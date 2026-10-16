#!/usr/bin/env python3
"""
Convert a large Word document (.doc/.docx/.odt/.rtf) into a folder of LLM-friendly Markdown files.

Pipeline: LibreOffice (only for non-.docx) -> pandoc (GFM) -> post-processing -> split by headings.

Usage:
  convert.py INPUT OUTPUT_DIR [--title "Document title"] [--split-level 3] [--max-lines 300]
             [--keep-intermediate] [--no-index]

Output:
  OUTPUT_DIR/
    README.md        index with a one-line description of every file
    llms.txt         same index in llms.txt format
    llms-full.txt    whole document concatenated (for one-shot context loading)
    media/           extracted images
    NN-slug.md or NN-slug/NN-slug.md   one file per section (split until files fit --max-lines)

Requires: pandoc, beautifulsoup4; soffice (LibreOffice) for .doc/.odt/.rtf inputs.
"""
import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("beautifulsoup4 is required: python3 -m pip install beautifulsoup4")

# ----------------------------------------------------------------------------- helpers
TRANSLIT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh", "з": "z", "и": "i",
    "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
    "у": "u", "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "",
    "э": "e", "ю": "yu", "я": "ya",
}
STOPWORDS = {"описание", "метод", "метода", "сервиса", "веб-сервиса", "и", "в", "на", "по", "с", "для", "к", "о", "об", "из", "от"}


def slugify(title, max_words=6):
    s = title.lower()
    s = re.sub(r"[«»\"'“”„]", "", s)
    s = "".join(TRANSLIT.get(ch, ch) for ch in s)
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    words = [w for w in s.split("-") if w]
    # drop stopwords only when something remains
    kept = [w for w in words if w not in {"".join(TRANSLIT.get(c, c) for c in sw) for sw in STOPWORDS}]
    words = kept or words
    return "-".join(words[:max_words])[:50].strip("-") or "section"


def pandoc_slug(title):
    """Reproduce pandoc's auto identifier for a heading (gfm)."""
    s = title.lower()
    s = re.sub(r"[^\w\s-]", "", s)
    s = re.sub(r"\s+", "-", s.strip())
    s = re.sub(r"^[^a-zа-яё]+", "", s)
    return s


def norm(s):
    s = re.sub(r"\[\^\d+\]", "", s)
    s = re.sub(r"^[\d.\s]+", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()


ANCHOR = re.compile(r'<span id="([^"]*)" class="anchor"></span>')
HEADING_ANCHOR = re.compile(r"^(__RefHeading|_Toc)")
LABEL_RE = re.compile(r"^(Приложение\s+[А-ЯA-Z0-9]+|[А-ЯA-Z]\.\d+(?:\.\d+)*|\d+(?:\.\d+)*)\.?\s+(.+)$", re.I)
TOC_LINE = re.compile(r"^\[(.+)\]\(#[^)]+\)$")


def toc_title(line):
    """Return the entry title of a pandoc TOC line, or None. Handles both
    '[1 Title 12](#a)' (LibreOffice) and '[1 Title [12](#a)](#a)' (Word)."""
    m = TOC_LINE.match(line)
    if not m:
        return None
    inner = m.group(1)
    inner = re.sub(r"\s*\[\d+\]\(#[^)]+\)$", "", inner)
    inner = re.sub(r"\s+\d+$", "", inner)
    return inner.strip()
CAPTION_LIST = re.compile(r"^\s*\d+\.\s+[–-]\s+(.+)$")
CAPTION_WORD = re.compile(r"^\s*(?:\*\*)?(Таблица|Рисунок|Схема|Табл\.|Рис\.)\s+([\dА-Яа-яA-Z.]+?)\s*(?:\*\*)?\s*[–\-.:]\s*(?:\*\*)?\s*(.+?)(?:\*\*)?\.?$")


# ----------------------------------------------------------------------------- 1. to markdown
def mhtml_to_html(src: Path, workdir: Path):
    """Word 'web page, single file' (.mht) saved under a .doc name: extract the HTML part and images."""
    import base64
    import email
    from email import policy
    msg = email.message_from_bytes(src.read_bytes(), policy=policy.default)
    html_path = None
    media = workdir / "media" / "media"
    media.mkdir(parents=True, exist_ok=True)
    for part in msg.walk():
        ctype = part.get_content_type()
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        if ctype == "text/html" and html_path is None:
            charset = (part.get_content_charset() or "utf-8").lower()
            if charset == "unicode":
                charset = "utf-16"
            try:
                html = payload.decode(charset, errors="replace")
            except LookupError:
                html = payload.decode("utf-8", errors="replace")
            html_path = workdir / "doc.html"
            html_path.write_text(html, encoding="utf-8")
        elif ctype.startswith("image/"):
            loc = part.get("Content-Location", "")
            name = Path(loc).name or f"image{len(list(media.iterdir())) + 1}"
            (media / name).write_bytes(payload)
    if html_path is None:
        sys.exit("MHTML container has no text/html part")
    return html_path


def to_markdown(src: Path, workdir: Path):
    src = src.resolve()
    docx = src
    head = src.read_bytes()[:64]
    is_html = src.suffix.lower() in (".html", ".htm", ".mht", ".mhtml")
    if is_html or head.startswith(b"MIME-Version") or head.lstrip().lower().startswith((b"<html", b"<!doctype")):
        if not shutil.which("pandoc"):
            sys.exit("pandoc is required: brew install pandoc")
        html = mhtml_to_html(src, workdir) if head.startswith(b"MIME-Version") else src
        md = workdir / "doc.md"
        subprocess.run(["pandoc", str(html), "-f", "html-native_divs-native_spans", "--strip-comments",
                        "-t", "gfm", "--wrap=none", "-o", str(md)], check=True)
        return md.read_text(encoding="utf-8")
    if src.suffix.lower() != ".docx":
        if not shutil.which("soffice"):
            sys.exit("soffice (LibreOffice) is required for non-.docx input: brew install --cask libreoffice")
        subprocess.run(["soffice", "--headless", "--convert-to", "docx", "--outdir", str(workdir), str(src)],
                       check=True, capture_output=True)
        docx = workdir / (src.stem + ".docx")
        if not docx.exists():
            sys.exit("LibreOffice conversion failed")
    if not shutil.which("pandoc"):
        sys.exit("pandoc is required: brew install pandoc")
    md = workdir / "doc.md"
    subprocess.run(["pandoc", str(docx), "-t", "gfm", "--wrap=none",
                    f"--extract-media={workdir / 'media'}", "-o", str(md)], check=True)
    return md.read_text(encoding="utf-8")


# ----------------------------------------------------------------------------- 2. parse structure
def parse(text, bold_headings=False):
    lines = text.split("\n")

    # TOC block: a run of [title page](#anchor) lines
    toc, toc_end, run_start = [], None, None
    for idx, ln in enumerate(lines):
        if toc_title(ln) is not None:
            if run_start is None:
                run_start = idx
            toc_end = idx
        elif ln.strip() and run_start is not None and idx - toc_end > 2:
            break
    if toc_end is not None:
        for ln in lines[run_start:toc_end + 1]:
            full = toc_title(ln)
            if full:
                mm = LABEL_RE.match(full)
                toc.append((mm.group(1), mm.group(2).strip()) if mm else ("", full))
    if len(toc) < 3:
        toc, toc_end, run_start = [], None, None
    toc_titles = {norm(t) for _, t in toc}
    from collections import Counter
    _cnt = Counter(norm(t) for _, t in toc)
    unique_titles = {t for t, n in _cnt.items() if n == 1}  # safe to promote plain text/list items to headings

    # join "1.  \" + indented continuation (LibreOffice heading-as-list artefact) and
    # hard-broken paragraphs that together equal a TOC title (headings typed as plain text)
    joined, i = [], 0
    while i < len(lines):
        ln = lines[i]
        if re.match(r"^\s*\d+\.\s+(<span[^>]*></span>)?\\$", ln) and i + 1 < len(lines):
            joined.append(re.sub(r"\\$", "", ln) + lines[i + 1].strip())
            i += 2
            continue
        if ln.endswith("\\") and i + 1 < len(lines) and norm(ANCHOR.sub("", ln[:-1] + " " + lines[i + 1].strip())) in toc_titles:
            joined.append(ln[:-1].rstrip() + " " + lines[i + 1].strip())
            i += 2
            continue
        joined.append(ln)
        i += 1
    lines = joined

    body = lines[toc_end + 1:] if toc_end is not None else lines
    preamble = [l for l in (lines[:run_start] if run_start is not None else []) if l.strip()]

    items = []  # ("H", title, level_hint) | ("L", text)
    fig_n = tab_n = 0
    for ln in body:
        ids = ANCHOR.findall(ln)
        s = ANCHOR.sub("", ln)
        s = re.sub(r'<span class="[^"]*"></span>', "", s)  # empty styled spans (e.g. Word "mark")
        s = re.sub(r'<span class="[^"]*">(.*?)</span>', r"\1", s)  # smallcaps/underline spans: keep text
        s = re.sub(r"\bstart_\d+\b", "", s)  # Word bookmark leaks
        m = re.match(r"^(#{1,6}) (.+)$", s)
        if m:
            items.append(("H", m.group(2).strip(), len(m.group(1))))
            continue
        is_anchor_heading = (ids and any(HEADING_ANCHOR.match(x) for x in ids) and not CAPTION_LIST.match(s)
                             and not s.lstrip().startswith("|") and len(s.strip()) <= 200
                             and (not toc or norm(re.sub(r"^\s*(\d+\.\s+)?", "", s)) in toc_titles))
        list_item = re.match(r"^\s*\d+\.\s+(.+)$", s)
        plain = re.sub(r"^\*\*|\*\*$", "", s.strip())
        anchored_item = list_item and ids and norm(list_item.group(1)) in toc_titles
        if is_anchor_heading or anchored_item or (list_item and norm(list_item.group(1)) in unique_titles) or (plain and norm(plain) in unique_titles):
            s = plain if not list_item else s
            title = re.sub(r"^\s*(\d+\.\s+)?", "", s).strip()
            if title:
                items.append(("H", title, 0))  # level resolved later from TOC label
                continue
        m = CAPTION_LIST.match(s)
        if m:
            prev = next((o for o in reversed(items) if not (o[0] == "L" and not o[1].strip())), None)
            if prev and prev[0] == "L" and prev[1].startswith("<img"):
                fig_n += 1
                items[items.index(prev)] = ("L", f"IMGCAP::{fig_n}::{m.group(1).strip()}")
            else:
                tab_n += 1
                items.append(("L", f"**Таблица {tab_n} – {m.group(1).strip()}**"))
            continue
        m = CAPTION_WORD.match(s)
        if m:
            items.append(("L", f"**{m.group(1)} {m.group(2)} – {m.group(3).strip()}**"))
            continue
        m = re.match(r"^\*\*([^*]{3,90})\*\*\s*$", s.strip())
        if bold_headings and m and not re.match(r"^(Таблица|Рисунок|Табл\.|Рис\.|Примечани|Запрос|Ответ)", m.group(1)):
            items.append(("H", m.group(1).strip().rstrip(":"), 2))
            continue
        if s.strip() == "<!-- -->":
            continue
        items.append(("L", s))

    # drop setext underlines (=== / ---) left behind when a plain-text line was promoted to a heading
    cleaned = []
    for it in items:
        if it[0] == "L" and re.match(r"^\s*(={3,}|-{3,})\s*$", it[1]):
            prev = next((o for o in reversed(cleaned) if not (o[0] == "L" and not o[1].strip())), None)
            if prev and prev[0] == "H":
                continue
        cleaned.append(it)
    items = cleaned

    # images
    out = []
    for it in items:
        if it[0] == "L" and it[1].startswith("IMGCAP::"):
            _, n, cap = it[1].split("::", 2)
            out.append(("L", f"![{cap}](media/image{n}.png)"))
            out.append(("L", f"*Рисунок {n} – {cap}*"))
        elif it[0] == "L" and "<img" in it[1]:
            def img(m):
                nonlocal fig_n
                fig_n += 1
                src = re.search(r'src="([^"]+)"', m.group(0))
                return f"![Рисунок {fig_n}](media/{Path(src.group(1)).name if src else ''})"
            out.append(("L", re.sub(r"<img[^>]*>", img, it[1])))
        else:
            out.append(it)
    items = out

    # assign labels from TOC (in-order greedy matching with a lookahead window)
    heads = [i for i, it in enumerate(items) if it[0] == "H"]
    labelled = {}
    ti = 0
    from collections import Counter
    title_count = Counter(norm(t) for _, t in toc)
    for hi in heads:
        title = norm(items[hi][1])
        window = len(toc) if title_count.get(title, 0) == 1 else ti + 6
        for look in range(ti, min(window, len(toc))):
            if norm(toc[look][1]) == title:
                labelled[hi] = toc[look][0]
                ti = look + 1
                break
    matched = len(labelled)
    used = set(toc[i][0] + " " + toc[i][1] for i in range(len(toc)) if any(
        labelled.get(h) == toc[i][0] and norm(items[h][1]) == norm(toc[i][1]) for h in heads))
    unmatched_toc = [l + " " + t for l, t in toc if (l + " " + t) not in used]

    final = []
    for i, it in enumerate(items):
        if it[0] == "H":
            lbl = labelled.get(i, "")
            lvl = it[2]
            if lbl:
                lvl = 1 if lbl.lower().startswith("приложение") else lbl.count(".") + 1
            if lvl == 0:
                lvl = 1
            final.append(("H", lbl, it[1], lvl))
        else:
            final.append(it)
    return final, toc, matched, len(heads), fig_n, tab_n, preamble, unmatched_toc


# ----------------------------------------------------------------------------- 3. tables
def cell_text(cell):
    for a in cell.select("a.footnote-ref"):
        a.replace_with(f" (сноска {a.get_text(strip=True)})")
    for s in cell.find_all("strong"):
        s.replace_with(f"**{s.get_text(' ', strip=True)}**")
    for c in cell.find_all("code"):
        c.replace_with(f"`{c.get_text(strip=True)}`")
    parts = []
    blocks = [b for b in cell.find_all(["p", "li"]) if not (b.name == "p" and b.parent.name == "li")]
    if blocks:
        for b in blocks:
            t = re.sub(r"\bstart_\d+\b", "", b.get_text(" ", strip=True))
            if not t:
                continue
            if b.name == "li" and b.parent.name == "ol":
                n = b.parent.find_all("li", recursive=False).index(b) + 1
                parts.append(f"{n}. {t}")
            else:
                parts.append(("- " + t) if b.name == "li" else t)
    else:
        t = cell.get_text(" ", strip=True)
        if t:
            parts.append(t)
    return [re.sub(r"\s+", " ", p).strip() for p in parts if p.strip()]


def join_cell(parts):
    return "<br>".join(parts).replace("|", "\\|")


def strip_bold(parts):
    if not parts:
        return parts
    parts = list(parts)
    parts[0] = re.sub(r"^\*\*", "", parts[0]).lstrip()
    parts[-1] = re.sub(r"\*\*$", "", parts[-1]).rstrip()
    return [p for p in parts if p.strip()]


def code_like(parts):
    first = parts[0].lstrip("*").lstrip() if parts else ""
    return (len(parts) >= 3 and first[:1] in "{[<") or (len(parts) >= 2 and re.match(r"^(GET|POST|PUT|DELETE|PATCH)\s", first) is not None) \
        or (len(parts) >= 2 and re.match(r"^ey[A-Za-z0-9_-]{20,}", first) is not None)


def html_table_to_md(html):
    soup = BeautifulSoup(html, "html.parser")
    fns = []
    for sec in soup.select("section.footnotes"):
        for li in sec.find_all("li"):
            t = re.sub(r"\s+", " ", li.get_text(" ", strip=True)).replace("↩︎", "").strip()
            fns.append(f"Сноска {li.get('id', '').replace('fn', '')}: {t}")
        sec.decompose()
    cap = soup.find("caption")
    captxt = None
    if cap:
        captxt = re.sub(r"\s+", " ", cap.get_text(" ", strip=True))
        cap.decompose()
    has_header = soup.find("th") is not None
    grid = {}  # (row, col) -> text
    raw = {}   # (row, col) -> list of paragraphs (for example tables)
    for r, tr in enumerate(soup.find_all("tr")):
        c = 0
        for td in tr.find_all(["td", "th"]):
            while (r, c) in grid:
                c += 1
            parts = cell_text(td)
            t = join_cell(parts)
            if td.name == "th" and t and not t.startswith("**"):
                t = f"**{t}**"
            raw[(r, c)] = parts
            try:
                rs, cs = int(td.get("rowspan", 1) or 1), int(td.get("colspan", 1) or 1)
            except ValueError:
                rs, cs = 1, 1
            for dr in range(rs):
                for dc in range(cs):
                    grid[(r + dr, c + dc)] = t if (dr == 0 and dc == 0) else ("〃" if dr > 0 and dc == 0 else "")
            c += cs
    if not grid:
        return fns, []
    nrows = max(k[0] for k in grid) + 1
    width = max(k[1] for k in grid) + 1
    rows = [[grid.get((r, c), "") for c in range(width)] for r in range(nrows)]

    # request/response example tables: two columns, right cell is multi-line code -> label + fenced block
    def opaque(parts):  # a single long token/base64/URL-ish string without spaces
        return len(parts) == 1 and " " not in parts[0] and len(parts[0]) > 40
    if width == 1 and all(code_like(strip_bold(raw.get((r, 0), []))) or opaque(strip_bold(raw.get((r, 0), []))) for r in range(nrows)):
        md = [f"**{captxt}**", ""] if captxt else []
        for r in range(nrows):
            parts = strip_bold(raw.get((r, 0), []))
            md += [f"```{guess_lang(parts[0]) or 'text'}"] + parts + ["```", ""]
        return fns, md
    if width == 2 and any(code_like(raw.get((r, 1), [])) for r in range(nrows)):
        md = [f"**{captxt}**", ""] if captxt else []
        for r in range(nrows):
            label = re.sub(r"^\*\*|\*\*$", "", " ".join(raw.get((r, 0), [])).strip())
            parts = strip_bold(raw.get((r, 1), []))
            if code_like(parts):
                lang = guess_lang(parts[0])
                md += [f"**{label}**", "", f"```{lang}"] + parts + ["```", ""]
            else:
                md += [f"**{label}:** {' '.join(parts)}", ""]
        return fns, md
    md = []
    if captxt:
        md += [f"**{captxt}**", ""]
    if has_header:
        md.append("| " + " | ".join(rows[0]) + " |")
        rows = rows[1:]
    else:
        md.append("|" + "   |" * width)
    md.append("|" + "---|" * width)
    for r in rows:
        md.append("| " + " | ".join(r) + " |")
    return fns, md


def convert_tables(items):
    out, i = [], 0
    while i < len(items):
        it = items[i]
        if it[0] == "L" and it[1].lstrip().startswith("<table"):
            buf = []
            while i < len(items):
                buf.append(items[i][1])
                if items[i][0] == "L" and items[i][1].strip() == "</table>":
                    i += 1
                    break
                i += 1
            j = i
            while j < len(items) and items[j][0] == "L" and not items[j][1].strip():
                j += 1
            extra = []
            if j < len(items) and items[j][0] == "L" and items[j][1].startswith("<section"):
                i = j
                while i < len(items):
                    extra.append(items[i][1])
                    if items[i][1].strip() == "</section>":
                        i += 1
                        break
                    i += 1
            fns, md = html_table_to_md("\n".join(buf + extra))
            out += [("L", l) for l in md]
            for fn in fns:
                out += [("L", ""), ("L", fn)]
            continue
        out.append(it)
        i += 1
    return out


# ----------------------------------------------------------------------------- 4. code blocks
def guess_lang(first):
    s = first.strip()
    if s.startswith("<"):
        return "xml"
    if s.startswith(("{", "[")):
        return "json"
    if re.match(r"^(openapi|swagger|info|paths|components|servers):", s):
        return "yaml"
    if re.match(r"^(GET|POST|PUT|DELETE|PATCH) ", s) or s.startswith("HTTP/"):
        return "http"
    return ""


def fence_json_paragraphs(items):
    """Word-HTML documents give one <p> per line: a lone '{' or '[' line starts a JSON block that ends when brackets balance."""
    out, i, in_fence = [], 0, False
    while i < len(items):
        it = items[i]
        if it[0] == "L" and it[1].startswith("```"):
            in_fence = not in_fence
        if not in_fence and it[0] == "L" and it[1].strip() in ("{", "["):
            depth, block, j = 0, [], i
            while j < len(items) and items[j][0] == "L":
                t = items[j][1]
                if t.strip():
                    block.append(t.rstrip())
                    depth += t.count("{") + t.count("[") - t.count("}") - t.count("]")
                    if depth <= 0:
                        break
                j += 1
            if depth <= 0 and len(block) > 1:
                out.append(("L", "```json"))
                out += [("L", b[4:] if b.startswith("    ") else b) for b in block]
                out.append(("L", "```"))
                i = j + 1
                continue
        out.append(it)
        i += 1
    return out


def join_footnotes(items):
    """pandoc writes long footnotes as '[^5]:' followed by indented/quoted lines; fold them into one line."""
    out, i = [], 0
    while i < len(items):
        it = items[i]
        m = it[0] == "L" and re.match(r"^\[\^(\d+)\]:\s*(.*)$", it[1])
        if m:
            text = [m.group(2).strip()] if m.group(2).strip() else []
            j = i + 1
            while j < len(items) and items[j][0] == "L" and (items[j][1].startswith("    ") or
                                                            (not items[j][1].strip() and j + 1 < len(items) and items[j + 1][1].startswith("    "))):
                t = items[j][1].strip()
                if t:
                    text.append(re.sub(r"^>\s?", "", t))
                j += 1
            out.append(("L", f"[^{m.group(1)}]: " + " ".join(text)))
            i = j
            continue
        out.append(it)
        i += 1
    return out


def fence_pipe_single_column(items):
    """pandoc's own pipe tables with one column holding a token/JSON example -> fenced block."""
    out, i = [], 0
    while i < len(items):
        if items[i][0] == "L" and items[i][1].startswith("|"):
            j = i
            while j < len(items) and items[j][0] == "L" and items[j][1].startswith("|"):
                j += 1
            rows = [items[k][1] for k in range(i, j)]
            if len(rows) >= 2 and re.match(r"^\|:?-+:?\|$", rows[1].strip()):
                cells = [r.strip().strip("|").strip() for r in rows if not re.match(r"^\|:?-+:?\|$", r.strip())]
                cells = [c for c in cells if c]
                if cells and all((" " not in c and len(c) > 40) or c[:1] in "{[<" for c in cells):
                    for c in cells:
                        out += [("L", f"```{guess_lang(c) or 'text'}"), ("L", c), ("L", "```"), ("L", "")]
                    i = j
                    continue
            out += items[i:j]
            i = j
            continue
        out.append(items[i])
        i += 1
    return out


def fence_code(items):
    out, i, in_fence = [], 0, False
    while i < len(items):
        it = items[i]
        if it[0] == "L" and it[1].startswith("```"):
            in_fence = not in_fence
        prev = next((o[1] for o in reversed(out) if o[0] == "L" and o[1].strip()), "")
        prev_is_list = bool(re.match(r"^\s*(\d+[.)]|[-*•])\s", prev)) or prev.endswith(":") and False
        starts_like_list = bool(re.match(r"^\s*(\d+[.)]|[-*•])\s", it[1]))
        if not in_fence and it[0] == "L" and it[1].startswith("    ") and it[1].strip() and not prev_is_list and not starts_like_list:
            block = []
            while i < len(items) and items[i][0] == "L" and (items[i][1].startswith("    ") or not items[i][1].strip()):
                block.append(items[i][1])
                i += 1
            while block and not block[-1].strip():
                block.pop()
            out.append(("L", f"```{guess_lang(block[0])}"))
            out += [("L", b[4:] if b.startswith("    ") else b) for b in block]
            out.append(("L", "```"))
            continue
        out.append(it)
        i += 1
    return out


# ----------------------------------------------------------------------------- 5. split
class Node:
    def __init__(self, lbl, title, level):
        self.lbl, self.title, self.level = lbl, title, level
        self.lines, self.children = [], []

    def size(self):
        return len(self.lines) + sum(c.size() for c in self.children)


def build_tree(items):
    root = Node("", "", 0)
    stack = [root]
    for it in items:
        if it[0] == "H":
            _, lbl, title, lvl = it
            node = Node(lbl, title, lvl)
            while stack and stack[-1].level >= lvl:
                stack.pop()
            stack[-1].children.append(node)
            stack.append(node)
        else:
            stack[-1].lines.append(it[1])
    return root


def plan_files(root, split_level, max_lines):
    """Return list of (path, node, include_children: bool)."""
    files = []

    def visit(node, dirpath, seq):
        name = f"{seq:02d}-{slugify(node.title)}"
        big = node.size() > max_lines and node.children and node.level < split_level
        if not big:
            files.append((dirpath / f"{name}.md", node, True))
            return
        sub = dirpath / name
        files.append((sub / "README.md", node, False))
        for k, ch in enumerate(node.children, 1):
            visit(ch, sub, k)

    for k, ch in enumerate(root.children, 1):
        visit(ch, Path("."), k)
    return files


def render(node, depth, include_children):
    prefix = f"{node.lbl} " if node.lbl else ""
    lines = [f"{'#' * min(depth, 6)} {prefix}{node.title}"] + node.lines
    if include_children:
        for ch in node.children:
            lines += render(ch, depth + 1, True)
    return lines


# ----------------------------------------------------------------------------- 6. write
LINK_RE = re.compile(r"\[([^\]]+)\]\(#([^)]+)\)")


def tidy(body):
    s = "\n".join(body)
    s = re.sub(r"\n{3,}", "\n\n", s)
    s = s.replace("\\\n", "  \n")
    s = re.sub(r"\\([-_<>\[\]+#~().])", r"\1", s)
    return s.strip() + "\n"


def natural_key(p):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", p.as_posix())]


def read_frontmatter(text):
    m = re.match(r"^---\n(.*?)\n---\n", text, flags=re.S)
    meta = {}
    if m:
        for l in m.group(1).split("\n"):
            k, _, v = l.partition(":")
            meta[k.strip()] = v.strip().strip('"')
    return meta


def rebuild_index(out: Path, title=None):
    files = sorted((p for p in out.rglob("*.md") if p.name not in ("README.md", "00-title-page.md") or p.parent != out),
                   key=lambda p: natural_key(p.relative_to(out)))
    # keep sub-folder README.md files (section intros) but not the top-level one
    written = []
    for p in files:
        text = p.read_text(encoding="utf-8")
        fm = read_frontmatter(text)
        t = fm.get("title") or re.sub(r"^#\s*", "", next((l for l in text.split("\n") if l.startswith("#")), p.stem))
        written.append((p.relative_to(out), t, fm.get("description", "")))
    readme_path = out / "README.md"
    intro = None
    if readme_path.exists():
        old = readme_path.read_text(encoding="utf-8")
        mm = re.match(r"^(# .*?)\n## Содержание\n", old, flags=re.S)
        if mm:
            intro = mm.group(1).rstrip("\n")
    if intro is None:
        intro = f"# {title or out.name}\n\nДокументация в формате Markdown для чтения людьми и языковыми моделями."
    doc_title = re.sub(r"^# ", "", intro.split("\n")[0])

    def entry(path, t, d, sep):
        return "  " * len(path.parent.parts) + f"- [{t}]({path.as_posix()}){sep}{d}"
    readme_path.write_text(intro + "\n\n## Содержание\n\n" + "\n".join(entry(p, t, d, " — ") for p, t, d in written) + "\n", encoding="utf-8")
    (out / "llms.txt").write_text(f"# {doc_title}\n\n> Markdown-версия документа. Файлы ниже самодостаточны; описание каждого — одна строка.\n\n## Файлы\n\n"
                                  + "\n".join(entry(p, t, d, ": ").lstrip() for p, t, d in written) + "\n", encoding="utf-8")
    full = []
    for path, _, _ in written:
        t = (out / path).read_text(encoding="utf-8")
        t = re.sub(r"^---\n.*?\n---\n", "", t, count=1, flags=re.S)
        full.append(f"<!-- file: {path.as_posix()} -->\n" + t)
    (out / "llms-full.txt").write_text("\n\n".join(full), encoding="utf-8")
    print(f"index rebuilt: {len(written)} files")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--title", help="document title for frontmatter/README (default: from the title page or file name)")
    ap.add_argument("--split-level", type=int, default=3, help="deepest heading level that may start a new file (default 3)")
    ap.add_argument("--max-lines", type=int, default=300, help="split a section into sub-files when it is longer than this (default 300)")
    ap.add_argument("--bold-headings", action="store_true",
                    help="treat short paragraphs that are entirely bold as level-2 headings (for documents without heading styles)")
    ap.add_argument("--keep-intermediate", action="store_true", help="keep the pandoc markdown next to the output (doc.md)")
    ap.add_argument("--no-index", action="store_true", help="do not write README.md / llms.txt / llms-full.txt")
    ap.add_argument("--index-only", action="store_true",
                    help="INPUT is ignored; rebuild README.md, llms.txt and llms-full.txt from the .md files already in OUTPUT (keeps the README intro)")
    args = ap.parse_args()
    if args.index_only:
        return rebuild_index(Path(args.output), args.title)

    src, out = Path(args.input), Path(args.output)
    workdir = Path(tempfile.mkdtemp(prefix="doc2md-"))
    text = to_markdown(src, workdir)

    items, toc, matched, nheads, fig_n, tab_n, preamble, unmatched_toc = parse(text, args.bold_headings)
    items = convert_tables(items)
    items = join_footnotes(items)
    items = fence_code(items)
    items = fence_json_paragraphs(items)
    items = fence_pipe_single_column(items)

    # preamble (title page etc.) before the first heading
    pre = []
    for it in items:
        if it[0] == "H":
            break
        pre.append(it[1])
    pre_text = preamble + [l for l in pre if l.strip()]
    def guess_title():
        cands = [re.sub(r"[*_#\\]", "", l).strip() for l in pre_text if len(l.strip()) > 10]
        kw = re.compile(r"рекомендаци|описание|руководство|порядок|инструкци|регламент|спецификац|требован|guide|specification", re.I)
        first_head = next((it[2] for it in items if it[0] == "H"), None)
        t = next((c for c in cands if kw.search(c)), cands[0] if cands else (first_head or src.stem))
        return t.capitalize() if t.isupper() else t
    doc_title = args.title or guess_title()
    items = items[len(pre):]

    root = build_tree(items)
    plan = plan_files(root, args.split_level, args.max_lines)

    # slug -> file map for intra-document links (pandoc slugs follow '#'-heading order)
    counts, slug_to_file, label_to_file = {}, {}, {}
    for path, node, inc in plan:
        def reg(n):
            base = pandoc_slug(n.title)
            k = counts.get(base, 0)
            counts[base] = k + 1
            slug_to_file[base if k == 0 else f"{base}-{k}"] = path
            if n.lbl:
                label_to_file[n.lbl] = path
            if inc:
                for c in n.children:
                    reg(c)
        reg(node)

    out.mkdir(parents=True, exist_ok=True)
    media_src = workdir / "media" / "media"
    if media_src.exists():
        (out / "media").mkdir(exist_ok=True)
        for p in media_src.iterdir():
            shutil.copy(p, out / "media" / p.name)

    unresolved = set()
    written = []  # (path, title, description)

    def fix_links(line, this):
        def rep(m):
            txt, slug = m.group(1), m.group(2)
            if re.match(r"^(Таблица|Рисунок|Табл\.|Рис\.)\s", txt):
                return txt
            target = slug_to_file.get(slug)
            if target is None:
                mm = LABEL_RE.match(txt + " x")
                if mm:
                    lbl = mm.group(1)
                    while lbl and lbl not in label_to_file:
                        lbl = lbl.rsplit(".", 1)[0] if "." in lbl else ""
                    target = label_to_file.get(lbl)
            if target is None:
                unresolved.add(slug)
                return txt
            if target == this:
                return txt
            rel = Path(*([".."] * len(this.parent.parts))) / target if this.parent.parts else target
            return f"[{txt}]({rel.as_posix()})"
        return LINK_RE.sub(rep, line)

    for path, node, inc in plan:
        body = render(node, 1, inc)
        body = [fix_links(l, path) for l in body]
        depth = len(path.parent.parts)
        body = [l.replace("](media/", "](" + "../" * depth + "media/") for l in body]
        first = next((l for l in body[1:] if l.strip() and not l.startswith(("#", "|", "!", "```", "**Таблица", "*Рисунок"))), "")
        first = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", first)
        desc = re.sub(r"\s+", " ", re.sub(r"[*_`]", "", first)).strip()
        desc = re.sub(r"\\([-_<>\[\]+#~])", r"\1", desc)
        if len(desc) > 200:
            desc = desc[:200].rsplit(" ", 1)[0] + "…"
        title = (node.lbl + " " if node.lbl else "") + node.title
        secs = ", ".join(l for l in ([node.lbl] + ([c.lbl for c in node.children] if inc else [])) if l)
        fm = ["---", f"title: \"{title.replace(chr(34), chr(39))}\"", f"source: \"{doc_title.replace(chr(34), chr(39))}\"",
              f"sections: \"{secs}\"", f"description: \"{desc.replace(chr(34), chr(39))}\"", "---", ""]
        p = out / path
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("\n".join(fm) + tidy(body), encoding="utf-8")
        written.append((path, title, desc))

    # move pandoc footnote definitions next to their references
    defs = {}
    texts = {}
    for path, _, _ in written:
        t = (out / path).read_text(encoding="utf-8")
        keep = []
        for l in t.split("\n"):
            m = re.match(r"^\[\^(\d+)\]: (.*)$", l)
            if m:
                defs[m.group(1)] = m.group(2)
            else:
                keep.append(l)
        texts[path] = "\n".join(keep)
    for path, t in texts.items():
        used = sorted({n for n in re.findall(r"\[\^(\d+)\](?!:)", t)}, key=int)
        for n in used:
            if n in defs:
                t = t.rstrip("\n") + f"\n\n[^{n}]: {defs[n]}\n"
        (out / path).write_text(t.rstrip("\n") + "\n", encoding="utf-8")

    if not args.no_index:
        def entry(path, title, desc, sep):
            depth = len(path.parent.parts)
            return "  " * depth + f"- [{title}]({path.as_posix()}){sep}{desc}"
        idx = [entry(p, t, d, " — ") for p, t, d in written]
        readme = [f"# {doc_title}", "",
                  f"Документация в формате Markdown, подготовленная из файла «{src.name}» для чтения людьми и языковыми моделями. "
                  "Каждый файл самодостаточен и начинается с frontmatter (title, sections, description). "
                  "Диаграммы лежат в `media/` как изображения.", "",
                  "<!-- TODO: при необходимости дополните вводную: для кого документ, как по нему ориентироваться -->", "",
                  "## Содержание", ""] + idx + [""]
        (out / "README.md").write_text("\n".join(readme), encoding="utf-8")
        llms = [f"# {doc_title}", "", f"> Markdown-версия документа «{src.name}». Файлы ниже самодостаточны; описание каждого — одна строка.", "",
                "## Файлы", ""] + [entry(p, t, d, ": ").lstrip() for p, t, d in written] + [""]
        (out / "llms.txt").write_text("\n".join(llms), encoding="utf-8")
        full = []
        for path, _, _ in written:
            t = (out / path).read_text(encoding="utf-8")
            t = re.sub(r"^---\n.*?\n---\n", "", t, count=1, flags=re.S)
            full.append(f"<!-- file: {path.as_posix()} -->\n" + t)
        (out / "llms-full.txt").write_text("\n\n".join(full), encoding="utf-8")

    if args.keep_intermediate:
        shutil.copy(workdir / "doc.md", out / "doc.md")
    if pre_text:
        raw = "\n".join(preamble + pre)
        def table_text(m):
            soup = BeautifulSoup(m.group(0), "html.parser")
            leaves = [c for c in soup.find_all(["td", "th"]) if not c.find("table")]
            return "\n\n".join(t for t in (re.sub(r"\s+", " ", c.get_text(" ", strip=True)) for c in leaves) if t)
        raw = re.sub(r"<table.*</table>", table_text, raw, flags=re.S)
        raw = re.sub(r"</?span[^>]*>", "", raw)
        raw = "\n".join(l for l in raw.split("\n") if not re.match(r"^#+\s*(содержание|оглавление)\s*$", l.strip(), re.I))
        (out / "00-title-page.md").write_text(tidy(raw.split("\n")), encoding="utf-8")

    # link check
    broken = []
    for p in out.rglob("*.md"):
        if p.name == "00-title-page.md":
            continue
        for m in re.finditer(r"\]\(([^)#]+\.(?:md|png|jpg|jpeg|gif|emf|wmf))\)", p.read_text(encoding="utf-8")):
            if not (p.parent / m.group(1)).resolve().exists():
                broken.append((p, m.group(1)))

    # report
    print(f"source:            {src.name}")
    print(f"title:             {doc_title}")
    print(f"TOC entries:       {len(toc)}   headings in body: {nheads}   numbered from TOC: {matched}")
    print(f"figures: {fig_n}   captioned tables: {tab_n}   files written: {len(written)}")
    print(f"unresolved links:  {len(unresolved)}   broken file links: {len(broken)}")
    for b in broken[:10]:
        print("   BROKEN", b)
    if unresolved:
        print("   unresolved anchors:", ", ".join(sorted(unresolved)[:10]))
    if toc and unmatched_toc:
        print(f"WARNING: {len(unmatched_toc)} TOC entries have no matching heading in the body (their text may differ or the style is not a heading):")
        for t in unmatched_toc[:15]:
            print("   -", t)
    if toc and matched < nheads:
        print(f"note: {nheads - matched} headings are deeper than the TOC and were left unnumbered (normal)")
    if not toc:
        print("NOTE: no table of contents found; section numbers not restored, levels taken from heading styles")
    print("files:")
    for path, title, _ in written:
        print(f"  {sum(1 for _ in open(out / path, encoding='utf-8')):5d}  {path.as_posix()}  —  {title[:70]}")
    print(f"\nintermediate files in {workdir}")


if __name__ == "__main__":
    main()
