---
name: doc-to-llm-docs
description: Turn a large Word document (.doc, .docx, .odt, .rtf, Word "web page" .mht saved as .doc) into a folder of LLM-friendly Markdown files — one self-contained file per section, with frontmatter, restored section numbers, Markdown tables, fenced code, an index README, llms.txt and llms-full.txt. Use this whenever the user wants a spec, methodology, API description, regulation or manual "converted for an LLM", "turned into documentation", "split into markdown", "made readable for Claude / RAG / an agent", or drops a .doc/.docx path and asks to extract or structure it — even if they only say "сделай из этого документацию" or "переведи в markdown". Do not use for PDF input (use marker/docling first, then this skill on the HTML) or for short notes where a single pandoc call is enough.
---

# doc-to-llm-docs

Converts one big Word document into a documentation folder that a model can navigate by index and load section by section. The heavy lifting is in `scripts/convert.py`; your job is to run it, read its report, spot-check the result, and fix the few things a script cannot judge (title, intro text, folder names).

## Why not just pandoc or textutil

- macOS `textutil` silently drops every table from `.doc` files and leaks Word field codes (`TOC`, `REF`, `NUMPAGES`) into the text. Never use it for this.
- Plain `pandoc` output of a 100+ page spec is one 8 000-line file with a 300-line table of contents, no section numbers (Word auto-numbering is not text), HTML tables for anything with multi-line cells, appendix headings rendered as numbered list items, and images referenced by absolute temp paths. A model given that file wastes most of its context on noise.

The script fixes all of that deterministically. `references/pitfalls.md` lists what it corrects and why; read it when a new document breaks in a new way, because most new failures are variations of a known one.

## Prerequisites

```bash
brew install pandoc                     # required
brew install --cask libreoffice         # only for .doc/.odt/.rtf (not needed for .docx or .mht)
python3 -m pip install beautifulsoup4   # required
```

## Workflow

1. **Run the converter** into a new folder next to the project or wherever the user keeps docs:

   ```bash
   python3 -I <skill-dir>/scripts/convert.py "<input.doc>" <output-dir> \
       --title "<Human title of the document>"
   ```

   `<skill-dir>` is the directory this SKILL.md lives in (the skill loader tells you its base directory; when installed as a plugin it is `${CLAUDE_PLUGIN_ROOT}/skills/doc-to-llm-docs`). Use that absolute path; do not copy the script into the user's project.

   Pass `--title` whenever you can see the document's name (file name, title page, user's words): the auto-guess reads the title page and often picks the system name instead of the document name.

   Options worth knowing:
   - `--max-lines 300` / `--split-level 3`: a section longer than max-lines is split into its subsections (down to split-level); shorter sections stay as one file. Raise max-lines for documents with long but coherent chapters. For API references where each method (level 3) has its own "поля запроса / ошибки / поля ответа / примеры" subsections and examples run to hundreds of lines, use `--split-level 4`: big methods become a folder with the examples in their own file, small methods stay as one file. Check the "files" list in the report: anything over ~400 lines that is not a single code example is a sign to go one level deeper.
   - `--bold-headings`: for documents without heading styles (the report shows `headings in body: 1–5` for a long document). Promotes short all-bold paragraphs to level-2 headings.
   - `--index-only` (`convert.py x <output-dir> --index-only`): rebuild README.md, llms.txt and llms-full.txt from the files already in the output folder, keeping the hand-written README intro. Run it after any manual edit, rename or deletion.
   - `--keep-intermediate`: keeps pandoc's raw markdown as `doc.md` in the output for debugging.

2. **Read the report.** The script prints what it found. The lines that matter:
   - `TOC entries / headings in body / numbered from TOC` — when the document has a table of contents, numbered-from-TOC should equal TOC entries. If the report lists unmatched TOC entries, those headings exist in the body as plain text that differs from the TOC (typo, extra word, footnote mark, hard line break). Find the text in the output and fix the heading by hand (`# <label> <title>` and the `sections:` line in frontmatter).
   - `unresolved links / broken file links` — should be 0. Unresolved anchors are cross-references to sections the splitter could not find; turn them into plain text or into a file link by hand.
   - `captioned tables` — if 0 but the document clearly has "Таблица N – …" captions, the caption style is new and the count is only cosmetic; note it in your report.
   - `files` list — the number of top-level files should match the number of top-level TOC entries. An extra file with a strange name (base64, a sentence, a table cell) is a false heading created by a Word bookmark; delete it, put its content back where it belongs, and leave the numbering gap or renumber and then run `--index-only`.

   If the script misbehaves on a pattern that looks general, fix the output by hand first and tell the user what the script got wrong; improving the script is a separate decision, not part of converting their document.

3. **Spot-check before declaring success**, because a report can be clean while content is wrong. Open one file of each kind and look at:
   - A parameter table: multi-paragraph cells joined with `<br>`, merged cells shown as `〃`, no leftover `<td>`/`<p>`, the header row is the real header (header-less tables get an empty header row, which is fine).
   - Every fenced block in a mid-sized file: it must contain code, not a nested list or prose that happened to be indented (`grep -n -A2 '^```' file.md`).
   - A diagram: `![…](../media/imageN.png)` resolves and the caption sits under it.
   - Footnotes: `grep -rn '^\[\^[0-9]*\]:\s*$'` must find nothing (empty definitions mean a multi-line footnote was lost) and headings must not carry `[^N]`.
   - `grep -rn '<span\|<img\|<td\|NUMPAGES\|PAGEREF\|HYPERLINK' <output-dir>` returns nothing.

4. **Finish the index by hand.** `README.md` has a generated file list and a `<!-- TODO -->` marker for the intro. Replace it with 3–6 lines a reader needs: what system the document describes, who it is for, where to start for each typical question (process → scenarios, request parameters → services, full examples → appendices, how to connect → connection section). This paragraph is what a model reads first; the generated list alone is not enough. For API references, list the entities and their endpoints in one line each: that is the lookup a model does most.

   File and folder names are transliterated section titles (`04-integratsionnyh-servisov/`). Rename them to something a developer would type (`03-services/`) when the output will live in a repo, then run `--index-only` and the link check below. Leave them if the output is a one-off dump.

5. **Verify links after any manual change:**

   ```bash
   python3 - <<'EOF'
   import re, pathlib
   for f in pathlib.Path("<output-dir>").rglob("*.md"):
       for m in re.finditer(r"\]\(([^)#]+\.(?:md|png))\)", f.read_text()):
           if not (f.parent / m.group(1)).exists(): print("BROKEN", f, m.group(1))
   EOF
   ```

6. **Tell the user** where the folder is, how many files/tables/code blocks/images it has, what you fixed by hand, what the script got wrong (so it can be improved later), and what the model still cannot read (diagrams are PNG; their textual steps usually sit in the table next to them). If two files in the user's folder turn out to be versions of the same document, say so and convert only the newest.

## Output contract

```
<output-dir>/
  README.md          index: every file with its title and one-line description, plus a hand-written intro
  llms.txt           same index in llmstxt.org format
  llms-full.txt      whole document concatenated with `<!-- file: … -->` markers (one-shot context loading)
  00-title-page.md   title page and version history, kept for reference (not part of the docs)
  media/             extracted images
  NN-slug.md         one file per section that fits --max-lines
  NN-slug/README.md  intro of a section that was split, followed by NN-slug/NN-sub.md files
```

Every `.md` starts with frontmatter: `title` (with the original section number), `source` (document title), `sections` (numbers covered), `description` (first paragraph, ≤200 chars). Headings inside keep their original numbers (`## 3.1.2 Метод «registerDocument»`) so references like "см. п. 3.1.2" stay greppable, and intra-document links are rewritten to relative file links.

## When the script is not enough

- **PDF input**: run `marker_single` or `docling` to get HTML/Markdown first, then point the script at the `.html` (it accepts HTML) or split the Markdown by hand following the same contract.
- **Scanned documents**: OCR first; this skill assumes real text.
- **Documents with a non-standard structure** (no headings, no TOC, everything in tables): the script will produce one big file. Decide with the user whether splitting by hand is worth it; often the single fenced/tabled file plus a README is already a big improvement.
- **Multiple documents for one system**: convert each into its own folder and write a top-level README that explains which document answers which question. Do not merge them into one tree; version drift between documents is real and the source should stay traceable.
