# Pitfalls of Word → Markdown conversion (and what convert.py does about them)

Keep this list in mind when a new document breaks: most new failures are a variation of one of these.

## Input formats

| Symptom | Cause | Handling |
|---|---|---|
| `.doc` converts but has **no tables** and lines like `На NUMPAGES 140 IF =MOD(…)` | macOS `textutil` was used | Never use textutil. Script uses LibreOffice (`soffice --headless --convert-to docx`) then pandoc. |
| `.doc` starts with `MIME-Version: 1.0` | It is a Word "web page, single file" (MHTML) renamed to .doc. LibreOffice opens it as text. | Script detects the header, extracts the `text/html` part with Python `email`, decodes `charset="unicode"` as UTF-16, saves image parts to `media/`, runs pandoc with `-f html-native_divs-native_spans --strip-comments` so `mso-*` spans do not leak. |
| Word HTML / MHTML headings are `<h1>` only; sections are bold paragraphs | Author used bold instead of heading styles | `--bold-headings` promotes short all-bold paragraphs (3–90 chars, not starting with Таблица/Рисунок/Запрос/Ответ) to H2. |
| JSON examples come out as paragraphs with blank lines between | Word HTML emits one `<p>` per line | `fence_json_paragraphs`: a lone `{` or `[` line starts a block that ends when brackets balance. XML in this form is not handled yet. |
| pandoc fails on `.docx` from LibreOffice with footnotes in tables | Footnotes end up as `<section class="footnotes">` after the table in HTML mode | Script absorbs the section into the table conversion and emits `Сноска N: …` lines. |

## Structure

| Symptom | Cause | Handling |
|---|---|---|
| Headings have **no numbers** (`# Общие сведения` instead of `# 1 Общие сведения`) | Word auto-numbering is not text; pandoc drops it | Script parses the TOC (`[1 Title 12](#a)` from LibreOffice, `[1 Title [12](#a)](#a)` from Word) and matches body headings to TOC entries in order. Labels like `2.1.3`, `Приложение Б`, `Б.1` give the level. |
| Appendix headings appear as `1.  <span id="__RefHeading…"></span>Title` or `1.  \` + indented line | LibreOffice exports "Appendix" styles as numbered list items | Lines carrying a `__RefHeading`/`_Toc` anchor, or list items whose text equals a TOC title, become headings. `\`-continued list lines are joined first. |
| A file named after a JWT token or a sentence appears; a table row went missing from its section | Word left a `_Toc` bookmark inside a table cell or a paragraph; the TOC link points there | Anchor-based promotion is refused for table rows (`\|…`), lines over 200 chars, and (when a TOC exists) text that is not a TOC title. |
| `ПРИЛОЖЕНИЕ 1.` not recognised as an appendix; heading with `[^3]` never matches its TOC entry | Upper-case label; footnote mark inside the heading text | `LABEL_RE` is case-insensitive; `norm()` drops `[^N]` and leading digits before comparing. |
| A TOC entry stays unmatched although the text exists | Heading typed as plain text with a hard line break (`Title part 1\` newline `part 2`), or a stray digit glued to the text (`8Метод …`) | Script joins hard-broken lines when the join equals a TOC title, and `norm()` ignores leading digits. Only titles that are **unique** in the TOC are promoted from plain text, otherwise "Участники" would become a heading every time it appears in a sentence. |
| Repeated short titles (Участники, Предварительные условия) matched to the wrong TOC entry | Greedy matching with unlimited lookahead | Lookahead window is 6 entries for non-unique titles, unlimited for unique ones. |
| 300 lines of `[2.1.1.1 Участники 11](#участники)` at the top | Pandoc keeps the TOC as links | The TOC block is detected and removed; the title page before it is saved as `00-title-page.txt`. |

## Tables, captions, figures

| Symptom | Cause | Handling |
|---|---|---|
| Tables come out as raw `<table>` HTML | Any cell with two paragraphs forces pandoc to HTML | BeautifulSoup rewrite to pipe tables: paragraphs joined with `<br>`, `<li>` → `- item` / `N. item`, `|` escaped, `<th>` bolded. |
| Rows shift left after a merged cell | `rowspan` | Grid placement; continuation cells show `〃`. |
| First data row becomes the header | Table has no `<th>` | Empty header row is emitted (GFM requires one). Looks odd for key-value tables but preserves data. |
| JSON/HTTP examples squeezed into a 2-column table (`Пример запроса` / `Пример ответа`) with `<br>` between lines | Authors put examples in tables | A 2-column table whose right cells are multi-line code (start with `{`, `[`, `<` or an HTTP verb) is rendered as `**label**` + fenced block per row instead of a pipe table. |
| Captions appear as `18. <span id="table_x"></span> – Общие сведения о методе` | Word caption numbering is a SEQ field; LibreOffice renders captions as list items | Pattern `N. – Title` → `**Таблица N – Title**` (counted sequentially) or, if the previous block is an image, `![Title](media/imageN.png)` + `*Рисунок N – Title*`. Explicit `Таблица N – …` lines are kept as bold. Sequential counting assumes every table has a caption; verify against in-text references ("см. Таблица 18"). |
| `<img src="/var/folders/…/media/media/image1.png" style="…">` | pandoc `--extract-media` with absolute path | Every `<img>` (also inline at the end of a paragraph) is rewritten to `![Рисунок N](media/<basename>)`, relative to the file depth. |
| `**Схема 1**. Title` / `Таблица 4. Title` not recognised as captions | Caption uses a dot or colon, bold ends before the separator, or the word is «Схема» | `CAPTION_WORD` accepts `–`, `-`, `.`, `:` and bold around the number. The sequential table counter in the report only counts `N. – Title` captions, so it may stay 0 for such documents; cosmetic. |
| One-cell table (or pandoc pipe table) holding a token or JSON example | Authors box examples in a table | Single-column tables whose cells are code-like or one long space-free string become fenced blocks (`text`/`json`). |
| Bookmark names leak into text: `…об ЭМД.start_1` | Word bookmarks inside runs | `start_\d+` stripped everywhere, including table cells. |

## Text

| Symptom | Cause | Handling |
|---|---|---|
| `\-`, `\_`, `\[base\]`, `\<Параметры\>` | pandoc escapes Markdown-significant characters | Unescaped for `- _ < > [ ] + # ~` after splitting (safe because tables are already rebuilt and `\|` is left alone). |
| Lines ending with `\` | Hard line breaks | Converted to two-space breaks. |
| Lines containing only a non-breaking space | Word empty paragraphs | Treated as blank. |
| `[Таблица 18](#table_registerDocument_1)` | Cross-reference links to captions | Reduced to plain `Таблица 18`. |
| `[2.1.2 Поиск записей](#поиск-записей-в-рэмд)` | Cross-reference to a section | Pandoc slugs are recomputed per heading (lowercase, strip punctuation, spaces→hyphens, `-1`/`-2` for duplicates) and mapped to the file that now holds the heading; links become relative paths. Anchors like `#AppB` are resolved by the link text (`Приложение Б`). |
| Footnote definitions all at the end of the last file | pandoc puts `[^1]: …` at document end | Definitions are moved into the file that references them. |
| Empty `[^5]:` lines and a fenced block of quoted prose nearby | pandoc writes long footnotes as `[^5]:` + indented `> text` lines, which the code fencer then swallows | `join_footnotes` folds the continuation lines into one definition before code fencing. Check with `grep -rn '^\[\^[0-9]*\]:\s*$'`. |
| A nested list inside a list item becomes a code block | Indented continuation after a blank line looks like indented code | Code fencing looks at the previous non-blank line and refuses blocks that start with a list marker. |
| Raw `<table>` of 27 columns in the title page file | Title pages are layout tables | Leaf cells are flattened to text lines; styled `<span>`s are stripped. |

## Splitting

- Split happens top-down: a section is one file unless it is longer than `--max-lines` and has children at a level ≤ `--split-level`; then its own text goes to `NN-slug/README.md` and each child is visited recursively.
- File names are `NN-<transliterated title, up to 6 words, ≤50 chars>` with a few stopwords dropped (описание, метод, сервиса, …). They are stable across reruns as long as the document structure is unchanged, so hand edits can be re-applied with a diff.
- Descriptions in frontmatter and README are the first prose paragraph of the file, trimmed to 200 chars, with links and emphasis stripped.

- After deleting a false file or renaming folders, run `convert.py x <output-dir> --index-only`: it rebuilds README (keeping the intro above `## Содержание`), llms.txt and llms-full.txt from the files on disk.

## Known gaps

- Diagrams are PNG; nothing is OCR'd. The scenario step tables next to them usually carry the same information.
- Figure/table numbering is sequential within the document and may drift from the original if a caption was not detected. The in-text references keep the original numbers.
- Header-less key-value tables get an empty header row.
- Symbol bullets from Word (•, ○) are sometimes dropped by pandoc, leaving flat paragraphs; restore them by hand when a list clearly lost its markers.
- Word documents with tracked changes: accept all changes in Word/LibreOffice first; pandoc shows both old and new text otherwise (`--track-changes=accept` can be added to the pandoc call if needed).
