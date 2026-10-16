# doc-to-llm-docs

A [Claude Code](https://claude.com/claude-code) skill that turns a large Word document (`.doc`, `.docx`, `.odt`, `.rtf`, or a Word "web page" `.mht` saved as `.doc`) into a folder of LLM-friendly Markdown: one self-contained file per section, frontmatter, restored section numbers, Markdown tables, fenced code, an index `README.md`, `llms.txt` and `llms-full.txt`.

Built for 100+ page specifications with a table of contents, numbered sections, parameter tables and XML/JSON examples (it was developed on Russian ЕГИСЗ integration specs: РЭМД, ФРМО, ФРМР, ИПС), but nothing in it is specific to that domain.

## Why

`pandoc` alone gives you one 8 000-line file with a 300-line table of contents, no section numbers (Word auto-numbering is not text), HTML tables for every multi-line cell, appendix headings rendered as list items and images pointing at temp paths. macOS `textutil` is worse: it silently drops every table from `.doc` files. The bundled script fixes all of that and splits the document by heading so a model can load only the section it needs.

## Install

As a plugin (recommended):

```
/plugin marketplace add kovalevvv/doc-to-llm-docs
/plugin install doc-to-llm-docs@kovalevvv
```

Or copy the skill folder:

```bash
git clone https://github.com/kovalevvv/doc-to-llm-docs
cp -R doc-to-llm-docs/skills/doc-to-llm-docs ~/.claude/skills/
```

Dependencies on the machine where Claude Code runs:

```bash
brew install pandoc                     # required
brew install --cask libreoffice         # only for .doc / .odt / .rtf input
python3 -m pip install beautifulsoup4   # required
```

## Use

Ask Claude Code in your own words, e.g. "преврати `~/Downloads/spec.docx` в документацию для LLM" or "split this Word spec into markdown for RAG". The skill runs the converter, reads its report, spot-checks tables, code blocks, images and footnotes, and writes the README intro by hand.

The script can also be run directly:

```bash
python3 -I skills/doc-to-llm-docs/scripts/convert.py "spec.docx" out/ --title "Spec title" [--split-level 4] [--bold-headings]
python3 -I skills/doc-to-llm-docs/scripts/convert.py x out/ --index-only   # rebuild the index after manual edits
```

Output:

```
out/
  README.md          index with a one-line description of every file (+ hand-written intro)
  llms.txt           same index in llmstxt.org format
  llms-full.txt      the whole document concatenated, for one-shot context loading
  00-title-page.md   title page and version history
  media/             extracted images
  NN-slug.md         one file per section that fits --max-lines
  NN-slug/README.md  intro of a split section, followed by NN-slug/NN-sub.md
```

## What the converter handles

See [`skills/doc-to-llm-docs/references/pitfalls.md`](skills/doc-to-llm-docs/references/pitfalls.md) for the full list. Highlights: TOC detection and removal with section numbers restored from it; headings that Word/LibreOffice export as list items or plain paragraphs; HTML tables to pipe tables with `rowspan`/`colspan`; request/response example tables to fenced blocks; indented and paragraph-per-line JSON to fenced blocks; multi-line footnotes; inline images; Word bookmark and field leaks; intra-document links rewritten to relative file links; a link check at the end.

Not handled: PDF input (run marker/docling first and feed the HTML), scanned documents, diagrams (kept as PNG).

## License

MIT
