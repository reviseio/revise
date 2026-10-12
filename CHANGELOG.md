# Changelog

## 3.1.4

### Fixed

- **The package README's heading** reads `revise`.

## 3.1.3

### Changed

- **The package README** has a plain heading and tagline. No changes to the
  CLI or the library.

## 3.1.2

### Changed

- **`revise skill install` installs where the agent running it reads
  skills:** `.claude/skills` from Claude Code or in a project with a
  `.claude` folder, otherwise `.agents/skills`, which Codex, Cursor, Copilot,
  and others read. It used to fall back to `.claude/skills`, which other
  agents don't read.
- **It ends by saying to read the installed SKILL.md now:** a running agent
  session doesn't load a newly installed skill; new sessions do. The
  quickest install for any agent with a terminal is
  `npx -y revise@latest skill install`.

## 3.1.1

### Changed

- **The agent skill is `revise-cli`.** `revise skill install` writes
  `.agents/skills/revise-cli/` (or `.claude/skills/revise-cli/`) and replaces
  the `revise/` folder earlier versions installed. As plugins it's
  `cli@revise`: `claude plugin install cli@revise`,
  `codex plugin add cli@revise`, or
  `npx skills add reviseio/skills --skill revise-cli`.

## 3.1.0

### Added

- **`--blocks` on `changes`, `accept`, and `reject`.** Select tracked
  changes by block: ids (`--blocks a1B2c3,d4E5f6`) or a range in document
  order (`--blocks a1B2c3..d4E5f6`). Ids of list items, table cells, and
  paragraphs inside them work. In the library: `{ blocks: [...] }`, with
  ranges as `{ from, to }`.
- **`--view tracked|accepted|original`** on `read`, `grep`, `search`,
  `export`, and `convert`. `read` now shows pending changes inline as
  `<del>` and `<ins>` with their change ids and authors; `grep` and `search`
  find text on both sides of a change and say which side it's on. `export`
  and `convert` with `--view original` write the document as it was before
  every pending change. `--final` still means `--view accepted`.
- **`replace-all --ignore-case` and `--exact`.**

### Changed

- **`replace-text` is now `replace-all`** (`replaceAll` in the library). It
  replaces in page headers and footers and in footnotes and endnotes as well
  as the body, and says where. Curly and straight quotes, no-break and plain
  spaces, and no-break and plain hyphens match each other unless `--exact`,
  and the replacement keeps the document's characters: `"net 30"` →
  `"net 45"` keeps its no-break space. Matches it leaves alone, like other
  capitalization, are reported as `nearMisses`. `replace-text` and
  `replaceText` still work, and JSON output run under `replace-text` still
  names it.
- **`--section` is a heading and everything under it.** `--section 8`
  includes 8.1 and 8.2. Numbers match whether typed or Word's automatic
  numbering. A title matches in full, or by text in only one heading; text in
  several headings is an error (`section_ambiguous`), and a section that
  doesn't exist is `section_not_found` instead of an empty result.
- **Clearer flag names:** `new --content` (was `--from`), `new --page-size`
  (was `--page`), `search --results-page` (was `--page`), and `--blocks` for
  block ids on `read` and `replace-all` (was `--ids`). The old names still
  work.
- **`changes` lists header, footer, and note changes as such** ("Header",
  "Footnote 1"), with their text.

### Fixed

- **Word-numbered headings.** Headings numbered by Word's heading styles, how
  most contracts number clauses, are headings again: `outline` lists them
  with their numbers ("8.1 Cap on liability") and `--section 8` finds them.
- **Comparing numbered documents.** Two copies of the same numbered document
  compare as unchanged, instead of every numbered clause deleted and
  reinserted.
- **Comparing against a file revise edited.** A one-word edit compares as a
  one-word change, not a whole paragraph replaced: editing a paragraph no
  longer adds fonts and sizes to it or its neighbors.
- **Tracked changes in headers and footers** are saved as Word revisions, and
  ones made in Word open as pending changes.
- **Markdown export** numbers nested list items per level (8.1 was written
  as "9."), and writes numbered headings as headings: `## 8.1 Cap on
liability`.

## 3.0.1

### Fixed

- **Windows-1252 text on Node 20.** Text that isn't UTF-8, like a .txt file
  from a Windows editor piped to `convert`, and RTF files read € “ ” ‘ ’ – —
  and the other characters Windows-1252 stores in bytes 0x80–0x9F as
  invisible control characters. Node 22 and later were not affected.

## 3.0.0

The first release of `revise` as a command-line tool and TypeScript library
for reading, editing, and redlining Word documents.

Versions 2.x of this package were the Revise editor for React. The editor is
now [`@reviseio/editor`](https://www.npmjs.com/package/@reviseio/editor); 2.x
stays installable.

### Added

- **The `revise` CLI.** Read a document as HTML with block ids, search it,
  edit it with the same tools the Revise agent uses, and save the file in
  place, as tracked changes by default in .docx. Review with `changes`,
  `diff`, `accept`, and `reject`; `undo` and `restore` work across edits;
  `render` draws pages as images; `export`, `convert`, `compare`, and `new`
  write .docx, PDF, Markdown, HTML, and text. `convert` also works in a
  pipe: `-` reads stdin (the format is recognized from its bytes) or writes
  stdout. `--out new.docx` on an edit saves the result as a new file and
  leaves the original as it is. Every command has its own
  `--help`, and colored output on a terminal (not in pipes unless
  `FORCE_COLOR` is set; never with `--json`).
- **`--json` on every command**: one envelope (`schema`, `ok`, `command`,
  `meta`, `data`, `warnings`, `error`) with stable error codes, structured
  partial results, and exit codes that say whether anything was saved.
  Specified in [JSON output](https://revise.io/developer/docs/revise/json) and shipped as a JSON Schema,
  `revise/cli-output.schema.json`.
- **`ReviseDocument`**, the same operations as a library: open a file or
  bytes, run typed tools, review changes, export with a fidelity report,
  and snapshot and restore state.
- **`agentTools(doc, { provider })`**: the editing tools in the shapes the
  Anthropic and OpenAI SDKs expect, with a matching result formatter.
- **Collaboration rooms.** `seedYDoc(room, file)` seeds a sync server's
  Y.Doc on every connection (idempotent, repairs two peers that seeded at
  once, keeps the original Word file in the room), and
  `ReviseDocument.fromYDoc(room)` edits and exports a live room.
- **Word fidelity.** Untouched content in a .docx exports byte for byte;
  `xml` commands read and patch the raw package for what the tools don't
  cover (styles, numbering, settings).
- **An agent skill**: `revise skill install` writes a SKILL.md that teaches
  coding agents the CLI.

### Removed

- **The editor entry points of 2.x.** `revise/editor`, `revise/editor/full`,
  and `revise/server` now throw an error pointing to `@reviseio/editor`, and
  `revise/style.css` is an empty stylesheet that says the same.
