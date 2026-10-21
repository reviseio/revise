# Changelog

## 4.0.1

Released with `revise` 4.0.1. No changes to the CLI or library.

## 4.0.0

The CLI and TypeScript library are the `@reviseio/cli` package; the command is
`revise`. Released with `revise` 4.0.0 (the editor) and `@reviseio/engine`
4.0.0.

### Added

- **Lists of figures and tables.** `<toc source="captions">` in tool HTML
  lists figure captions with page numbers; `label="Table"` or any other
  label collects that kind. A Word `TOC \c "Label"` field in a .docx opens
  as one and is written back as the same field, and it renders in PDF,
  Markdown, RTF, and text output.
- **`join_paragraphs` takes `ranges`.** Up to 100
  `{ firstId, lastId, separator? }` ranges join runs of consecutive
  paragraphs, such as PDF line-wrapped text, in one call. If any range is
  invalid, every problem is reported and nothing changes.
- **Comment anchors across wrapped lines.** A `leave_comment` anchor or
  start/end pair whose text continues into the following paragraphs is
  anchored across them when every spanned block has been read.
- **Cell borders per edge and padding shorthand** in `style_blocks`
  (`borderTop`, `borderRight`, `borderBottom`, `borderLeft`; `none` for no
  border), and `fontVariant: "small-caps"`.

### Changed

- **Install `@reviseio/cli`.** `npm i -g @reviseio/cli` for the command,
  `npx @reviseio/cli` to run it once, and `@reviseio/cli` for library
  imports. Remove an earlier global install of `revise` first; see
  [upgrading](https://revise.io/developer/docs/revise/introduction#upgrade-from-an-earlier-version).
- **New table rows and columns match their neighbors**
  (`insert_table_row`, `insert_table_column`): borders, padding, shading,
  and vertical alignment come from the adjacent cell, rows under a shaded
  header take the body's formatting, and alternating row shading continues.
- **`replace_block` keeps paragraph layout.** Alignment, indents, spacing,
  line height, named style, and direction carry over unless the HTML sets
  them.
- **A `replace` find may span paragraphs.** A plain-text `find` with line
  breaks can match the target block and the paragraphs right after it.
- **`join_paragraphs` keeps line-end hyphens.** "exam-" + "ple" joins as
  "exam-ple" when no `separator` is given.
- **`remove_blocks` needs no read for page breaks and empty paragraphs.**
  Section breaks and formatted empty paragraphs still do.
- **Empty numbered items show their number** in `read` output
  (`data-marker`).
- **Clearer tool errors** for negative `leftIndent`, indents on list items,
  and `edit_table_cells` targets inside merged cells.

### Fixed

- **Pending footnotes and endnotes stay with their references** through
  later edits. An edit that would rebuild a block holding one fails with
  `pending_footnote_reference` and changes nothing. A footnote reference
  inside a tracked insertion or deletion in a .docx opens with its note as
  part of the same pending change.
- **`replace` no longer rewrites similar paragraphs next to an exact
  match** in a table or list, and `replace_all` never fuzzy-matches.
- **Fewer "not fully read" refusals on tables:** inserting next to a partly
  read table and exact-text `replace` inside one no longer need a full read.
- **`th` in `style_blocks` selectors** matches only a table's header row
  when it has one.
- **Your own edits no longer count as someone else's.** A block restyled by
  a broad `style_blocks` can be edited next without a re-read, your edits no
  longer invalidate your other search results, and cells of a row or column
  you inserted as a suggestion can be filled.
- **Images in `replace_block`** can become text, and an image replacement
  keeps the original's settings and takes px width and height.
- **A `replace` matching an image's alt text** fails with a pointer to
  `replace_block` instead of reporting success without a change.
- **`letterSpacing: "normal"`** is accepted.
- **Web-search citation markers** from OpenAI models are removed from tool
  arguments.
- **PDF output** embeds TrueType fonts with spec-conformant dictionaries, so
  strict PDF readers and printers handle it.

### Security

- **jsPDF 4.2.1.** PDF output is written with a fork of jsPDF 4.2.1
  (was 3.0.1), which clears the advisories against 3.0.1. The third-party
  notices list jsPDF and the packages it inlines.

## 3.1.7

Released with the editor 3.1.7: one version and one build for all three
packages. A Y.Doc or .docx moves between the editor and the library
unchanged.

### Changed

- **`remove_blocks` takes up to 200 ids** (was 50). Past the limit it removes
  nothing and says how to proceed: pass a search's `search_result_id` to
  remove every match, or send the ids in batches.

### Fixed

- **Page numbers next to an existing header or footer.** `set_page_layout`
  with a page-number position, or a preset that numbers pages (APA puts them
  top right), used to report success while a hand-authored header or footer
  kept the number from appearing. The number now goes into that header's or
  footer's empty zone at the requested side, in the authored text's font. When
  the zone already holds text or the header already shows a page number, the
  result says so and lists the headers and footers instead of claiming the
  layout changed.
- **A `replace` whose `find` and `replace` are identical** fails with an error
  that says so, instead of the generic "no changes detected".

## 3.1.6

### Changed

- **The agent skill knows when revise can't run.** In the Claude and ChatGPT
  apps, a document attached to the chat lives in a sandbox that can't install
  revise. The skill now stops before running anything there and offers the
  Revise app, which opens the document as an editable editor in the chat. If
  the app is connected, it uses it. Agents with a terminal are unaffected, and
  an install that fails on a network error is treated the same way instead of
  being worked around.

## 3.1.5

Agents finish document tasks in fewer commands. Measured on the revise
benchmark, Claude Code and Codex finish its tasks for about 40% less with
revise than without it.

### Added

- **`revise batch <file> '[...]'`** runs several tool calls (for example a
  table row and the total under it, or a `replace_all` and a `replace`) as one
  edit with one save. If any step fails, nothing is saved.
- **`read --section 4,7,9`** reads several sections at once, `--blocks a..b`
  reads a range of ids, and `--count N` limits how many blocks are shown.
- **`grep` takes several phrases** and shows about 80 characters around each
  match instead of the whole block (at most 40 blocks).
- **`outline` lists headings typed as bold body text** first, each with the
  command that turns it into a real heading.
- **PDF exports print the page count and paper size.**

### Changed

- **One `replace` call can edit several blocks:** leave out `id` and each
  `find` edits the one block you've read or grepped that contains it.
- **Saves say what they rewrote:** "Rewrote N edited blocks; the rest of the
  file is copied byte for byte."
- **`read` output is shorter:** it leaves out formatting that writing the
  block back restores anyway (a heading's own style, default table cell
  borders and padding, and the like).
- **`outline` shows one id range per section** (`first..last`) and its word
  count.
- **`compare` says what differs outside the redline:** header and footer
  text and page setup, or that the rest is formatting only.
- **`--dry-run` with `--out`** previews instead of failing.
- **The agent skill is a third shorter** and lists every field of `replace`
  and `leave_comment`, so agents don't need to look them up.

### Fixed

- **`read --blocks a,b`** shows only those blocks, not every block between
  them.
- **`outline` and `read --section`** see every section, not just the first 100.
- **Rows added with `insert_table_row`** take the alignment, spacing,
  indents, and font of the row they're copied from.

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

The first release of the command-line tool and TypeScript library for
reading, editing, and redlining Word documents.

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
