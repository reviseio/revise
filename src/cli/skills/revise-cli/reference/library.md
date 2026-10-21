# revise as a TypeScript library

Read this when the work is code, not shell commands: a Node script or service
that edits documents, your own model loop with Revise's editing tools, or a
collaboration server. Install `@reviseio/cli` (the same package as the CLI); the
CLI is a shell over this API. ESM only, Node 20+.
Full docs: https://revise.io/developer/docs/revise/library.md

```ts
import { writeFile } from "node:fs/promises";
import { ReviseDocument } from "@reviseio/cli";

const doc = await ReviseDocument.open("contract.docx", { author: "Legal bot" });
await doc.replaceAll("Provider", "Northstar"); // tracked changes
await doc.reject({ section: "8" });
const { bytes, fidelity } = await doc.export({ format: "docx" });
await writeFile("contract-out.docx", bytes);
doc.dispose();
```

## Open

- `ReviseDocument.open(source, options?)`: a path, or `{ bytes, filename }`
  (.docx, .odt, .rtf, .md, .html, .txt). A Word file keeps its package, so
  untouched content exports byte for byte.
- `ReviseDocument.create({ html?, markdown?, title? })`, `ReviseDocument.restore(snapshot)`,
  `ReviseDocument.fromYDoc(ydoc)` (a Y.Doc you own, such as a live room).
- Options: `author`, `mode` (`"suggesting"`: tracked changes, the default for
  Word files; `"editing"`: direct), `documentId`, `images`.
- Calls run one at a time in order; call `dispose()` when done.

## Read

`status()`, `outline()`, `summary()`, `grep(text, { ignoreCase })` →
`[{ id, blockIndex, count, text }]`.

## Edit

- `replaceAll(find, replace, { ids? })` → `{ replaced, blocks, changeIds, failures }`.
- Everything else is an editing tool: `doc.tools.definitions()` lists them with
  JSON schemas; `doc.tools.call(name, input)` runs one with typed input;
  `doc.tools.run(name, input)` with untyped input, as a model sends it. Neither
  throws for a bad edit: `{ ok: true, value }` or `{ ok: false, error }`.
- `await doc.setMode("editing" | "suggesting")`.

## Review

`changes({ section? })`, `preview(selection)`, `accept(selection)`,
`reject(selection)` → `{ resolved, missing, unresolved }`, `diff()`. A selection
is exactly one of `ids`, `all`, or `section` / `exceptSection` (`"8"`, `"8."`,
or the heading text).

## Output

`export({ format, view? })` with `docx`, `pdf`, `markdown`, `html`, `txt`;
`view: "accepted"` accepts everything, `"original"` rejects it. Returns `bytes`, `mode`, `view`,
`pendingTrackedChanges`, `fidelity` (.docx: `"graft"` kept untouched content
byte for byte, `"regenerated"` wrote it fresh), `warnings`.
`renderPages({ startPage, pageCount, view })` → one PNG (needs `@napi-rs/canvas`).
`compare(a, b)` → a redline document; `convert(source, format)` → bytes.
Raw package: `listParts`, `catPart`, `patchPackage`.

## Your own model loop

`agentTools(doc, { provider })` returns the tools in the shape a model API
expects and a `handle` that runs one call and returns the result message.
Providers: `"anthropic"` (Messages `tool_use` → `tool_result`), `"openai"` (Chat
Completions `tool_calls` → `role: "tool"`), `"openai-responses"`
(`function_call` → `function_call_output`). No model SDK is a dependency.

```ts
import Anthropic from "@anthropic-ai/sdk";
import { ReviseDocument, agentTools } from "@reviseio/cli";

const anthropic = new Anthropic();
const doc = await ReviseDocument.open("contract.docx", {
  author: "Contract bot",
});
const agent = agentTools(doc, { provider: "anthropic" });
const messages: Anthropic.MessageParam[] = [
  { role: "user", content: "Change the governing law to New York." },
];
for (;;) {
  const reply = await anthropic.messages.create({
    model: "claude-sonnet-5-5",
    max_tokens: 4096,
    tools: agent.tools,
    messages,
  });
  messages.push({ role: "assistant", content: reply.content });
  const calls = reply.content.filter((block) => block.type === "tool_use");
  if (calls.length === 0) break;
  messages.push({
    role: "user",
    content: await Promise.all(calls.map(agent.handle)),
  });
}
```

`handle` never throws for a bad call; the model gets the error and corrects
itself. Options: `include`, `exclude`, `maxResultChars`. After compacting the
transcript, call `agent.reset()`. Docs: https://revise.io/developer/docs/revise/agents.md

## Collaboration rooms

Every Revise document is a Y.js document. On a sync server:
`seedYDoc(room, "contract.docx")` seeds an empty room (safe on every
connection); `ReviseDocument.fromYDoc(room, { author })` edits a live room as
tracked changes that appear in every open Revise editor (the `revise` package); `export()` downloads
it. Dispose the ReviseDocument, not the room. Docs:
https://revise.io/developer/docs/revise/collaboration.md

## Errors

Expected failures (missing file, unsupported format, malformed patch) throw
`ReviseError` with a stable `code`. Tool calls return errors as values.
