# @reviseio/cli

[![](https://img.shields.io/npm/v/@reviseio/cli)](https://www.npmjs.com/package/@reviseio/cli) [![](https://img.shields.io/badge/read_the-docs-blue)](https://revise.io/developer/docs/revise/cli)

The swiss army knife for documents.

Read, edit, and redline Word documents. Available as CLI and TypeScript library.

```bash
npm i -g @reviseio/cli    # installs the `revise` command
npx @reviseio/cli read contract.docx    # or run it once without installing
```

Looking for the React document editor? That's [`revise`](https://www.npmjs.com/package/revise).

## For Agents

The `revise` CLI makes AI agents faster and cheaper.

Claude Code and Codex finish the same document tasks about 1.5-2× faster and using 30-70% less usage compared to their default `docx` toolset. See our [benchmarks on GitHub](https://github.com/reviseio/revise/deployments/benchmark).

To get started, tell your agent:

```text
Read https://revise.io/skills/cli.md and follow it.
```

Or install the skill directly from a terminal:

```bash
npx -y @reviseio/cli@latest skill install
```

## CLI

The `revise` CLI lets you inspect, edit, and convert .docx files.

```bash
revise outline contract.docx                              # headings with block ranges
revise read contract.docx 0                               # blocks as HTML, with ids
revise replace-all contract.docx "Provider" "Northstar"  # a tracked change, saved in place
revise replace-all contract.docx "net 30" "net 45" --author "Legal bot" --out v2.docx
revise changes contract.docx                              # what's pending
revise accept contract.docx --all
revise reject contract.docx --section 8
revise undo contract.docx
revise compare v1.docx v2.docx --out redline.docx
revise export contract.docx contract.pdf                  # .docx .pdf .md .html .txt
cat notes.md | revise convert - notes.docx
```

`revise` was designed primarily for AI agents, and has a whole suite of agent-ready tools for granular editing:

```bash
revise tools                        # every editing tool
revise call contract.docx set_header_footer '{"side":"footer","center":"Page {PAGE} of {PAGES}"}'
revise status contract.docx --json  # one versioned JSON object
revise examples                     # more
```

[JSON output contract](https://revise.io/developer/docs/revise/json)

## Agent skills

Quick install:

```bash
revise skill install                                                               # anything else
```

Claude:

```bash
claude plugin marketplace add reviseio/skills && claude plugin install cli@revise  # Claude Code
```

Codex:

```bash
codex plugin marketplace add reviseio/skills && codex plugin add cli@revise        # Codex
```

Via `skills`:

```bash
npx skills add reviseio/skills --skill revise-cli                                  # Cursor, Copilot, Windsurf…
```

## Library

`@reviseio/cli` is also a [TypeScript library](https://revise.io/developer/docs/revise/library).

```ts
import { ReviseDocument } from "@reviseio/cli";

const doc = await ReviseDocument.open("contract.docx", { author: "Legal bot" });
await doc.replaceAll("Provider", "Northstar");
await doc.reject({ section: "8" });
const { bytes } = await doc.export({ format: "docx" });
doc.dispose();
```

Give the editing tools to a model:

```ts
import Anthropic from "@anthropic-ai/sdk";
import { ReviseDocument, agentTools } from "@reviseio/cli";

const doc = await ReviseDocument.open("brief.docx");
const agent = agentTools(doc, { provider: "anthropic" }); // or "openai"
const reply = await new Anthropic().messages.create({
  model,
  max_tokens,
  messages,
  tools: agent.tools,
});
const results = await Promise.all(
  reply.content.filter((b) => b.type === "tool_use").map(agent.handle),
);
messages.push(
  { role: "assistant", content: reply.content },
  { role: "user", content: results },
);
```

Edit a live collaboration room (Y.js):

```ts
import { ReviseDocument, seedYDoc } from "@reviseio/cli";

await seedYDoc(room, "contract.docx"); // safe on every connection
const doc = await ReviseDocument.fromYDoc(room, { author: "Review bot" });
await doc.replaceAll("draft", "final");
const { bytes } = await doc.export({ format: "docx" });
```

## License

`@reviseio/cli` is available under the MIT License.

[`@reviseio/engine`](https://www.npmjs.com/package/@reviseio/engine) is free for any use under its own license.
