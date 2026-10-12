// @vitest-environment node
// Ported from Revise's main repository on each release; changes land there.

import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { ReviseDocument, agentTools } from "../src/lib";

// The adapters as an agent loop uses them: provider-shaped tool lists in,
// provider-shaped tool calls through `handle`, provider-shaped results out.
// No model is called; the "model" here is a scripted list of tool calls.

const CONTRACT = fileURLToPath(
  new URL(
    "./fixtures/legal-negotiation-before.docx",
    import.meta.url,
  ),
);

describe("agentTools", () => {
  it("speaks each provider's tool format", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    try {
      const anthropic = agentTools(doc, { provider: "anthropic" });
      const replace = anthropic.tools.find((t) => t.name === "replace")!;
      expect(replace.input_schema).toMatchObject({ type: "object" });
      expect(replace.description.length).toBeGreaterThan(0);

      const openai = agentTools(doc, { provider: "openai" });
      expect(openai.tools[0]).toMatchObject({
        type: "function",
        function: { name: expect.any(String), parameters: expect.any(Object) },
      });

      const responses = agentTools(doc, { provider: "openai-responses" });
      expect(responses.tools[0]).toMatchObject({
        type: "function",
        name: expect.any(String),
        parameters: expect.any(Object),
      });

      const onlyReads = agentTools(doc, {
        provider: "anthropic",
        include: ["read_blocks_from_index", "search_document"],
      });
      expect(onlyReads.tools.map((t) => t.name).sort()).toEqual([
        "read_blocks_from_index",
        "search_document",
      ]);
    } finally {
      doc.dispose();
    }
  });

  it("runs a scripted Anthropic turn: search, then edit as a tracked change", async () => {
    const doc = await ReviseDocument.open(CONTRACT, { author: "Agent X" });
    try {
      const agent = agentTools(doc, { provider: "anthropic" });
      const before = Buffer.from(agent.version());

      const search = await agent.handle({
        type: "tool_use",
        id: "call_1",
        name: "search_document",
        input: { queries: ["terminate without cause"] },
      });
      expect(search).toMatchObject({
        type: "tool_result",
        tool_use_id: "call_1",
      });
      expect(search.is_error).toBeUndefined();
      const id =
        /<p id="([^"]+)"[^>]*>(?:(?!<\/p>).)*terminate without cause/s.exec(
          search.content,
        )?.[1];
      expect(id).toBeTruthy();

      const edit = await agent.handle({
        type: "tool_use",
        id: "call_2",
        name: "replace",
        input: {
          id,
          replacements: [{ find: "thirty days", replace: "ninety days" }],
        },
      });
      expect(edit.is_error).toBeUndefined();
      expect(edit.content).toContain("Tracked changes:");
      expect(Buffer.from(agent.version()).equals(before)).toBe(false);
      const changes = await doc.changes();
      expect(changes.some((c) => c.inserted.includes("ninety"))).toBe(true);
      expect(changes.every((c) => c.author === "Agent X")).toBe(true);
    } finally {
      doc.dispose();
    }
  });

  it("takes OpenAI calls with string arguments, and returns errors to the model", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    try {
      const chat = agentTools(doc, { provider: "openai" });
      const read = await chat.handle({
        id: "c1",
        type: "function",
        function: {
          name: "read_blocks_from_index",
          arguments: JSON.stringify({ index: 0 }),
        },
      });
      expect(read).toMatchObject({ role: "tool", tool_call_id: "c1" });
      expect(read.content).toContain("<");

      const badJson = await chat.handle({
        id: "c2",
        function: { name: "read_blocks_from_index", arguments: "{oops" },
      });
      expect(badJson.content).toMatch(/^error invalid_json/);

      const responses = agentTools(doc, { provider: "openai-responses" });
      const unknown = await responses.handle({
        type: "function_call",
        call_id: "r1",
        name: "delete_everything",
        arguments: "{}",
      });
      expect(unknown).toEqual({
        type: "function_call_output",
        call_id: "r1",
        output: expect.stringMatching(/^error unknown_tool/),
      });

      const anthropic = agentTools(doc, { provider: "anthropic" });
      const failed = await anthropic.handle({
        id: "a1",
        name: "replace",
        input: {
          id: "no-such-block",
          replacements: [{ find: "a", replace: "b" }],
        },
      });
      expect(failed.is_error).toBe(true);
    } finally {
      doc.dispose();
    }
  });
});
