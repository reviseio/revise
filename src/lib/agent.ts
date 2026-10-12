/**
 * Revise's editing tools for your own agent loop, in the shape each model
 * API expects: hand `tools` to the model, pass each tool call it makes to
 * `handle`, and send back what `handle` returns.
 *
 *   const agent = agentTools(doc, { provider: "anthropic" });
 *   const reply = await anthropic.messages.create({ model, tools: agent.tools, messages });
 *   const results = await Promise.all(
 *     reply.content.filter((b) => b.type === "tool_use").map(agent.handle),
 *   );
 *   messages.push({ role: "assistant", content: reply.content }, { role: "user", content: results });
 *
 * OpenAI Responses: pass `agent.tools` as `tools`, append every output item
 * the model returned (reasoning and function_call items included) to your
 * input, then the `function_call_output` items `handle` returns, and call
 * again until the model stops calling tools.
 *
 * No model SDK is a dependency: plain objects in, plain objects out. Every
 * call runs on the document's one tool session, so reads, searches, and
 * edits share context across turns, and calls run one at a time (parallel
 * tool calls are safe; they queue). The session remembers what the model
 * has read: when your host compacts or drops the transcript, call
 * `reset()` so it stops assuming the model still has that text. Tool
 * descriptions mention "turns": here a turn is one model response.
 */
import type { ReviseJsonObject, ReviseToolResult } from "@reviseio/engine";
import type { ReviseDocument } from "./document";

export type AgentProvider = "anthropic" | "openai" | "openai-responses";

export interface AgentToolsOptions {
  provider: AgentProvider;
  /** Cap on one tool result's text, in characters (default 100,000). Longer
   * results are cut with a note telling the model how to page. */
  maxResultChars?: number;
  /** Only these tools (by name). */
  include?: string[];
  /** Every tool except these. */
  exclude?: string[];
}

/** Anthropic Messages API: a `tool_use` content block. */
export interface AnthropicToolUse {
  type?: "tool_use";
  id: string;
  name: string;
  input: unknown;
}

/** OpenAI Chat Completions: an entry of `message.tool_calls`. */
export interface OpenAIToolCall {
  id: string;
  type?: "function";
  function: { name: string; arguments: string };
}

/** OpenAI Responses API: a `function_call` output item. */
export interface OpenAIResponsesFunctionCall {
  type?: "function_call";
  call_id: string;
  name: string;
  arguments: string;
}

interface ProviderShapes {
  anthropic: {
    tool: {
      name: string;
      description: string;
      input_schema: { type: "object"; [key: string]: unknown };
    };
    call: AnthropicToolUse;
    result: {
      type: "tool_result";
      tool_use_id: string;
      content: string;
      is_error?: boolean;
    };
  };
  openai: {
    tool: {
      type: "function";
      function: {
        name: string;
        description: string;
        parameters: { [key: string]: unknown };
      };
    };
    call: OpenAIToolCall;
    result: { role: "tool"; tool_call_id: string; content: string };
  };
  "openai-responses": {
    tool: {
      type: "function";
      name: string;
      description: string;
      parameters: { [key: string]: unknown };
      /** Revise's schemas use optional fields, which strict mode rejects. */
      strict: false;
    };
    call: OpenAIResponsesFunctionCall;
    result: { type: "function_call_output"; call_id: string; output: string };
  };
}

export interface AgentTools<P extends AgentProvider> {
  /** Tool definitions in the provider's format. */
  tools: Array<ProviderShapes[P]["tool"]>;
  /** Run one tool call; returns the provider's tool-result message. Never
   * throws for a bad call: the model gets the error to correct itself. */
  handle: (
    call: ProviderShapes[P]["call"],
  ) => Promise<ProviderShapes[P]["result"]>;
  /** Did the last handled calls change the document? Compare to decide
   * whether to save. */
  version: () => Uint8Array;
  /** Forget what the model has read (after compacting its transcript). */
  reset: () => Promise<void>;
}

/** A tool result as the text a model reads: the message, any tracked
 * change ids, the document HTML it loaded, and its structured data (counts,
 * paging, what changed), within a size budget. */
export function formatToolResult(
  result: ReviseToolResult<string>,
  maxChars = 100_000,
): {
  text: string;
  isError: boolean;
} {
  if ("error" in result) {
    return {
      text: `error ${result.error.code}: ${result.error.message}`,
      isError: true,
    };
  }
  const value = result.value;
  const lines: string[] = [];
  if (value.message) lines.push(value.message);
  if (value.suggestionIds?.length) {
    lines.push(`Tracked changes: ${value.suggestionIds.join(", ")}`);
  }
  if (value.data !== null && value.data !== undefined) {
    const data = JSON.stringify(value.data);
    if (data !== "{}" && data !== "[]") {
      lines.push(
        `data: ${data.length <= 4_000 ? data : compactData(value.data)}`,
      );
    }
  }
  if (value.context?.html) lines.push("", value.context.html);
  let text = lines.join("\n") || "Done.";
  if (text.length > maxChars) {
    text = `${text.slice(0, maxChars)}\n… (result cut at ${maxChars} characters; read fewer blocks at a time, e.g. a later index or specific ids)`;
  }
  return { text, isError: false };
}

/** Large structured data, kept valid and useful: every scalar field (counts,
 * paging, ids), and arrays summarized by length. */
function compactData(data: unknown): string {
  const summarize = (value: unknown, depth: number): unknown => {
    if (Array.isArray(value)) {
      return value.length <= 5 && depth < 2
        ? value.map((item) => summarize(item, depth + 1))
        : `[${value.length} items]`;
    }
    if (value && typeof value === "object") {
      if (depth >= 2) return "{…}";
      return Object.fromEntries(
        Object.entries(value).map(([key, inner]) => [
          key,
          summarize(inner, depth + 1),
        ]),
      );
    }
    return typeof value === "string" && value.length > 200
      ? `${value.slice(0, 200)}…`
      : value;
  };
  return `${JSON.stringify(summarize(data, 0))} (summarized; the full data is larger)`;
}

export function agentTools<P extends AgentProvider>(
  doc: ReviseDocument,
  options: AgentToolsOptions & { provider: P },
): AgentTools<P> {
  const definitions = doc.tools
    .definitions()
    .filter((d) => !options.include || options.include.includes(d.name))
    .filter((d) => !options.exclude?.includes(d.name));
  const allowed = new Set<string>(definitions.map((d) => d.name));

  const tools = definitions.map((d) => {
    switch (options.provider) {
      case "anthropic":
        return {
          name: d.name,
          description: d.description,
          input_schema: { ...d.inputSchema, type: "object" as const },
        };
      case "openai":
        return {
          type: "function" as const,
          function: {
            name: d.name,
            description: d.description,
            parameters: d.inputSchema,
          },
        };
      default:
        return {
          type: "function" as const,
          name: d.name,
          description: d.description,
          parameters: d.inputSchema,
          strict: false as const,
        };
    }
  }) as AgentTools<P>["tools"];

  const run = async (
    name: string,
    rawInput: unknown,
  ): Promise<{ text: string; isError: boolean }> => {
    if (!allowed.has(name)) {
      return {
        text: `error unknown_tool: There is no tool named "${name}".`,
        isError: true,
      };
    }
    let input: unknown = rawInput;
    if (typeof rawInput === "string") {
      try {
        input = rawInput.trim() ? JSON.parse(rawInput) : {};
      } catch (error) {
        return {
          text: `error invalid_json: The arguments are not valid JSON (${(error as Error).message}).`,
          isError: true,
        };
      }
    }
    if (!input || typeof input !== "object" || Array.isArray(input)) {
      return {
        text: "error invalid_arguments: Tool arguments must be a JSON object.",
        isError: true,
      };
    }
    return formatToolResult(
      await doc.tools.run(name, input as ReviseJsonObject),
      options.maxResultChars,
    );
  };

  const handle = async (
    call: ProviderShapes[P]["call"],
  ): Promise<ProviderShapes[P]["result"]> => {
    switch (options.provider) {
      case "anthropic": {
        const c = call as AnthropicToolUse;
        const { text, isError } = await run(c.name, c.input);
        return {
          type: "tool_result",
          tool_use_id: c.id,
          content: text,
          ...(isError ? { is_error: true } : {}),
        } as ProviderShapes[P]["result"];
      }
      case "openai": {
        const c = call as OpenAIToolCall;
        const { text } = await run(c.function.name, c.function.arguments);
        return {
          role: "tool",
          tool_call_id: c.id,
          content: text,
        } as ProviderShapes[P]["result"];
      }
      default: {
        const c = call as OpenAIResponsesFunctionCall;
        const { text } = await run(c.name, c.arguments);
        return {
          type: "function_call_output",
          call_id: c.call_id,
          output: text,
        } as ProviderShapes[P]["result"];
      }
    }
  };

  return {
    tools,
    handle,
    version: () => doc.version(),
    reset: () => doc.resetToolContext(),
  };
}
