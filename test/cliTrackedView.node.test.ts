// @vitest-environment node
// Ported from Revise's main repository on each release; changes land there.

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { main } from "../src/cli/main";
import { ReviseDocument } from "../src/lib";

// Reading a redline: read, grep, and search show the other side's pending
// changes by default, and --view picks accepted or original text instead.

const root = mkdtempSync(path.join(tmpdir(), "revise-view-"));
let redline: string;

async function revise(...argv: string[]) {
  let stdout = "";
  let stderr = "";
  const code = await main(argv, {
    stdout: (text) => {
      stdout += text;
    },
    stderr: (text) => {
      stderr += text;
    },
  });
  return { code, stdout, stderr };
}

async function reviseJson(...argv: string[]) {
  const result = await revise(...argv, "--json");
  const envelope = JSON.parse(result.stdout);
  return { ...result, json: envelope.data, error: envelope.error };
}

beforeAll(async () => {
  process.env.REVISE_WORKSPACE = path.join(root, ".revise");
  const draft = await ReviseDocument.create({
    html: "<p>Fees are $100 per hour.</p><p>Payment is due in 30 days.</p>",
  });
  const jane = await ReviseDocument.restore(await draft.snapshot(), {
    author: "Jane Smith",
  });
  draft.dispose();
  await jane.setMode("suggesting");
  await jane.replaceAll("$100", "$120");
  redline = path.join(root, "redline.docx");
  writeFileSync(redline, (await jane.export({ format: "docx" })).bytes);
  jane.dispose();
});

afterAll(() => {
  delete process.env.REVISE_WORKSPACE;
  rmSync(root, { recursive: true, force: true });
});

describe("--view", () => {
  it("read marks pending changes, by default", async () => {
    const read = await reviseJson("read", redline);
    expect(read.json.view).toBe("tracked");
    expect(read.json.html).toMatch(
      /\$<del data-change-id="[^"]+" data-author="Jane Smith">100<\/del><ins data-change-id="[^"]+" data-author="Jane Smith">120<\/ins> per hour/,
    );
    const human = await revise("read", redline);
    expect(human.stdout).toContain("<del");

    const accepted = await reviseJson("read", redline, "--view", "accepted");
    expect(accepted.json.html).toContain("$120 per hour");
    expect(accepted.json.html).not.toMatch(/<del|<ins/);
    const original = await reviseJson("read", redline, "--view", "original");
    expect(original.json.html).toContain("$100 per hour");
    expect(original.json.html).not.toMatch(/<del|<ins|\$120/);
  });

  it("grep finds deleted text, labeled", async () => {
    const tracked = await reviseJson("grep", redline, "$100");
    expect(tracked.json.view).toBe("tracked");
    expect(tracked.json.blocks).toHaveLength(1);
    expect(tracked.json.blocks[0]).toMatchObject({
      text: "Fees are $[-100-]{+120+} per hour.",
      matches: [{ text: "$100", change: "deleted" }],
    });
    expect((await revise("grep", redline, "$100")).stdout).toContain(
      "1 in deleted text",
    );
    expect(
      (await reviseJson("grep", redline, "$100", "--view", "accepted")).json
        .blocks,
    ).toEqual([]);
    expect(
      (await reviseJson("grep", redline, "$100", "--view", "original")).json
        .blocks[0].text,
    ).toBe("Fees are $100 per hour.");
  });

  it("search adds matches in deleted text; original searches the old text", async () => {
    const tracked = await reviseJson("search", redline, "$100");
    expect(tracked.json.view).toBe("tracked");
    expect(tracked.json.viewMatches).toEqual([
      expect.objectContaining({
        query: "$100",
        text: "Fees are $100 per hour.",
        changeIds: [expect.any(String)],
      }),
    ]);
    expect((await revise("search", redline, "$100")).stdout).toContain(
      "Also in deleted text",
    );
    const original = await reviseJson(
      "search",
      redline,
      "$100",
      "--view",
      "original",
    );
    expect(original.json.viewMatches).toHaveLength(1);
    expect(original.json.html).toBeNull();
    const accepted = await reviseJson(
      "search",
      redline,
      "$120",
      "--view",
      "accepted",
    );
    expect(accepted.json.viewMatches).toBeUndefined();
    expect(accepted.json.html).toContain("$120");
  });

  it("export and convert write the original text", async () => {
    const md = path.join(root, "original.md");
    const exported = await reviseJson(
      "export",
      redline,
      md,
      "--view",
      "original",
    );
    expect(exported.json).toMatchObject({ view: "original", final: true });
    expect(readFileSync(md, "utf8")).toContain("$100 per hour");

    const txt = path.join(root, "converted.txt");
    const converted = await reviseJson(
      "convert",
      redline,
      txt,
      "--view",
      "original",
    );
    expect(converted.json.view).toBe("original");
    expect(readFileSync(txt, "utf8")).toContain("$100 per hour");

    // A plain-text copy can't hold changes: tracked comes out accepted.
    const plain = path.join(root, "plain.txt");
    expect((await reviseJson("convert", redline, plain)).json.view).toBe(
      "accepted",
    );
    expect(readFileSync(plain, "utf8")).toContain("$120 per hour");
  });

  it("keeps --final as --view accepted, and refuses a contradiction", async () => {
    const out = path.join(root, "final.txt");
    const final = await reviseJson("convert", redline, out, "--final");
    expect(final.json.view).toBe("accepted");
    const both = await reviseJson(
      "read",
      redline,
      "--final",
      "--view",
      "original",
    );
    expect(both.error.code).toBe("usage");
    expect(
      (await reviseJson("read", redline, "--view", "nope")).error.code,
    ).toBe("usage");
    // Not in help: --view is the one name.
    expect((await revise("export", "--help")).stdout).not.toContain("--final");
  });
});
