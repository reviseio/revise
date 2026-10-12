// @vitest-environment node
// Ported from Revise's main repository on each release; changes land there.

/**
 * The --json contract (revise.io/developer/docs/revise/json, cli/output.schema.json):
 * every command's real output, successes and failures, checked against the
 * published schema, strictly (no undocumented keys), plus the invariants
 * the schema can't express: one object on stdout, nothing on stderr, exit
 * codes that agree with `ok`. And color: only where it belongs.
 */
import {
  chmodSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { main } from "../src/cli/main";
import { sniffFormat } from "../src/cli/sniff";
import { SKILL_MARKDOWN, SKILL_REFERENCES } from "../src/cli/skill";
import schema from "../src/cli/output.schema.json";

const fixture = (relative: string) =>
  fileURLToPath(new URL(`./fixtures/${relative}`, import.meta.url));
const CONTRACT = fixture("legal-negotiation-before.docx");
const CONTRACT_V2 = fixture(
  "legal-negotiation-after-clean.docx",
);

// ── a strict JSON Schema validator for the subset the schema uses ──────────

type Schema = Record<string, unknown>;

function resolve(ref: string): Schema {
  return ref
    .replace(/^#\//, "")
    .split("/")
    .reduce<Schema>((node, key) => node[key] as Schema, schema as Schema);
}

function typeOf(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  if (typeof value === "number")
    return Number.isInteger(value) ? "integer" : "number";
  return typeof value;
}

/** Problems with `value` against `node`. Strict: an object schema whose
 * `required` names every property allows no other keys. */
function check(value: unknown, node: Schema, at = "$"): string[] {
  if (node.$ref) return check(value, resolve(node.$ref as string), at);
  const problems: string[] = [];
  if (node.type) {
    const types = ([] as string[]).concat(node.type as string | string[]);
    const actual = typeOf(value);
    if (
      !types.includes(actual) &&
      !(actual === "integer" && types.includes("number"))
    ) {
      return [`${at}: ${actual}, expected ${types.join("|")}`];
    }
  }
  if ("const" in node && value !== node.const) {
    problems.push(
      `${at}: ${JSON.stringify(value)} is not ${JSON.stringify(node.const)}`,
    );
  }
  if (node.enum && !(node.enum as unknown[]).includes(value)) {
    problems.push(
      `${at}: ${JSON.stringify(value)} is not one of ${JSON.stringify(node.enum)}`,
    );
  }
  if (
    typeof node.minimum === "number" &&
    typeof value === "number" &&
    value < node.minimum
  ) {
    problems.push(`${at}: ${value} < ${node.minimum}`);
  }
  if (typeOf(value) === "object") {
    const object = value as Record<string, unknown>;
    const properties = (node.properties ?? {}) as Record<string, Schema>;
    const required = (node.required ?? []) as string[];
    for (const key of required) {
      if (!(key in object)) problems.push(`${at}: missing ${key}`);
    }
    for (const [key, sub] of Object.entries(properties)) {
      if (key in object)
        problems.push(...check(object[key], sub, `${at}.${key}`));
    }
    const strict =
      node.properties &&
      required.length === Object.keys(properties).length &&
      node.additionalProperties === undefined;
    for (const key of Object.keys(object)) {
      if (key in properties) continue;
      if (
        node.additionalProperties &&
        typeof node.additionalProperties === "object"
      ) {
        problems.push(
          ...check(
            object[key],
            node.additionalProperties as Schema,
            `${at}.${key}`,
          ),
        );
      } else if (strict || node.additionalProperties === false) {
        problems.push(`${at}: undocumented key ${key}`);
      }
    }
  }
  if (Array.isArray(value) && node.items) {
    value.forEach((item, i) =>
      problems.push(...check(item, node.items as Schema, `${at}[${i}]`)),
    );
  }
  if (node.oneOf) {
    const passing = (node.oneOf as Schema[]).filter(
      (sub) => !check(value, sub, at).length,
    );
    if (passing.length !== 1) {
      problems.push(
        `${at}: matches ${passing.length} of oneOf (${(node.oneOf as Schema[])
          .map((sub) => check(value, sub, at)[0] ?? "ok")
          .join(" / ")})`,
      );
    }
  }
  for (const sub of (node.allOf ?? []) as Schema[])
    problems.push(...check(value, sub, at));
  if (node.if) {
    const branch = check(value, node.if as Schema, at).length
      ? node.else
      : node.then;
    if (branch) problems.push(...check(value, branch as Schema, at));
  }
  return problems;
}

// ── running the CLI ────────────────────────────────────────────────────────

interface Run {
  code: number;
  stdout: string;
  stderr: string;
  /** Raw bytes written to stdout (`revise convert … -`). */
  stdoutBytes: Uint8Array;
}

async function run(
  argv: string[],
  options: {
    tty?: boolean;
    env?: Record<string, string>;
    stdin?: Uint8Array | string;
  } = {},
): Promise<Run> {
  let stdout = "";
  let stderr = "";
  const chunks: Uint8Array[] = [];
  const stdin = options.stdin;
  const code = await main(argv, {
    stdout: (text) => {
      stdout += text;
    },
    stderr: (text) => {
      stderr += text;
    },
    stdoutIsTTY: options.tty ?? false,
    stderrIsTTY: options.tty ?? false,
    env: options.env ?? {},
    stdinIsTTY: stdin === undefined,
    readStdin: async () =>
      typeof stdin === "string"
        ? new TextEncoder().encode(stdin)
        : (stdin ?? new Uint8Array()),
    writeStdout: (bytes) => {
      chunks.push(bytes);
      stdout += new TextDecoder().decode(bytes);
    },
  });
  return { code, stdout, stderr, stdoutBytes: Buffer.concat(chunks) };
}

const seen = new Map<string, number>();

/** Run with --json and hold the output to the contract. */
async function contract(...argv: string[]) {
  return contractWith({}, ...argv);
}

async function contractWith(
  options: Parameters<typeof run>[1],
  ...argv: string[]
) {
  const result = await run([...argv, "--json"], options);
  // One object, one line, nothing else on stdout; nothing on stderr.
  expect(result.stdout.endsWith("\n")).toBe(true);
  expect(result.stdout.trimEnd().split("\n")).toHaveLength(1);
  expect(result.stderr).toBe("");
  const envelope = JSON.parse(result.stdout);
  const problems = check(envelope, schema as Schema);
  expect(problems, `${argv.join(" ")}\n${result.stdout}`).toEqual([]);
  expect(envelope.ok).toBe(envelope.error === null);
  if (envelope.ok) expect(result.code).toBe(0);
  else expect([1, 2, 70]).toContain(result.code);
  // Partly done means something was done: there is data.
  if (result.code === 2) expect(envelope.data).not.toBeNull();
  const key = `${envelope.command}:${envelope.ok ? "ok" : "error"}`;
  seen.set(key, (seen.get(key) ?? 0) + 1);
  return { ...result, envelope, data: envelope.data, error: envelope.error };
}

const root = mkdtempSync(path.join(tmpdir(), "revise-contract-"));
let work: string;
let doc: string;

beforeAll(() => {
  work = mkdtempSync(path.join(root, "w-"));
  process.env.REVISE_WORKSPACE = path.join(work, ".revise");
  doc = path.join(work, "contract.docx");
  copyFileSync(CONTRACT, doc);
});

afterAll(() => {
  delete process.env.REVISE_WORKSPACE;
  rmSync(root, { recursive: true, force: true });
});

describe("--json contract", () => {
  it("is a valid JSON Schema: no keyword holds the wrong kind of value", () => {
    // A command named like a keyword ("examples") once sat in a container
    // validators read as a schema; standard validators then rejected the file.
    const arrayKeywords = new Set([
      "examples",
      "enum",
      "required",
      "allOf",
      "anyOf",
      "oneOf",
      "prefixItems",
    ]);
    const bad: string[] = [];
    const walk = (node: unknown, at: string) => {
      if (!node || typeof node !== "object") return;
      for (const [key, value] of Object.entries(node)) {
        const here = `${at}/${key}`;
        const isSchemaMap = [
          "$defs",
          "properties",
          "patternProperties",
          "dependentSchemas",
        ].includes(key);
        if (
          !Array.isArray(node) &&
          !isSchemaMap &&
          arrayKeywords.has(key) &&
          !Array.isArray(value)
        )
          bad.push(here);
        if (isSchemaMap && value && typeof value === "object") {
          for (const [name, sub] of Object.entries(value))
            walk(sub, `${here}/${name}`);
        } else walk(value, here);
      }
    };
    walk(schema, "#");
    expect(bad).toEqual([]);
  });

  it("holds for every command, in success and failure", async () => {
    // Reading.
    const status = await contract("status", doc);
    expect(status.data).toMatchObject({
      mode: "suggest",
      trackedChangeCount: 0,
    });
    expect(status.data.sourceFidelity.fidelity).toBe("exact");
    await contract("outline", doc, "--limit", "2");
    const hits = await contract("grep", doc, "Provider");
    const block = hits.data.blocks[0].blockId as string;
    await contract("read", doc, "--ids", block);
    await contract("read", doc, "3");
    await contract("search", doc, "Provider");
    expect((await contract("grep", doc, "[", "--regex")).error.code).toBe(
      "invalid_regex",
    );

    // Editing.
    const replaced = await contract(
      "replace-all",
      doc,
      "Provider",
      "Northstar",
    );
    expect(replaced.data.saved).toBe(true);
    expect(replaced.data.changeIds.length).toBeGreaterThan(0);
    const noop = await contract("replace-all", doc, "Customer", "Customer");
    expect(noop.code).toBe(0);
    expect(noop.data.saved).toBe(false);
    const edit = await contract(
      "call",
      doc,
      "replace",
      JSON.stringify({
        id: block,
        replacements: [{ find: "Cedar Vale", replace: "Cedar Hill" }],
      }),
    );
    expect(edit.data.saved, JSON.stringify(edit.data)).toBe(true);
    const refused = await contract(
      "call",
      doc,
      "replace",
      JSON.stringify({
        id: "nope",
        replacements: [{ find: "a", replace: "b" }],
      }),
    );
    expect(refused.code).toBe(1);
    expect(refused.error.code).toBe("tool_error");
    expect(refused.data.saved).toBe(false);
    await contract("tools");
    await contract("tools", "replace");
    expect((await contract("tools", "nope")).error.code).toBe("unknown_tool");
    expect((await contract("mode", doc)).data.changed).toBe(false);
    expect((await contract("mode", doc, "edit")).data.changed).toBe(true);
    await contract("mode", doc, "suggest");

    // Review.
    const changes = await contract("changes", doc);
    const ids = changes.data.changes.map(
      (c: { changeId: string }) => c.changeId,
    );
    expect(ids.length).toBeGreaterThan(1);
    const dry = await contract("accept", doc, "--all", "--dry-run");
    expect(dry.data).toMatchObject({
      dryRun: true,
      resolved: [],
      saved: false,
    });
    // Partly done: one accepted and saved, one unknown.
    const partial = await contract("accept", doc, ids[0], "no-such-change");
    expect(partial.code).toBe(2);
    expect(partial.data).toMatchObject({
      resolved: [ids[0]],
      notFound: ["no-such-change"],
      saved: true,
    });
    expect(partial.error.code).toBe("change_not_found");
    const none = await contract("reject", doc, "no-such-change");
    expect(none.code).toBe(1);
    expect(none.data.saved).toBe(false);
    await contract("diff", doc);
    await contract("diff", doc, "--since", "0");
    const history = await contract("history", doc);
    expect(history.data.checkpoint).toBe(
      history.data.checkpoints.at(-1).checkpoint,
    );
    const undone = await contract("undo", doc);
    expect(undone.data.discarded).toEqual([history.data.checkpoint]);
    await contract("undo", doc, "0");
    await contract("restore", doc, "1");
    expect((await contract("restore", doc, "99")).error.code).toBe(
      "checkpoint_not_found",
    );

    // Output.
    const exported = await contract(
      "export",
      doc,
      path.join(work, "copy.docx"),
    );
    expect(exported.data.docxExport.method).toBe("graft");
    expect(exported.data.summary).not.toBeNull();
    await contract("export", doc, path.join(work, "copy.md"), "--final");
    expect(
      (await contract("export", doc, path.join(work, "copy.docx"))).error.code,
    ).toBe("would_overwrite");

    // Create and convert.
    const draft = path.join(work, "draft.md");
    writeFileSync(draft, "# Memo\n\n- one\n- two\n");
    const created = await contract(
      "new",
      path.join(work, "memo.docx"),
      "--content",
      draft,
    );
    expect(created.data.mode).toBe("edit");
    await contract("convert", doc, path.join(work, "contract.txt"));
    const compared = await contract(
      "compare",
      CONTRACT,
      CONTRACT_V2,
      "--out",
      path.join(work, "redline.docx"),
    );
    expect(compared.data.verification.scope).toBe("normalized_text");

    // Raw package.
    const parts = await contract("xml", "ls", doc);
    expect(
      parts.data.parts.some(
        (p: { path: string }) => p.path === "word/document.xml",
      ),
    ).toBe(true);
    await contract("xml", "cat", doc, "word/styles.xml", "--match", "Heading1");
    expect(
      (await contract("xml", "cat", doc, "word/nope.xml")).error.code,
    ).toBe("no_such_part");

    // Everything else.
    await contract("skill");
    const examples = await contract("examples");
    expect(examples.data.examples.length).toBeGreaterThanOrEqual(5);
    await contract("help");
    await contract("help", "render");
    await contract("render", "--help");
    const version = await contract("version");
    expect(version.data.version).toBe(version.envelope.meta.version);
    expect((await contract("frobnicate")).error.code).toBe("unknown_command");
    expect((await contract("status", "--frob")).error.code).toBe("usage");
    expect(
      (await contract("status", path.join(work, "missing.docx"))).error.code,
    ).toBe("file_not_found");
    expect((await contract("accept")).error.hint).toBe("revise accept --help");

    // Every command was seen at least once.
    const commands = new Set([...seen.keys()].map((k) => k.split(":")[0]));
    for (const name of [
      "status",
      "outline",
      "read",
      "grep",
      "search",
      "replace-all",
      "call",
      "tools",
      "mode",
      "changes",
      "diff",
      "accept",
      "reject",
      "history",
      "undo",
      "restore",
      "export",
      "new",
      "compare",
      "convert",
      "xml ls",
      "xml cat",
      "skill",
      "examples",
      "help",
      "version",
    ]) {
      expect(commands, name).toContain(name);
    }
  }, 60_000);

  it("puts warnings in the envelope, and on stderr only in human output", async () => {
    const file = path.join(work, "outside.md");
    writeFileSync(file, "Alpha clause\n");
    await run(["status", file]);
    writeFileSync(file, "Inserted clause\n\nAlpha clause\n");
    const json = await contract("status", file);
    expect(json.envelope.warnings[0]).toMatchObject({
      code: "file_changed_outside",
      details: expect.objectContaining({ newBlockCount: 1 }),
    });
    writeFileSync(file, "Again\n");
    const human = await run(["status", file]);
    expect(human.stderr).toMatch(/^warning\[file_changed_outside\]: /);
  });

  it("indents on a terminal and stays one line in a pipe", async () => {
    const tty = await run(["version", "--json"], { tty: true });
    expect(tty.stdout.split("\n").length).toBeGreaterThan(3);
    expect(tty.stdout).not.toMatch(/\x1b\[/);
  });
});

describe("--json contract, round 2", () => {
  it("describes what export wrote, not the source", async () => {
    const file = path.join(work, "counts.docx");
    copyFileSync(CONTRACT, file);
    await run(["replace-all", file, "Cedar Vale", "Cedar Hill"]);
    const md = await contract("export", file, path.join(work, "counts.md"));
    expect(md.data).toMatchObject({ final: true, trackedChangeCount: 0 });
    expect(md.envelope.warnings.map((w: { code: string }) => w.code)).toEqual([
      "changes_accepted",
    ]);
    const out = path.join(work, "counts-out.docx");
    const docx = await contract("export", file, out);
    const reopened = await contract("status", out);
    expect(docx.data.trackedChangeCount).toBe(reopened.data.trackedChangeCount);
    expect(docx.data.summary.trackedChangeCount).toBe(
      docx.data.trackedChangeCount,
    );
  });

  it("names missing inputs file_not_found, everywhere", async () => {
    const missing = path.join(work, "missing input.docx");
    for (const argv of [
      ["convert", missing, path.join(work, "x.txt")],
      ["compare", missing, CONTRACT, "--out", path.join(work, "r.docx")],
      ["new", path.join(work, "n.docx"), "--from", path.join(work, "no.html")],
      ["new", path.join(work, "n.docx"), "--from", path.join(work, "no.md")],
      ["call", doc, "replace", `@${path.join(work, "no.json")}`],
    ]) {
      const result = await contract(...argv);
      expect(result.error.code, argv.join(" ")).toBe("file_not_found");
      expect(result.code).toBe(1);
    }
  });

  it("checks new's page setup before writing anything", async () => {
    const out = path.join(work, "bad-page.docx");
    for (const flags of [
      ["--page-size", "bogus"],
      ["--page", "bogus"],
      ["--orientation", "sideways"],
      ["--margins", "wide"],
    ]) {
      const result = await contract("new", out, ...flags);
      expect(result.error.code).toBe("usage");
      expect(existsSync(out)).toBe(false);
    }
  });

  it("gives xml failures a hint for this file", async () => {
    const patch = await contract(
      "xml",
      "patch",
      doc,
      "word/document.xml",
      "--find",
      "QQQQ-not-there",
      "--replace",
      "x",
    );
    expect(patch.error.code).toBe("find_not_found");
    expect(patch.error.hint).toBe(
      `revise xml cat ${doc} word/document.xml --match '<text>'`,
    );
    const part = await contract("xml", "cat", doc, "word/nope.xml");
    expect(part.error.hint).toBe(`revise xml ls ${doc}`);
  });

  it("says truncated only when there are more matches than shown", async () => {
    // A package whose document has exactly 20, then 21, occurrences.
    const zipWith = async (n: number) => {
      const zip = await JSZip.loadAsync(readFileSync(CONTRACT));
      zip.file(
        "word/notes.xml",
        `<notes>${Array.from({ length: n }, (_, i) => `<w:p><w:t>NEEDLE ${i}</w:t></w:p>`).join("")}</notes>`,
      );
      const file = path.join(work, `needles-${n}.docx`);
      writeFileSync(file, await zip.generateAsync({ type: "uint8array" }));
      return file;
    };
    const exactly = await contract(
      "xml",
      "cat",
      await zipWith(20),
      "word/notes.xml",
      "--match",
      "NEEDLE",
    );
    expect(exactly.data.fragments).toHaveLength(20);
    expect(exactly.data.truncated).toBe(false);
    const more = await contract(
      "xml",
      "cat",
      await zipWith(21),
      "word/notes.xml",
      "--match",
      "NEEDLE",
    );
    expect(more.data.fragments).toHaveLength(20);
    expect(more.data.truncated).toBe(true);
  });

  it("names every command the same way in help and in the envelope", async () => {
    expect((await contract("help", "version")).data.command.name).toBe(
      "version",
    );
    expect((await contract("help", "skill", "install")).data.command.name).toBe(
      "skill install",
    );
    const versionHelp = await contract("version", "--help");
    expect(versionHelp.envelope.command).toBe("help");
    const install = await contract(
      "skill",
      "install",
      "--dir",
      path.join(work, "skills dir"),
    );
    expect(install.envelope.command).toBe("skill install");
    expect(existsSync(install.data.output)).toBe(true);
    const md = path.join(work, "notes.markdown");
    writeFileSync(md, "# Notes\n");
    expect((await contract("status", md)).data.format).toBe("md");
  });

  it("installs the skill as revise-cli, replacing 3.x's revise/ folder", async () => {
    const base = path.join(work, "agent-skills");
    // What revise 3.0 and 3.1 installed, next to someone else's skill.
    mkdirSync(path.join(base, "revise", "reference"), { recursive: true });
    writeFileSync(
      path.join(base, "revise", "SKILL.md"),
      "---\nname: revise\ndescription: old\n---\n",
    );
    mkdirSync(path.join(base, "other"), { recursive: true });
    writeFileSync(
      path.join(base, "other", "SKILL.md"),
      "---\nname: other\n---\n",
    );
    const install = await contract("skill", "install", "--dir", base);
    expect(install.data.output).toBe(path.join(base, "revise-cli", "SKILL.md"));
    expect(readFileSync(install.data.output, "utf8")).toMatch(
      /^---\nname: revise-cli\n/,
    );
    expect(existsSync(path.join(base, "revise"))).toBe(false);
    expect(existsSync(path.join(base, "other", "SKILL.md"))).toBe(true);
    // A running agent session won't load it, so the output says to read it.
    const human = await run(["skill", "install", "--dir", base]);
    expect(human.stdout).toContain(
      `read ${path.join(base, "revise-cli", "SKILL.md")} now`,
    );
  });

  it("installs the skill where the agent running it reads skills", async () => {
    const cwd = process.cwd();
    const project = (name: string, folders: string[] = []) => {
      const dir = path.join(work, `skill-dir-${name}`);
      for (const folder of ["", ...folders]) {
        mkdirSync(path.join(dir, folder), { recursive: true });
      }
      return dir;
    };
    const installedIn = async (
      dir: string,
      env: Record<string, string> = {},
    ) => {
      process.chdir(dir);
      try {
        const result = await contractWith({ env }, "skill", "install");
        return path.relative(realpathSync(dir), result.data.output);
      } finally {
        process.chdir(cwd);
      }
    };
    const skill = (base: string) => path.join(base, "revise-cli", "SKILL.md");
    // Codex, Cursor, and the rest read the shared folder.
    expect(await installedIn(project("bare"))).toBe(skill(".agents/skills"));
    // Claude Code reads only its own, and says when it's the one running.
    expect(await installedIn(project("claude-code"), { CLAUDECODE: "1" })).toBe(
      skill(".claude/skills"),
    );
    // Otherwise the folders a project already has decide.
    expect(await installedIn(project("claude-project", [".claude"]))).toBe(
      skill(".claude/skills"),
    );
    expect(
      await installedIn(project("both", [".claude", ".agents/skills"])),
    ).toBe(skill(".agents/skills"));
  });

  it("keeps new's receipt when the cache can't be written", async () => {
    // Round 3: a plain filesystem error after the file exists.
    const blocked = path.join(work, "blocked-state");
    mkdirSync(blocked, { recursive: true });
    writeFileSync(path.join(blocked, "docs"), "not a directory");
    const saved = process.env.REVISE_WORKSPACE;
    process.env.REVISE_WORKSPACE = blocked;
    try {
      const out = path.join(work, "cache failure.docx");
      const result = await contract("new", out);
      expect(result.code).toBe(2);
      expect(result.error.code).toBe("io_error");
      expect(result.data).toMatchObject({
        output: realpathSync(out),
        outputWritten: true,
      });
      expect(existsSync(out)).toBe(true);
    } finally {
      process.env.REVISE_WORKSPACE = saved;
    }
  });

  it("labels every failure error[code], partial ones too", async () => {
    const file = path.join(work, "partial.docx");
    copyFileSync(CONTRACT, file);
    await run(["replace-all", file, "Cedar Vale", "Cedar Hill"]);
    const [id] = (await contract("changes", file)).data.changes.map(
      (c: { changeId: string }) => c.changeId,
    );
    const partial = await run(["accept", file, id, "missing"]);
    expect(partial.code).toBe(2);
    expect(partial.stderr).toMatch(/^error\[change_not_found\]: /);
  });
});

describe("human output", () => {
  it("prints errors on stderr with a hint, and data on stdout", async () => {
    const result = await run(["read", doc, "--ids", "nope"]);
    expect(result.code).toBe(1);
    expect(result.stdout).toBe("");
    expect(result.stderr).toMatch(/^error\[tool_error\]: /);
    const usage = await run(["accept", doc]);
    expect(usage.stderr).toContain("hint: revise accept --help");
  });

  it("shows each command's own help", async () => {
    const help = await run(["render", "--help"]);
    expect(help.stdout).toMatch(/^Usage: revise render <file>/);
    expect(help.stdout).toContain("--view <view>");
    expect((await run(["help", "render"])).stdout).toBe(help.stdout);
  });

  it("prints hints you can paste", async () => {
    const outline = await run(["outline", doc, "--limit", "1"]);
    expect(outline.stdout).toContain(
      `revise outline ${doc} --offset 1 --limit 1`,
    );
  });

  it("colors only a terminal, and honors NO_COLOR and --no-color", async () => {
    const piped = await run(["changes", doc]);
    expect(piped.stdout).not.toMatch(/\x1b\[/);
    const tty = await run(["changes", doc], { tty: true });
    expect(tty.stdout).toMatch(/\x1b\[36m/);
    for (const options of [
      { tty: true, env: { NO_COLOR: "1" } },
      { tty: true, env: { TERM: "dumb" } },
      { tty: true, env: { FORCE_COLOR: "0" } },
    ]) {
      expect((await run(["changes", doc], options)).stdout).not.toMatch(
        /\x1b\[/,
      );
    }
    expect(
      (await run(["changes", doc, "--no-color"], { tty: true })).stdout,
    ).not.toMatch(/\x1b\[/);
    expect(
      (await run(["changes", doc], { env: { FORCE_COLOR: "1" } })).stdout,
    ).toMatch(/\x1b\[/);
    // Color never replaces the words or signs that carry the meaning.
    const strip = (text: string) => text.replace(/\x1b\[\d+m/g, "");
    expect(strip(tty.stdout)).toBe(piped.stdout);
  });
});

it("ships the schema the tests hold the CLI to", () => {
  const text = readFileSync(
    fileURLToPath(new URL("../src/cli/output.schema.json", import.meta.url)),
    "utf8",
  );
  expect(JSON.parse(text).properties.schema.const).toBe(1);
});

describe("convert through stdin and stdout", () => {
  const zip = async (files: Record<string, string>) => {
    const z = new JSZip();
    for (const [name, text] of Object.entries(files)) z.file(name, text);
    return z.generateAsync({ type: "uint8array" });
  };

  it("recognizes each format from its bytes; text is Markdown", async () => {
    const enc = (text: string) => new TextEncoder().encode(text);
    const format = (bytes: Uint8Array) => sniffFormat(bytes).format;
    expect(format(readFileSync(CONTRACT))).toBe("docx");
    expect(
      format(
        await zip({
          mimetype: "application/vnd.oasis.opendocument.text",
          "content.xml": "<office:document-content/>",
        }),
      ),
    ).toBe("odt");
    expect(format(enc("{\\rtf1\\ansi Hello}"))).toBe("rtf");
    expect(format(enc("\uFEFF  <!DOCTYPE html><html><body>x"))).toBe("html");
    expect(format(enc("<p>Hello</p>"))).toBe("html");
    expect(format(enc("# Title\n\n- one\n"))).toBe("md");
    expect(format(enc("Just a note.\nSecond line."))).toBe("md");
    expect(() => format(enc("%PDF-1.7\n"))).toThrow(/PDF/);
    expect(() => format(new Uint8Array())).toThrow(/Nothing came in/);
    expect(() =>
      format(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 1])),
    ).toThrow(/binary/);
    const plainZip = await zip({ "a.txt": "x" });
    expect(() => format(plainZip)).toThrow(/not a .docx/);
  });

  it("gets the cases a first-bytes guess gets wrong", async () => {
    const enc = (text: string) => new TextEncoder().encode(text);
    const format = (bytes: Uint8Array) => sniffFormat(bytes).format;
    // HTML after a comment, as a fragment of inline tags, after whitespace.
    expect(
      format(enc("<!--generated--><html><body><p>Hi</p></body></html>")),
    ).toBe("html");
    expect(format(enc("<strong>Hello</strong> <em>world</em>"))).toBe("html");
    expect(format(enc(" ".repeat(9000) + "<p>Late</p>"))).toBe("html");
    expect(format(enc("<p>1. Scope</p>\n<p>2. Fees</p>"))).toBe("html");
    // A README's banner doesn't make its Markdown HTML.
    expect(
      format(
        enc(
          '<div align="center">Note</div>\n\n# Heading\n\n**Bold** and [a](https://x.y).\n',
        ),
      ),
    ).toBe("md");
    // A .docx whose document part comes after a large media entry.
    const original = await JSZip.loadAsync(readFileSync(CONTRACT));
    const late = new JSZip();
    late.file("word/media/big.bin", new Uint8Array(1_100_000).fill(7));
    for (const [name, entry] of Object.entries(original.files)) {
      if (!entry.dir) late.file(name, await entry.async("uint8array"));
    }
    expect(format(await late.generateAsync({ type: "uint8array" }))).toBe(
      "docx",
    );
    // UTF-16 text is text, re-encoded as UTF-8 for import.
    const utf16 = new Uint8Array([
      0xff,
      0xfe,
      ...Buffer.from("# Sample\nHi", "utf16le"),
    ]);
    const sniffed = sniffFormat(utf16);
    expect(sniffed.format).toBe("md");
    expect(new TextDecoder().decode(sniffed.bytes)).toBe("# Sample\nHi");
    // Round 2: weigh the whole text, not the first tag.
    expect(
      format(enc("<div>\n<p>\n1. Scope\n</p>\n<p>\n2. Fees\n</p>\n</div>")),
    ).toBe("html");
    expect(
      format(
        enc(
          "<strong>Note:</strong> **Important** and [docs](https://example.com).\n",
        ),
      ),
    ).toBe("md");
    expect(
      format(
        enc(
          '<div align="center">Project</div>\n\n*Important* project notes.\n',
        ),
      ),
    ).toBe("md");
    expect(format(enc("<p>snake_case_name and a * b</p>"))).toBe("html");
    // UTF-16 without a mark, and Windows-1252.
    const bare = sniffFormat(
      new Uint8Array(Buffer.from("# Café — 世界\nHi", "utf16le")),
    );
    expect(new TextDecoder().decode(bare.bytes)).toBe("# Café — 世界\nHi");
    const cp1252 = sniffFormat(
      new Uint8Array([0x43, 0x61, 0x66, 0xe9, 0x20, 0x80]),
    );
    expect(new TextDecoder().decode(cp1252.bytes)).toBe("Café €");
    // A truncated .docx is an input problem, not an internal error.
    const truncated = readFileSync(CONTRACT).subarray(0, 20000);
    expect(() => format(truncated)).toThrow(/damaged or incomplete/);
    const forced = await run(
      ["convert", "-", path.join(work, "t.md"), "--from", "docx"],
      { stdin: truncated },
    );
    expect(forced.code).toBe(1);
    expect(forced.stderr).toContain("error[import_failed]");
    // Hints run as written, or there is none.
    const empty = await run(["convert", "-", path.join(work, "e.md")], {
      stdin: "",
    });
    expect(empty.stderr).toMatch(/hint: cat notes\.md \| revise convert - /);
    const pdf = await run(["convert", "-", path.join(work, "p.md")], {
      stdin: "%PDF-1.4",
    });
    expect(pdf.code).toBe(1);
    expect(pdf.stderr).not.toContain("hint:");
  });

  it("reads a piped document and writes a file", async () => {
    const out = path.join(work, "piped.docx");
    const { data } = await contractWith(
      { stdin: "# Piped\n\nFrom **stdin**.\n" },
      "convert",
      "-",
      out,
    );
    expect(data).toMatchObject({
      input: null,
      inputFormat: "md",
      outputWritten: true,
      format: "docx",
    });
    const back = await run(["convert", out, "-"]);
    expect(back.code).toBe(0);
    expect(back.stdout).toContain("Piped");
    expect(back.stdout).toContain("stdin");
    // Human output names where it came from.
    const human = await run(
      ["convert", "-", path.join(work, "piped.html"), "--force"],
      { stdin: "Plain words." },
    );
    expect(human.stdout).toMatch(/from stdin \(md\)/);
  });

  it("writes the document, and only the document, to stdout", async () => {
    const md = await run(["convert", doc, "-"]);
    expect(md.code).toBe(0);
    expect(md.stderr).toBe("");
    expect(md.stdout).not.toMatch(/Wrote/);
    expect(md.stdout.length).toBeGreaterThan(100);
    const html = await run(["convert", doc, "-", "--to", "html"]);
    expect(html.stdout).toMatch(/<(p|h1|html)/i);
    const docx = await run(["convert", doc, "-", "--to", "docx"]);
    expect([...docx.stdoutBytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 3, 4]);
    // A docx piped through comes back out.
    const round = await run(["convert", "-", "-"], {
      stdin: docx.stdoutBytes,
    });
    expect(round.code).toBe(0);
    expect(round.stdout.length).toBeGreaterThan(100);
  });

  it("refuses what can't work, with a hint", async () => {
    const cases: Array<[string[], Parameters<typeof run>[1], string]> = [
      [["convert", doc, "-", "--to", "docx"], { tty: true }, "usage"],

      [["convert", doc, path.join(work, "x.md"), "--to", "html"], {}, "usage"],
      [["convert", doc, path.join(work, "y.md"), "--from", "md"], {}, "usage"],
      [["convert", "-", path.join(work, "z.docx")], {}, "usage"],
      [
        ["convert", "-", path.join(work, "z.docx"), "--from", "pages"],
        { stdin: "x" },
        "unsupported_format",
      ],
      [["convert", doc, "-", "--to", "odt"], {}, "unsupported_format"],
    ];
    for (const [argv, options, code] of cases) {
      const result = await run(argv, options);
      expect(result.code, argv.join(" ")).toBe(1);
      expect(result.stderr, argv.join(" ")).toContain(`error[${code}]`);
      expect(result.stderr, argv.join(" ")).toContain("hint:");
    }
    expect(existsSync(path.join(work, "z.docx"))).toBe(false);
    // --json owns stdout, so it can't share it with the document.
    const json = await contract("convert", doc, "-");
    expect(json.error.code).toBe("usage");
    expect(json.error.hint).toMatch(/--json$/);
    expect(json.stdout).not.toMatch(/Northstar|Provider/);
  });

  it("--from gets the same encoding handling", async () => {
    const out = path.join(work, "utf16-named.md");
    const result = await run(["convert", "-", out, "--from", "md", "--force"], {
      stdin: new Uint8Array(Buffer.from("# Café — 世界\n", "utf16le")),
    });
    expect(result.code).toBe(0);
    expect(readFileSync(out, "utf8")).toContain("Café — 世界");
  });

  it("--from overrides recognition", async () => {
    const out = path.join(work, "literal.docx");
    const { data } = await contractWith(
      { stdin: "# not a heading" },
      "convert",
      "-",
      out,
      "--from",
      "txt",
    );
    expect(data.inputFormat).toBe("txt");
    const back = await run(["convert", out, "-", "--to", "txt"]);
    expect(back.stdout).toContain("# not a heading");
  });
});

describe("--out: edit a copy", () => {
  it("saves the edit as a new file and leaves the original untouched", async () => {
    const original = path.join(work, "out-original.docx");
    copyFileSync(CONTRACT, original);
    const before = readFileSync(original);
    // Ids read from the original work on the copy.
    const read = await contract("grep", original, "Provider security lead");
    const id = read.data.blocks[0].blockId;
    const copy = path.join(work, "out-copy.docx");
    const called = await contract(
      "call",
      original,
      "replace",
      JSON.stringify({
        id,
        replacements: [{ find: "Provider", replace: "Northstar" }],
      }),
      "--out",
      copy,
    );
    expect(called.data.copiedFrom).toBe(realpathSync(original));
    expect(called.data.file).toBe(realpathSync(copy));
    expect(called.data.saved).toBe(true);
    expect(readFileSync(original).equals(before)).toBe(true);
    expect(
      (await contract("grep", copy, "Northstar security lead")).data.matchCount,
    ).toBe(1);
    expect(
      (await contract("grep", original, "Northstar security lead")).data
        .matchCount,
    ).toBe(0);
    // The copy's history starts at the original's state: undo returns to it.
    const history = await run(["history", copy]);
    expect(history.stdout).toMatch(/Copied from out-original\.docx/);
    await contract("undo", copy);
    expect(
      (await contract("grep", copy, "Provider security lead")).data.matchCount,
    ).toBe(1);
    // Without --out, copiedFrom is null.
    expect(
      (await contract("replace-all", copy, "Cedar Vale", "Cedar Hill")).data
        .copiedFrom,
    ).toBeNull();
    // Human output says what was copied where.
    const human = await run([
      "replace-all",
      original,
      "Customer",
      "Client",
      "--out",
      path.join(work, "out-human.docx"),
    ]);
    expect(human.stdout).toMatch(/Copied out-original\.docx to /);
    expect(readFileSync(original).equals(before)).toBe(true);
  });

  it("refuses what it can't do, and a failed edit leaves no copy", async () => {
    const original = path.join(work, "out-refuse.docx");
    copyFileSync(CONTRACT, original);
    const taken = path.join(work, "out-taken.docx");
    copyFileSync(CONTRACT, taken);
    expect(
      (await contract("replace-all", original, "Provider", "X", "--out", taken))
        .error.code,
    ).toBe("would_overwrite");
    expect(
      (
        await contract(
          "replace-all",
          original,
          "Provider",
          "X",
          "--out",
          taken,
          "--force",
        )
      ).data.saved,
    ).toBe(true);
    expect(
      (
        await contract(
          "replace-all",
          original,
          "Provider",
          "X",
          "--out",
          path.join(work, "x.md"),
        )
      ).error.code,
    ).toBe("usage");
    expect(
      (
        await contract(
          "accept",
          original,
          "--all",
          "--dry-run",
          "--out",
          path.join(work, "dry.docx"),
        )
      ).error.code,
    ).toBe("usage");
    const failed = path.join(work, "out-failed.docx");
    const result = await contract("accept", original, "nope", "--out", failed);
    expect(result.code).toBe(1);
    expect(existsSync(failed)).toBe(false);
    expect(
      (
        await contract(
          "replace-all",
          original,
          "Provider",
          "X",
          "--out",
          original,
        )
      ).error.code,
    ).toBe("same_file");
  });
});

describe("tool shorthand: help and usage are about the tool", () => {
  it("revise <tool>, with or without --help, shows the tool's input", async () => {
    for (const argv of [["leave_comment"], ["leave_comment", "--help"]]) {
      const help = await contract(...argv);
      expect(help.envelope.command).toBe("help");
      expect(help.data.command.name).toBe("leave_comment");
      expect(help.data.command.usage).toMatch(/^revise leave_comment <file>/);
      expect(help.data.inputSchema.type).toBe("object");
      const human = await run(argv);
      expect(human.code).toBe(0);
      expect(human.stdout).toMatch(/Usage: revise leave_comment <file>/);
      expect(human.stdout).toMatch(/Input \(a JSON object\)/);
      // One line per field, required ones first, with their types.
      expect(human.stdout).toMatch(/^ {2}comment \(string, required\)/m);
    }
  });

  it("saves a cell edit across commands after a read", async () => {
    const html = (await contract("read", doc)).data.html as string;
    const table = /<table id="([^"]+)"/.exec(html)?.[1];
    const rows = [...html.matchAll(/<tr id="([^"]+)"/g)].map((m) => m[1]);
    expect(table).toBeTruthy();
    expect(rows.length).toBeGreaterThan(1);
    const edited = await contract(
      "call",
      doc,
      "edit_table_cells",
      JSON.stringify({
        table_id: table,
        mode: "replace",
        cells: [{ row_id: rows[1], column: 0, text: "Cell marker" }],
      }),
    );
    expect(edited.envelope.ok).toBe(true);
    expect(edited.data.saved).toBe(true);
  });

  it("refuses editor-only layout settings instead of saving nothing", async () => {
    const refused = await contract(
      "call",
      doc,
      "set_page_layout",
      '{"pageless":true,"blockSpacing":"0.25in"}',
    );
    expect(refused.envelope.ok).toBe(false);
    expect(refused.error.code).toBe("usage");
    expect(refused.error.message).toMatch(
      /pageless, blockSpacing are a revise.io editor display setting/,
    );
  });

  it("refuses inherited names as references, and hints on 'requires'", async () => {
    for (const name of ["constructor", "toString", "__proto__"]) {
      expect((await contract("skill", "--reference", name)).error.code).toBe(
        "usage",
      );
    }
    const requires = await contract(
      "call",
      doc,
      "leave_comment",
      '{"comment":"x"}',
    );
    expect(requires.error.hint).toBe("revise leave_comment --help");
    const help = await run(["leave_comment", "--help"]);
    expect(help.stdout).toMatch(
      /anchor \(\{ text \} \| \{ end_text, start_text \} \| \{ whole_block: true \}\)/,
    );
    expect(help.stdout).toMatch(/^ {6}whole_block \(true, required\)/m);
  });

  it("names the input a tool needs when it is missing or wrong", async () => {
    const missing = await contract("call", doc, "insert_table_row");
    expect(missing.error.code).toBe("usage");
    expect(missing.error.message).toMatch(/tableId \(string, required\)/);
    expect(missing.error.hint).toBe("revise insert_table_row --help");
    const wrong = await contract(
      "call",
      doc,
      "insert_table_row",
      '{"tableId":"x","content":[]}',
    );
    expect(wrong.error.code).toBe("tool_error");
    expect(wrong.error.hint).toBe("revise insert_table_row --help");
  });
});

describe("--out never touches the original or what it replaces", () => {
  const historyOf = (file: string) => {
    const dir = path.join(work, ".revise", "docs");
    const name = readdirSync(dir).find((d) =>
      d.startsWith(path.basename(file).toLowerCase()),
    );
    return name
      ? readFileSync(path.join(dir, name, "meta.json"), "utf8")
      : null;
  };

  it("leaves the original's history alone, even when it has catching up to do", async () => {
    const original = path.join(work, "hist-original.docx");
    copyFileSync(CONTRACT, original);
    // First touch through --out: no history is created for the original.
    await contract(
      "replace-all",
      original,
      "Customer",
      "Client",
      "--out",
      path.join(work, "hist-a.docx"),
    );
    expect(historyOf(original)).toBeNull();
    // A tracked original, then changed outside revise: --out reconciles in
    // the copy, not in the original's history.
    await contract("replace-all", original, "Provider", "Northstar");
    const before = historyOf(original);
    copyFileSync(CONTRACT_V2, original);
    const copied = await contract(
      "replace-all",
      original,
      "Customer",
      "Client",
      "--out",
      path.join(work, "hist-b.docx"),
    );
    expect(copied.data.saved).toBe(true);
    expect(historyOf(original)).toBe(before);
  });

  it("puts back what --force replaced when the edit fails, and hints at the original", async () => {
    const original = path.join(work, "force-original.docx");
    copyFileSync(CONTRACT, original);
    const existing = path.join(work, "force-existing.docx");
    copyFileSync(CONTRACT, existing);
    await contract("replace-all", existing, "Customer", "ExistingClient");
    const bytes = readFileSync(existing);
    const history = historyOf(existing);
    const failed = await contract(
      "accept",
      original,
      "not-real",
      "--out",
      existing,
      "--force",
    );
    expect(failed.code).toBe(1);
    expect(readFileSync(existing).equals(bytes)).toBe(true);
    expect(historyOf(existing)).toBe(history);
    expect(failed.error.hint ?? "").not.toContain("force-existing");
    // Success with --force replaces it for good.
    const ok = await contract(
      "replace-all",
      original,
      "Provider",
      "Northstar",
      "--out",
      existing,
      "--force",
    );
    expect(ok.data.saved).toBe(true);
    expect(readFileSync(existing).equals(bytes)).toBe(false);
    expect(readdirSync(work).some((f) => f.includes("revise-replaced"))).toBe(
      false,
    );
  });
});

describe("--out edge cases", () => {
  it("refuses a folder, and a failed move-aside puts the file back", async () => {
    const original = path.join(work, "edge-original.docx");
    copyFileSync(CONTRACT, original);
    const folder = path.join(work, "folder.docx");
    mkdirSync(folder);
    writeFileSync(path.join(folder, "keep.txt"), "Keep me");
    const refused = await contract(
      "replace-all",
      original,
      "Customer",
      "X",
      "--out",
      folder,
      "--force",
    );
    expect(refused.error.code).toBe("usage");
    expect(readFileSync(path.join(folder, "keep.txt"), "utf8")).toBe("Keep me");
    if (process.getuid?.() === 0) return; // root ignores permissions
    const dir = mkdtempSync(path.join(work, "perm-"));
    const existing = path.join(dir, "existing.docx");
    copyFileSync(CONTRACT, existing);
    const workspace = process.env.REVISE_WORKSPACE;
    process.env.REVISE_WORKSPACE = path.join(dir, ".revise");
    try {
      await contract("replace-all", existing, "Customer", "ExistingClient");
      const bytes = readFileSync(existing);
      chmodSync(path.join(dir, ".revise", "docs"), 0o555);
      const failed = await contract(
        "accept",
        original,
        "not-real",
        "--out",
        existing,
        "--force",
      );
      chmodSync(path.join(dir, ".revise", "docs"), 0o755);
      expect(failed.code).toBe(1);
      expect(readFileSync(existing).equals(bytes)).toBe(true);
    } finally {
      process.env.REVISE_WORKSPACE = workspace;
    }
  });
});

describe("the agent skill's commands run", () => {
  /** Shell words, for the quoting the skill uses: '…', "…" with \\ escapes. */
  const words = (line: string): string[] => {
    const out: string[] = [];
    let current = "";
    let started = false;
    for (let i = 0; i < line.length; i += 1) {
      const c = line[i];
      if (c === "'") {
        const end = line.indexOf("'", i + 1);
        current += line.slice(i + 1, end);
        i = end;
        started = true;
      } else if (c === '"') {
        i += 1;
        while (line[i] !== '"') {
          if (line[i] === "\\") i += 1;
          current += line[i];
          i += 1;
        }
        started = true;
      } else if (c === " ") {
        if (started) out.push(current);
        current = "";
        started = false;
      } else {
        current += c;
        started = true;
      }
    }
    if (started) out.push(current);
    return out;
  };

  /** Every `revise …` the skill shows, in code blocks and inline code. */
  const commands = (markdown: string): string[] => {
    const found = new Set<string>();
    let fenced = false;
    for (const line of markdown.split("\n")) {
      if (line.trimStart().startsWith("```")) {
        fenced = !fenced;
        continue;
      }
      if (fenced && line.startsWith("revise ")) {
        found.add(line.replace(/\s{2,}#.*$/, "").trim());
      }
      for (const [, inline] of line.matchAll(/`(revise [^`]+)`/g))
        found.add(inline);
    }
    // Pipes, redirects, and heredocs need a shell; the pipe tests cover them.
    return [...found].filter((c) => !/[|<>]\s|\s-\s|\s-$|<<|\s>\s/.test(c));
  };

  it("never with arguments the CLI refuses", async () => {
    const shown = [SKILL_MARKDOWN, ...Object.values(SKILL_REFERENCES)].flatMap(
      commands,
    );
    expect(shown.length).toBeGreaterThan(20);
    const refused: string[] = [];
    // Outputs the examples name (out.pdf, redline.docx) land in the scratch
    // folder, not wherever the tests run.
    const cwd = process.cwd();
    process.chdir(work);
    try {
      for (const command of shown) {
        const file = path.join(
          work,
          `skill-${Math.random().toString(36).slice(2)}.docx`,
        );
        copyFileSync(CONTRACT, file);
        const argv = words(
          command
            .replace(/^revise /, "")
            .replace(/<file>|\bf\.docx\b|\bcontract\.docx\b|\bx\.docx\b/g, file)
            .replace(/--reference <name>/g, "--reference library")
            .replace(/<name>/g, "replace")
            .replace(/<tool>/g, "insert_table_row")
            .replace(/<index>/g, "0"),
        );
        const { stdout } = await run([...argv, "--json"]);
        const envelope = JSON.parse(stdout);
        if (
          ["usage", "unknown_command", "unknown_tool", "invalid_json"].includes(
            envelope.error?.code,
          )
        ) {
          refused.push(
            `${command}\n  -> ${envelope.error.code}: ${envelope.error.message}`,
          );
        }
      }
    } finally {
      process.chdir(cwd);
    }
    expect(refused, refused.join("\n")).toEqual([]);
  }, 180_000);
});
