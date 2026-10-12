// @vitest-environment node
// Ported from Revise's main repository on each release; changes land there.
// TODO(mirror): fixtures/table.rtf stands in for a file not exported: the main repository uses a third-party WordPad table (Apache Tika test file); this is a hand-written two-cell table.

import { exitWhenFlushed } from "../src/cli/exit";
import {
  chmodSync,
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { hostname, tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { afterAll, beforeEach, describe, expect, it } from "vitest";
import { main, writeDocumentFile } from "../src/cli/main";
import {
  readMeta,
  sessionNameFor,
  stageCommit,
  withFileLock,
  withFileLocks,
} from "../src/cli/store";

const fixture = (relative: string) =>
  fileURLToPath(new URL(`./fixtures/${relative}`, import.meta.url));
const CONTRACT = fixture("legal-negotiation-before.docx");
const CONTRACT_V2 = fixture(
  "legal-negotiation-after-clean.docx",
);
const RTF = fixture(
  "table.rtf",
);

const root = mkdtempSync(path.join(tmpdir(), "revise-cli-"));
let workdir: string;
let contract: string;

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
  return { ...result, envelope, json: envelope.data, error: envelope.error };
}

const part = async (file: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(file))).file(name)!.async("string");

describe("revise CLI (file-first)", () => {
  beforeEach(() => {
    workdir = mkdtempSync(path.join(root, "case-"));
    process.env.REVISE_WORKSPACE = path.join(workdir, ".revise");
    contract = path.join(workdir, "contract.docx");
    copyFileSync(CONTRACT, contract);
  });

  afterAll(() => {
    delete process.env.REVISE_WORKSPACE;
    rmSync(root, { recursive: true, force: true });
  });

  it("edits a file in place as signed Word revisions, and undo restores it", async () => {
    const hit = await reviseJson("grep", contract, "terminate without cause");
    expect(hit.code).toBe(0);
    const edit = await reviseJson(
      "call",
      contract,
      "replace",
      JSON.stringify({
        id: hit.json.blocks[0].blockId,
        replacements: [{ find: "thirty days", replace: "ninety days" }],
      }),
      "--author",
      "Claude",
    );
    expect(edit.code).toBe(0);
    expect(edit.json.checkpoint).toBe(1);
    const body = await part(contract, "word/document.xml");
    expect(body).toMatch(/<w:ins [^>]*w:author="Claude"/);
    expect(body).toMatch(/<w:del [^>]*w:author="Claude"/);
    expect((await reviseJson("diff", contract)).json.text).toContain(
      "[-thirty-]{+ninety+}",
    );
    expect((await revise("undo", contract)).code).toBe(0);
    expect(readFileSync(contract).equals(readFileSync(CONTRACT))).toBe(true);
  });

  it("re-reads a file that changed outside revise", async () => {
    await revise("grep", contract, "Provider");
    copyFileSync(CONTRACT_V2, contract);
    const after = await revise("grep", contract, "discounted");
    expect(after.code).toBe(0);
    expect(after.stderr).toMatch(/changed outside revise/);
    expect(after.stdout).toContain("discounted");
  });

  it("replaces every occurrence in one tracked step", async () => {
    const before = await reviseJson("grep", contract, "Provider");
    const total = before.json.blocks.reduce(
      (n: number, hit: { matchCount: number }) => n + hit.matchCount,
      0,
    );
    const replaced = await reviseJson(
      "replace-all",
      contract,
      "Provider",
      "Northstar",
    );
    expect(replaced.json.replacedOccurrences).toBe(total);
    expect(
      (await reviseJson("grep", contract, "Provider", "--view", "accepted"))
        .json.blocks,
    ).toEqual([]);
    // The tracked view (the default) still finds it, as deleted text.
    const tracked = await reviseJson("grep", contract, "Provider");
    expect(
      tracked.json.blocks.flatMap((hit: { matches: { change: string }[] }) =>
        hit.matches.map((m) => m.change),
      ),
    ).toEqual(Array(total).fill("deleted"));
    expect((await revise("reject", contract, "--all")).code).toBe(0);
    expect(
      (await reviseJson("grep", contract, "Provider")).json.blocks.length,
    ).toBe(before.json.blocks.length);
  });

  it("still takes 3.0's names: replace-text, new --from/--page, search --page", async () => {
    const old = await reviseJson(
      "replace-text",
      contract,
      "Provider",
      "Northstar",
    );
    expect(old.code).toBe(0);
    // Output keeps the name it was run under, so 3.0 scripts read it as before.
    expect(old.envelope.command).toBe("replace-text");
    expect(old.json.replacedOccurrences).toBeGreaterThan(0);
    expect((await revise("help", "replace-text")).stdout).toMatch(
      /replace-all/,
    );

    const draft = path.join(workdir, "draft.md");
    writeFileSync(draft, "# Memo\n\nBody.\n");
    const memo = path.join(workdir, "old-names.docx");
    const created = await reviseJson(
      "new",
      memo,
      "--from",
      draft,
      "--page",
      "a4",
    );
    expect(created.code).toBe(0);
    expect(created.json.summary.pageSize).toBe("a4");

    const searched = await reviseJson(
      "search",
      contract,
      "Northstar",
      "--page",
      "1",
    );
    expect(searched.code).toBe(0);
  });

  it("comments on a block read earlier, not on unread ones", async () => {
    const search = await reviseJson("search", contract, "Availability.");
    const read = search.json.result.matches[0].blockId as string;
    const comment = await reviseJson(
      "call",
      contract,
      "leave_comment",
      JSON.stringify({
        id: read,
        anchor: { text: "99.0%" },
        comment: "Ask for 99.9%.",
      }),
    );
    expect(comment.error).toBeNull();

    const other = path.join(workdir, "other.docx");
    copyFileSync(CONTRACT, other);
    const outline = await reviseJson("outline", other);
    const unread = outline.json.sections[3].startBlockId as string;
    const refused = await reviseJson(
      "call",
      other,
      "leave_comment",
      JSON.stringify({
        id: unread,
        anchor: { whole_block: true },
        comment: "x",
      }),
    );
    expect(refused.error?.details?.toolCode).toBe("out_of_context_edit");
  });

  it("keeps search_result_id usable across invocations", async () => {
    const first = await reviseJson("search", contract, "Provider");
    const id = first.json.result.search_result_id as string;
    await revise("search", contract, "Customer");
    const bulk = await reviseJson(
      "call",
      contract,
      "replace",
      JSON.stringify({
        search_result_id: id,
        replacements: [
          { find: "Provider", replace: "Northstar", occurrence: "all" },
        ],
      }),
    );
    expect(bulk.code).toBe(0);
    expect(bulk.json.result.effect.changed_target_count).toBe(
      first.json.result.totalMatches,
    );
  });

  it("reports tool rejections without a checkpoint or a write", async () => {
    const result = await reviseJson(
      "call",
      contract,
      "replace",
      JSON.stringify({
        id: "no-such-block",
        replacements: [{ find: "a", replace: "b" }],
      }),
    );
    // A refused call changed nothing: exit 1, with the receipt.
    expect(result.code).toBe(1);
    expect(result.error.code).toBe("tool_error");
    expect(result.json).toMatchObject({ tool: "replace", saved: false });
    expect(
      (await reviseJson("history", contract)).json.checkpoints.length,
    ).toBe(1);
    expect(readFileSync(contract).equals(readFileSync(CONTRACT))).toBe(true);
  });

  it("writes a redline file and reviews it by name", async () => {
    const revised = path.join(workdir, "revised.docx");
    copyFileSync(CONTRACT_V2, revised);
    const redline = path.join(workdir, "redline.docx");
    const compared = await reviseJson(
      "compare",
      contract,
      revised,
      "--out",
      redline,
      "--author",
      "Counterparty",
    );
    expect(compared.code).toBe(0);
    expect(compared.json.verification).toEqual({
      scope: "normalized_text",
      acceptAllMatchesRevised: true,
      rejectAllMatchesOriginal: true,
    });
    const changes = await reviseJson("changes", redline);
    expect(changes.json.changes.length).toBeGreaterThan(0);
    expect(
      changes.json.changes.every(
        (c: { author?: string }) => c.author === "Counterparty",
      ),
    ).toBe(true);
    expect((await revise("accept", redline, "--all")).code).toBe(0);
    const accepted = path.join(workdir, "accepted.md");
    const expected = path.join(workdir, "expected.md");
    await revise("export", redline, accepted);
    await revise("convert", revised, expected);
    expect(readFileSync(accepted, "utf8")).toBe(readFileSync(expected, "utf8"));
  });

  it("creates a document in one step and won't export over itself", async () => {
    const draft = path.join(workdir, "draft.html");
    writeFileSync(
      draft,
      '<p data-style-id="Title">Q4 Planning</p><h1>Goals</h1><ul><li>A</li><li>B</li></ul><table><tr><th>Item</th><th>Amount</th></tr><tr><td>x</td><td>1</td></tr></table>',
    );
    const memo = path.join(workdir, "memo.docx");
    const created = await reviseJson(
      "new",
      memo,
      "--content",
      draft,
      "--page-size",
      "a4",
      "--footer",
      "Page {PAGE} of {PAGES}",
    );
    expect(created.code).toBe(0);
    expect(created.json.mode).toBe("edit");
    expect(created.json.summary).toMatchObject({
      pageSize: "a4",
      title: "Q4 Planning",
      headingCounts: { H1: 1 },
      lists: [{ itemCount: 2 }],
      tables: [{ rowCount: 2, columnCount: 2 }],
      headersFooters: [
        expect.objectContaining({
          side: "footer",
          text: "Page {PAGE} of {PAGES}",
        }),
      ],
    });
    expect(await part(memo, "word/document.xml")).toContain(
      '<w:pStyle w:val="Title"/>',
    );
    expect((await reviseJson("export", memo, memo)).error.code).toBe(
      "same_file",
    );
  });

  it("converts without a cache", async () => {
    const out = path.join(workdir, "contract.md");
    expect((await revise("convert", contract, out)).code).toBe(0);
    expect(readFileSync(out, "utf8")).toContain("Software Services Agreement");
  });

  it("patches the package directly; the model re-learns it on next use", async () => {
    await revise("grep", contract, "Parties");
    // The model's state is the latest checkpoint's.
    const docDir = path.join(
      workdir,
      ".revise/docs",
      sessionNameFor(path.resolve(contract)),
    );
    const latestState = () => {
      const meta = JSON.parse(
        readFileSync(path.join(docDir, "meta.json"), "utf8"),
      ) as { checkpoints: Array<{ n: number }> };
      const n = meta.checkpoints.at(-1)!.n;
      return readFileSync(
        path.join(docDir, "checkpoints", `${String(n).padStart(4, "0")}.bin`),
      );
    };
    const stateBefore = latestState();
    const [style] = (
      await reviseJson(
        "xml",
        "cat",
        contract,
        "word/styles.xml",
        "--match",
        'w:styleId="Heading1"',
      )
    ).json.fragments as string[];
    const recolored = style.replace(
      /<w:color w:val="[0-9A-Fa-f]{6}"\/>/,
      '<w:color w:val="C00000"/>',
    );
    const withColor =
      recolored === style
        ? style.replace("<w:rPr>", '<w:rPr><w:color w:val="C00000"/>')
        : recolored;
    const patched = await reviseJson(
      "xml",
      "patch",
      contract,
      "word/styles.xml",
      "--find",
      style,
      "--replace",
      withColor,
    );
    expect(patched.code).toBe(0);
    // The file has the patch; the model was not touched by it.
    expect(await part(contract, "word/styles.xml")).toContain('w:val="C00000"');
    expect(await part(contract, "word/document.xml")).toBe(
      await part(CONTRACT, "word/document.xml"),
    );
    expect(latestState().equals(stateBefore)).toBe(true);
    // The next command re-learns it, folded into the patch's checkpoint.
    expect(
      (await reviseJson("history", contract)).json.checkpoints.length,
    ).toBe(2);
    expect(latestState().equals(stateBefore)).toBe(false);

    // The patch survives a tool edit that regenerates a heading.
    const heading = (
      await reviseJson("grep", contract, "Parties and interpretation")
    ).json.blocks[0].blockId as string;
    await revise(
      "call",
      contract,
      "replace",
      JSON.stringify({
        id: heading,
        replacements: [{ find: "Parties", replace: "The parties" }],
      }),
    );
    expect(await part(contract, "word/styles.xml")).toContain('w:val="C00000"');

    const shape = await reviseJson(
      "xml",
      "patch",
      contract,
      "word/document.xml",
      "--find",
      "<w:body>",
      "--replace",
      "<w:body><w:p/>",
    );
    expect(shape.error.code).toBe("body_shape_changed");
    const bad = await reviseJson(
      "xml",
      "patch",
      contract,
      "word/styles.xml",
      "--find",
      'w:styleId="Heading1"',
      "--replace",
      "<oops",
    );
    expect(bad.error.code).toBe("malformed_xml");

    await revise("undo", contract, "2");
    expect(readFileSync(contract).equals(readFileSync(CONTRACT))).toBe(true);
  });

  it("patches XML seen by xml cat even after tool edits", async () => {
    await revise("replace-all", contract, "Provider", "Northstar");
    const pendingDiff = (await reviseJson("diff", contract)).json.text;
    const units = (
      await reviseJson(
        "xml",
        "cat",
        contract,
        "word/document.xml",
        "--match",
        "Northstar",
      )
    ).json.fragments as string[];
    const paragraph = units.find(
      (u) => u.startsWith("<w:p") && u.includes("<w:pPr>"),
    )!;
    expect(paragraph).toBeDefined();
    const centered = paragraph.replace(
      "<w:pPr>",
      '<w:pPr><w:jc w:val="center"/>',
    );
    const patched = await reviseJson(
      "xml",
      "patch",
      contract,
      "word/document.xml",
      "--find",
      paragraph,
      "--replace",
      centered,
    );
    expect(patched.code).toBe(0);
    expect(await part(contract, "word/document.xml")).toContain(centered);
    // The model re-learns the file; the pending tracked changes survive
    // (re-read from Word's revisions, so the same edits with new ids).
    expect((await reviseJson("diff", contract)).json.text).toBe(pendingDiff);
    expect(
      (await reviseJson("grep", contract, "Northstar")).json.blocks.length,
    ).toBeGreaterThan(0);
  });

  it("reads formats it can't write, and won't edit them in place", async () => {
    const rtf = path.join(workdir, "table.rtf");
    copyFileSync(RTF, rtf);
    expect((await revise("read", rtf)).code).toBe(0);
    const edit = await reviseJson("replace-all", rtf, "a", "b");
    expect(edit.error.code).toBe("read_only_format");
  });

  it("settles by section correctly on a redline with deleted blocks", async () => {
    const revised = path.join(workdir, "revised.docx");
    copyFileSync(CONTRACT_V2, revised);
    const redline = path.join(workdir, "redline.docx");
    await revise("compare", contract, revised, "--out", redline);
    const before = (await reviseJson("changes", redline)).json
      .changes as Array<{
      changeId: string;
      section: string | null;
    }>;
    const inEight = before.filter((c) => c.section?.startsWith("8."));
    expect(inEight.length).toBeGreaterThan(0);

    const dry = await reviseJson(
      "accept",
      redline,
      "--except-section",
      "8",
      "--dry-run",
    );
    expect(dry.json.wouldResolve.length).toBe(before.length - inEight.length);
    expect((await reviseJson("history", redline)).json.checkpoints.length).toBe(
      1,
    );

    await revise("accept", redline, "--except-section", "8");
    const after = (await reviseJson("changes", redline)).json.changes as Array<{
      changeId: string;
      section: string | null;
    }>;
    expect(after.map((c) => c.changeId).sort()).toEqual(
      inEight.map((c) => c.changeId).sort(),
    );
    expect(after.every((c) => c.section?.startsWith("8."))).toBe(true);
  });

  it("treats an edit that changes nothing as success", async () => {
    const hit = await reviseJson("grep", contract, "terminate without cause");
    const noop = await reviseJson(
      "call",
      contract,
      "replace",
      JSON.stringify({
        id: hit.json.blocks[0].blockId,
        replacements: [{ find: "thirty days", replace: "thirty days" }],
      }),
    );
    expect([0, 2]).toContain(noop.code);
    expect(readFileSync(contract).equals(readFileSync(CONTRACT))).toBe(true);
  });

  it("serializes commands on one file", async () => {
    const { withFileLock } = await import("../src/cli/store");
    const events: string[] = [];
    const job = (label: string) =>
      withFileLock(path.resolve(contract), async () => {
        events.push(`${label}:start`);
        await new Promise((resolve) => setTimeout(resolve, 30));
        events.push(`${label}:end`);
      });
    await Promise.all([job("a"), job("b"), job("c")]);
    for (let i = 0; i < events.length; i += 2) {
      expect(events[i].replace(":start", "")).toBe(
        events[i + 1].replace(":end", ""),
      );
    }
  });

  describe("integrity", () => {
    it("records nothing when the save fails", async () => {
      const file = path.join(workdir, "denied.md");
      writeFileSync(file, "Alpha clause\n");
      expect((await revise("status", file)).code).toBe(0);
      chmodSync(file, 0o444);
      try {
        const failed = await reviseJson("replace-all", file, "Alpha", "Bravo");
        expect(failed.code).not.toBe(0);
        expect(failed.error.code).toBe("write_failed");
      } finally {
        chmodSync(file, 0o644);
      }
      expect(readFileSync(file, "utf8")).toContain("Alpha");
      expect((await reviseJson("grep", file, "Bravo")).json.blocks).toEqual([]);
      expect(
        (await reviseJson("grep", file, "Alpha")).json.blocks,
      ).toHaveLength(1);
      expect((await reviseJson("history", file)).json.checkpoints).toHaveLength(
        1,
      );
    });

    it("keeps model and file together through two xml patches and undo", async () => {
      const draft = path.join(workdir, "tiny.html");
      writeFileSync(draft, "<p>Alpha clause</p><p>Beta clause</p>");
      const doc = path.join(workdir, "patch.docx");
      expect((await revise("new", doc, "--from", draft)).code).toBe(0);
      const patch = (find: string, replace: string) =>
        revise(
          "xml",
          "patch",
          doc,
          "word/document.xml",
          "--find",
          find,
          "--replace",
          replace,
        );
      expect((await patch("Alpha clause", "Bravo clause")).code).toBe(0);
      expect((await patch("Bravo clause", "Delta clause")).code).toBe(0);
      expect((await revise("undo", doc)).code).toBe(0);
      const xml = await part(doc, "word/document.xml");
      expect(xml).toContain("Bravo clause");
      const texts = (await reviseJson("grep", doc, "clause")).json.blocks.map(
        (hit: { text: string }) => hit.text,
      );
      expect(texts).toEqual(["Bravo clause", "Beta clause"]);
      const stale = await reviseJson("replace-all", doc, "Alpha", "Omega");
      expect(stale.json.replacedOccurrences).toBe(0);
      expect(await part(doc, "word/document.xml")).not.toContain("Omega");
    });

    it("never moves an id to another paragraph when the file changes (Markdown)", async () => {
      const file = path.join(workdir, "ids.md");
      writeFileSync(file, "Alpha clause\n\nBeta clause\n");
      const [beta] = (await reviseJson("grep", file, "Beta")).json.blocks;
      writeFileSync(file, "Inserted clause\n\nAlpha clause\n\nBeta clause\n");
      const reread = await revise("grep", file, "clause");
      expect(reread.stderr).toContain("changed outside revise");
      const hits = (await reviseJson("grep", file, "clause")).json
        .blocks as Array<{
        blockId: string;
        text: string;
      }>;
      expect(hits.find((h) => h.blockId === beta.blockId)?.text).toBe(
        "Beta clause",
      );
      const rewrite = await reviseJson(
        "call",
        file,
        "replace_block",
        JSON.stringify({
          id: beta.blockId,
          content: ["<p>Overwritten Beta</p>"],
        }),
      );
      expect(rewrite.code).toBe(0);
      expect(readFileSync(file, "utf8")).toMatch(
        /Inserted clause\s+Alpha clause\s+Overwritten Beta/,
      );
    });

    it("follows Word's paragraph ids when a .docx is edited elsewhere", async () => {
      const draft = path.join(workdir, "ids.html");
      writeFileSync(draft, "<p>Alpha clause</p><p>Beta clause</p>");
      const doc = path.join(workdir, "ids.docx");
      expect((await revise("new", doc, "--from", draft)).code).toBe(0);
      const xml = await part(doc, "word/document.xml");
      expect(xml).toMatch(/<w:p [^>]*w14:paraId="[0-9A-F]{8}"/);
      const [beta] = (await reviseJson("grep", doc, "Beta")).json.blocks;
      // "Word": insert a paragraph at the top and edit Beta's text.
      const zip = await JSZip.loadAsync(readFileSync(doc));
      zip.file(
        "word/document.xml",
        xml
          .replace(
            /<w:body>/,
            "<w:body><w:p><w:r><w:t>Inserted clause</w:t></w:r></w:p>",
          )
          .replace("Beta clause", "Beta clause, amended"),
      );
      writeFileSync(doc, await zip.generateAsync({ type: "nodebuffer" }));
      const reread = await revise("status", doc);
      expect(reread.stderr).toMatch(/1 kept blocks? have new content/);
      const hits = (await reviseJson("grep", doc, "clause")).json
        .blocks as Array<{
        blockId: string;
        text: string;
      }>;
      expect(hits.map((h) => h.text)).toEqual([
        "Inserted clause",
        "Alpha clause",
        "Beta clause, amended",
      ]);
      expect(hits.find((h) => h.blockId === beta.blockId)?.text).toBe(
        "Beta clause, amended",
      );
    });

    it("re-reads a file another program saved between commands, losing nothing", async () => {
      const file = path.join(workdir, "race.md");
      writeFileSync(file, "Alpha clause\n");
      expect((await revise("status", file)).code).toBe(0);
      writeFileSync(file, "Alpha clause\n\nWritten by someone else\n");
      // The next command sees the change on load and re-reads it: nothing lost.
      expect((await revise("replace-all", file, "Alpha", "Bravo")).code).toBe(
        0,
      );
      expect(readFileSync(file, "utf8")).toContain("Written by someone else");
      expect(readFileSync(file, "utf8")).toContain("Bravo");
    });
  });

  it("refuses to overwrite a file another program saved during a command", async () => {
    const file = path.join(workdir, "race.txt");
    writeFileSync(file, "theirs\n");
    const { Doc } = await import("yjs");
    await expect(
      writeDocumentFile(
        {
          name: "race",
          file,
          fileHash: "hash-of-what-revise-read",
          documentId: "race",
          source: null,
          sourceFormat: "txt",
          createdAt: "",
          mode: "editing",
          checkpoints: [],
        },
        new Doc(),
        null,
        { bytes: new TextEncoder().encode("ours\n") },
      ),
    ).rejects.toMatchObject({ code: "file_changed" });
    expect(readFileSync(file, "utf8")).toBe("theirs\n");
  });

  describe("round 2", () => {
    it("refuses a saved search after the file changed", async () => {
      const file = path.join(workdir, "search.md");
      writeFileSync(file, "Alpha clause\n\nBeta clause\n");
      const found = await reviseJson("search", file, "clause");
      const id = found.json.result.search_result_id as string;
      writeFileSync(
        file,
        "Alpha clause\n\nBeta clause\n\nUnseen important clause\n",
      );
      const stale = await reviseJson(
        "call",
        file,
        "replace_block",
        JSON.stringify({
          search_result_id: id,
          content: ["<p>Overwritten</p>"],
        }),
      );
      expect(stale.code).not.toBe(0);
      expect(readFileSync(file, "utf8")).not.toContain("Overwritten");
    });

    it("puts the file back when its history can't be recorded", async () => {
      const file = path.join(workdir, "late.md");
      writeFileSync(file, "Alpha clause\n");
      expect((await revise("status", file)).code).toBe(0);
      const metaFile = path.join(
        workdir,
        ".revise/docs",
        sessionNameFor(file),
        "meta.json",
      );
      chmodSync(metaFile, 0o444);
      try {
        const failed = await reviseJson("replace-all", file, "Alpha", "Bravo");
        expect(failed.error.code).toBe("commit_failed");
      } finally {
        chmodSync(metaFile, 0o644);
      }
      expect(readFileSync(file, "utf8")).toBe("Alpha clause\n");
      expect((await reviseJson("history", file)).json.checkpoints).toHaveLength(
        1,
      );
      expect(
        (await reviseJson("grep", file, "Alpha")).json.blocks,
      ).toHaveLength(1);
    });

    it("settles a save interrupted after its commit, even after an outside edit", async () => {
      const file = path.join(workdir, "committed.md");
      writeFileSync(file, "Alpha clause\n");
      expect((await revise("replace-all", file, "Alpha", "Bravo")).code).toBe(
        0,
      );
      const name = sessionNameFor(file);
      const docDir = path.join(workdir, ".revise/docs", name);
      // "Crash" between committing the log and removing pending.json.
      const meta = readMeta(name);
      writeFileSync(
        path.join(docDir, "pending.json"),
        JSON.stringify({
          kind: "commit",
          tx: meta.lastTx,
          n: 1,
          command: "replace-all",
          summary: "",
          fileHash: meta.fileHash,
        }),
      );
      writeFileSync(file, "External clause\n");
      expect((await revise("status", file)).code).toBe(0);
      const history = (await reviseJson("history", file)).json
        .checkpoints as Array<{ checkpoint: number }>;
      expect(history.map((c) => c.checkpoint)).toEqual([0, 1, 2]);
      expect((await revise("undo", file)).code).toBe(0);
      expect(readFileSync(file, "utf8")).toContain("Bravo");
    });

    it("finishes a save that was interrupted after writing the file", async () => {
      const file = path.join(workdir, "crash.md");
      writeFileSync(file, "Alpha clause\n");
      expect((await revise("status", file)).code).toBe(0);
      const name = sessionNameFor(file);
      const meta = readMeta(name);
      // Stage a checkpoint, write the file, then "crash" before the commit.
      const written = "Bravo clause\n";
      const { createHash } = await import("node:crypto");
      stageCommit(
        meta,
        readFileSync(
          path.join(workdir, ".revise/docs", name, "checkpoints/0000.bin"),
        ),
        { command: "replace-all", summary: "interrupted" },
        createHash("sha256").update(written).digest("hex"),
      );
      writeFileSync(file, written);
      const next = await revise("history", file);
      expect(next.stderr).toMatch(/warning\[save_recovered\]: Finished saving/);
      expect((await reviseJson("history", file)).json.checkpoints).toHaveLength(
        2,
      );
    });

    it("takes several locks in one order, so opposite commands can't deadlock", async () => {
      const a = path.join(workdir, "a.md");
      const b = path.join(workdir, "b.md");
      writeFileSync(a, "a\n");
      writeFileSync(b, "b\n");
      const order: string[] = [];
      await Promise.all([
        withFileLocks([a, b], async () => {
          order.push("ab");
          await new Promise((r) => setTimeout(r, 30));
        }),
        withFileLocks([b, a], async () => {
          order.push("ba");
          await new Promise((r) => setTimeout(r, 30));
        }),
      ]);
      expect(order.sort()).toEqual(["ab", "ba"]);
    });

    it("parses positional numbers strictly", async () => {
      expect((await revise("restore", contract, "0garbage")).code).toBe(1);
      expect((await revise("read", contract, "3x")).code).toBe(1);
    });

    it("never replaces an output file without --force", async () => {
      const out = path.join(workdir, "out.md");
      writeFileSync(out, "OTHER WRITER MUST SURVIVE\n");
      const refused = await reviseJson("convert", contract, out);
      expect(refused.error.code).toBe("would_overwrite");
      expect(readFileSync(out, "utf8")).toBe("OTHER WRITER MUST SURVIVE\n");
      expect((await revise("convert", contract, out, "--force")).code).toBe(0);
    });
  });

  describe("locks", () => {
    it("ignores an exited owner, waits for live ones, and never evicts them", async () => {
      const file = path.join(workdir, "locked.md");
      writeFileSync(file, "x\n");
      const dir = path.join(workdir, ".revise", "locks", sessionNameFor(file));
      mkdirSync(dir, { recursive: true });
      const entry = (id: string, value: object) =>
        writeFileSync(path.join(dir, `${id}.entry`), JSON.stringify(value));
      // An exited process's entry, even with the lowest number: ignored.
      entry("a-dead", {
        pid: 2 ** 22 + 12345,
        host: hostname(),
        choosing: false,
        number: 1,
      });
      expect(await withFileLock(file, async () => "ran")).toBe("ran");
      // A live owner with a lower number holds the lock.
      entry("b-live", {
        pid: process.pid,
        host: hostname(),
        choosing: false,
        number: 1,
      });
      await expect(
        withFileLock(file, async () => "ran", { timeoutMs: 200 }),
      ).rejects.toMatchObject({ code: "file_busy" });
      // Giving up removed only the waiter's own entry.
      expect(readFileSync(path.join(dir, "b-live.entry"), "utf8")).toContain(
        '"number":1',
      );
      // A live owner still choosing its number (a slow starter) is waited
      // for too, whatever number it ends up with.
      rmSync(path.join(dir, "b-live.entry"));
      entry("c-choosing", {
        pid: process.pid,
        host: hostname(),
        choosing: true,
        number: 0,
      });
      await expect(
        withFileLock(file, async () => "ran", { timeoutMs: 200 }),
      ).rejects.toMatchObject({ code: "file_busy" });
      rmSync(path.join(dir, "c-choosing.entry"));
      // An entry that can't be read as one is treated as a live owner,
      // never as absence.
      writeFileSync(path.join(dir, "d-garbled.entry"), "{not json");
      await expect(
        withFileLock(file, async () => "ran", { timeoutMs: 200 }),
      ).rejects.toMatchObject({ code: "file_busy" });
      rmSync(path.join(dir, "d-garbled.entry"));
      expect(await withFileLock(file, async () => "ran")).toBe("ran");
    });

    it("lets one holder in at a time under contention", async () => {
      const file = path.join(workdir, "contended.md");
      writeFileSync(file, "x\n");
      let inside = 0;
      let most = 0;
      await Promise.all(
        Array.from({ length: 8 }, () =>
          withFileLock(file, async () => {
            inside += 1;
            most = Math.max(most, inside);
            await new Promise((r) => setTimeout(r, 15));
            inside -= 1;
          }),
        ),
      );
      expect(most).toBe(1);
    });

    it("tool shorthand takes the file's lock like every other edit", async () => {
      const file = path.join(workdir, "short.md");
      writeFileSync(file, "Alpha clause\n");
      const [hit] = (await reviseJson("grep", file, "Alpha")).json.blocks;
      const order: string[] = [];
      const held = withFileLock(file, async () => {
        order.push("held");
        await new Promise((r) => setTimeout(r, 300));
        order.push("released");
      });
      await new Promise((r) => setTimeout(r, 30));
      const edit = revise(
        "replace_block",
        file,
        JSON.stringify({ id: hit.blockId, content: ["<p>Bravo clause</p>"] }),
      ).then((result) => {
        order.push("edited");
        return result;
      });
      await held;
      expect((await edit).code).toBe(0);
      expect(order).toEqual(["held", "released", "edited"]);
    });
  });

  describe("honest output", () => {
    it("describes the exported file, not the source", async () => {
      const tracked = path.join(workdir, "tracked.docx");
      copyFileSync(
        fixture("legal-negotiation-after-tracked.docx"),
        tracked,
      );
      const final = path.join(workdir, "final.docx");
      const exported = await reviseJson("export", tracked, final, "--final");
      expect(exported.json.trackedChangeCount).toBe(0);
      expect(exported.json.summary.trackedChangeCount).toBe(0);
      expect((await reviseJson("status", final)).json.trackedChangeCount).toBe(
        0,
      );
      const again = await reviseJson("export", tracked, final);
      expect(again.error.code).toBe("would_overwrite");
    });

    it("rejects unknown options and partial results", async () => {
      expect(
        (await revise("grep", contract, "Provider", "--regexp")).code,
      ).toBe(1);
      const noop = await reviseJson(
        "replace-all",
        contract,
        "Provider",
        "Provider",
      );
      expect(noop.code).toBe(0);
      expect(noop.json.failures).toEqual([]);
      // Nothing asked for was pending: nothing done, exit 1.
      const missing = await reviseJson("accept", contract, "no-such-change");
      expect(missing.code).toBe(1);
      expect(missing.error.code).toBe("change_not_found");
      expect(missing.json.notFound).toEqual(["no-such-change"]);
      const sameOut = await reviseJson(
        "compare",
        contract,
        CONTRACT_V2,
        "--out",
        contract,
      );
      expect(sameOut.error.code).toBe("same_file");
    });
  });

  it("has no open step", async () => {
    const open = await reviseJson("open", contract);
    expect(open.error.message).toMatch(/name the file in each command/);
  });
});

describe("exit", () => {
  it("waits for buffered output before exiting", async () => {
    const order: string[] = [];
    let pending: (() => void) | null = null;
    // A pipe-like stream whose buffered data flushes later.
    const slow = {
      write: (_chunk: string, done?: () => void) => {
        pending = () => {
          order.push("flushed");
          done?.();
        };
        return false;
      },
    };
    const exited = exitWhenFlushed(0, [slow], () => order.push("exit"));
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(order).toEqual([]);
    (pending as unknown as () => void)();
    await exited;
    expect(order).toEqual(["flushed", "exit"]);
    process.exitCode = 0;
  });
});
