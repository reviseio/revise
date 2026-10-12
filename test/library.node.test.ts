// @vitest-environment node
// Ported from Revise's main repository on each release; changes land there.
// TODO(mirror): fixtures/image.png stands in for a file not exported: the main repository uses the Revise logo; this is a generated gradient.

import { existsSync, readFileSync } from "node:fs";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import * as Y from "yjs";
import {
  ReviseDocument,
  compare,
  convert,
  formatForPath,
  seedYDoc,
} from "../src/lib";
import { ydocToDocument, type ReviseJsonObject } from "@reviseio/engine";

// The library as a Node app uses it: everything in memory, bytes in and
// bytes out, nothing written to disk unless the caller writes it.

const fixture = (relative: string) =>
  fileURLToPath(new URL(`./fixtures/${relative}`, import.meta.url));
const CONTRACT = fixture("legal-negotiation-before.docx");
const PNG = fixture("image.png");
const CONTRACT_V2 = fixture(
  "legal-negotiation-after-clean.docx",
);

const part = async (bytes: Uint8Array, name: string) =>
  (await JSZip.loadAsync(bytes)).file(name)!.async("string");

// Word revisions in a saved .docx, and helpers to compare them.
const REVISION =
  /<w:(ins|del|rPrChange|pPrChange|trPrChange|tcPrChange|tblPrChange)\b([^>]*?)(\/?)>/g;
// "tag|id|author|date|text" for every revision element in the body.
const revisions = async (bytes: Uint8Array) =>
  revisionsInXml(await part(bytes, "word/document.xml"));
const revisionsInXml = (xml: string) => {
  const out: string[] = [];
  for (const match of xml.matchAll(REVISION)) {
    const [whole, tag, attrs, selfClosing] = match;
    const attr = (name: string) =>
      new RegExp(`w:${name}="([^"]*)"`).exec(attrs)?.[1] ?? "";
    let text = "";
    if (!selfClosing) {
      const start = match.index! + whole.length;
      let depth = 1;
      const tags = new RegExp(`<(/?)w:${tag}\\b[^>]*?(/?)>`, "g");
      tags.lastIndex = start;
      let end = xml.length;
      for (let m; (m = tags.exec(xml)); ) {
        if (m[2]) continue;
        depth += m[1] ? -1 : 1;
        if (depth === 0) {
          end = m.index;
          break;
        }
      }
      text = [
        ...xml
          .slice(start, end)
          .matchAll(/<w:(?:t|delText)\b[^>]*>([^<]*)<\/w:(?:t|delText)>/g),
      ]
        .map((t) => t[1])
        .join("");
    }
    out.push([tag, attr("id"), attr("author"), attr("date"), text].join("|"));
  }
  return out;
};
const minus = (a: string[], b: string[]) => {
  const rest = [...b];
  return a.filter((item) => {
    const at = rest.indexOf(item);
    if (at < 0) return true;
    rest.splice(at, 1);
    return false;
  });
};
const docxOf = async (doc: ReviseDocument) =>
  (await doc.export({ format: "docx" })).bytes;
const blockText = (node: any): string =>
  typeof node?.text === "string"
    ? node.text
    : (node?.children ?? []).map(blockText).join("");

/** A contract numbered the way Word numbers clauses: through the heading
 * styles, so no clause number is in the text. */
async function numberedContract(): Promise<Uint8Array> {
  const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
  const h = (style: string, text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const p = (text: string) =>
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const body = [
    h("Heading1", "Definitions"),
    p("The Provider defines terms."),
    h("Heading1", "Allocation of risk"),
    p("The Provider bears risk."),
    h("Heading2", "Cap on liability"),
    p("The Provider caps liability."),
    h("Heading1", "Termination"),
    p("The Provider may terminate."),
  ].join("");
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
  );
  zip.file(
    "word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>`,
  );
  zip.file(
    "word/numbering.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

/** A contract with the places replace-all must reach: a page header, a
 * footnote, curly quotes, a no-break space, and other capitalization. */
async function contractWithStories(): Promise<Uint8Array> {
  const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
  const p = (text: string) =>
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const body = [
    p("\u201cProvider\u201d means Acme Corp."),
    `<w:p><w:r><w:t xml:space="preserve">The Provider shall deliver.</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`,
    p("Payment is due net\u00a030 days."),
    p("PROVIDER will invoice monthly."),
  ].join("");
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`,
  );
  zip.file(
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W}>${p("Provider Confidential")}</w:hdr>`,
  );
  zip.file(
    "word/footnotes.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${p("Provider includes affiliates.")}</w:footnote></w:footnotes>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

describe("revise library", () => {
  it('exports "md" as markdown and names unknown formats', async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    try {
      const md = await doc.export({ format: "md" as never });
      expect(new TextDecoder().decode(md.bytes)).toMatch(/\S/);
      await expect(doc.export({ format: "rtfx" as never })).rejects.toThrow(
        /Unknown export format "rtfx"/,
      );
    } finally {
      doc.dispose();
    }
  }, 30_000);

  it("set_title saves the title into the file's properties", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    try {
      const result = await doc.tools.run("set_title", {
        title: "Master Services Agreement",
      });
      expect(result.ok).toBe(true);
      const { bytes } = await doc.export({ format: "docx" });
      expect(await part(bytes, "docProps/core.xml")).toContain(
        "<dc:title>Master Services Agreement</dc:title>",
      );
      const html = await doc.export({ format: "html" });
      expect(new TextDecoder().decode(html.bytes)).toContain(
        "<title>Master Services Agreement</title>",
      );
      const reopened = await ReviseDocument.open({
        bytes,
        filename: "renamed.docx",
      });
      // A fresh open titles its exports from the saved property too.
      const reopenedHtml = await reopened.export({ format: "html" });
      expect(new TextDecoder().decode(reopenedHtml.bytes)).toContain(
        "<title>Master Services Agreement</title>",
      );
      const again = await reopened.export({ format: "docx" });
      expect(await part(again.bytes, "docProps/core.xml")).toContain(
        "<dc:title>Master Services Agreement</dc:title>",
      );
      reopened.dispose();
    } finally {
      doc.dispose();
    }
  }, 30_000);

  it("edits a Word file as tracked changes and exports it, writing nothing", async () => {
    const original = new Uint8Array(readFileSync(CONTRACT));
    const doc = await ReviseDocument.open(
      { bytes: original, filename: "contract.docx" },
      { author: "Legal bot" },
    );
    try {
      expect(doc.mode).toBe("suggesting");
      const hits = doc.grep("Provider");
      expect(hits.length).toBeGreaterThan(0);

      const before = doc.version();
      const result = await doc.replaceAll("Provider", "Northstar");
      expect(result.replaced).toBeGreaterThan(0);
      expect(result.failures).toEqual([]);
      expect(Buffer.from(doc.version()).equals(Buffer.from(before))).toBe(
        false,
      );

      const changes = await doc.changes();
      expect(changes.length).toBeGreaterThan(0);
      expect(changes.every((c) => c.author === "Legal bot")).toBe(true);

      const exported = await doc.export({ format: "docx" });
      expect(exported.fidelity?.method).toBe("graft");
      expect(exported.pendingTrackedChanges).toBeGreaterThan(0);
      const out = exported.bytes;
      const xml = await part(out, "word/document.xml");
      expect(xml).toContain('w:author="Legal bot"');
      // Untouched parts carry over from the original package.
      expect(await part(out, "word/numbering.xml")).toBe(
        await part(original, "word/numbering.xml"),
      );
      expect(existsSync(".revise")).toBe(false);
    } finally {
      doc.dispose();
    }
  });

  it("keeps a table's untouched cells byte for byte when one cell is edited", async () => {
    // The bench's bulk rename: "Provider" also sits in three cells of the
    // security table. Only paragraphs that contain it may change.
    const original = new Uint8Array(readFileSync(CONTRACT));
    const doc = await ReviseDocument.open(CONTRACT, { author: "Claude" });
    await doc.replaceAll("Provider", "Northstar");
    const { bytes } = await doc.export({ format: "docx" });
    doc.dispose();
    const xml = async (data: Uint8Array) =>
      (await JSZip.loadAsync(data)).file("word/document.xml")!.async("string");
    const paragraphs = (text: string) =>
      text.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];
    const kept = new Set(paragraphs(await xml(bytes)));
    const changed = paragraphs(await xml(original)).filter((p) => !kept.has(p));
    expect(changed.length).toBeGreaterThan(0);
    for (const p of changed) expect(p).toMatch(/Provider/);
  });

  it("settles tracked changes by section, with a preview first", async () => {
    const { document, verified } = await compare(CONTRACT, CONTRACT_V2, {
      author: "Counsel",
    });
    try {
      expect(verified).toEqual({
        acceptAllGivesRevised: true,
        rejectAllGivesOriginal: true,
      });
      const all = await document.changes();
      const inEight = await document.changes({ section: "8" });
      expect(inEight.length).toBeGreaterThan(0);

      const preview = await document.preview({ exceptSection: "8" });
      expect(preview.length).toBe(all.length - inEight.length);
      expect((await document.changes()).length).toBe(all.length);

      await document.accept({ exceptSection: "8" });
      const left = await document.changes();
      expect(left.map((c) => c.id).sort()).toEqual(
        inEight.map((c) => c.id).sort(),
      );
      const diff = await document.diff();
      expect(diff.changed).toBeGreaterThan(0);
    } finally {
      document.dispose();
    }
  });

  it("round-trips through a snapshot, the hand-off format", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    await doc.replaceAll("Provider", "Northstar");
    const snapshot = await doc.snapshot();
    const pending = (await doc.changes()).length;
    doc.dispose();

    const restored = await ReviseDocument.restore(snapshot);
    try {
      expect((await restored.changes()).length).toBe(pending);
      expect(restored.grep("Northstar").length).toBeGreaterThan(0);
    } finally {
      restored.dispose();
    }
  });

  it("creates documents from HTML and converts between formats", async () => {
    const doc = await ReviseDocument.create({
      html: '<p data-style-id="Title">Q4 Planning</p><h1>Goals</h1><ul><li>Ship the CLI</li></ul>',
    });
    try {
      const summary = await doc.summary();
      expect(summary.data.title).toBe("Q4 Planning");
      expect(summary.data.headings).toEqual({ H1: 1 });
      expect(summary.data.lists).toEqual([1]);
      const markdown = new TextDecoder().decode(
        (await doc.export({ format: "markdown" })).bytes,
      );
      expect(markdown).toContain("Ship the CLI");
    } finally {
      doc.dispose();
    }

    const md = new TextDecoder().decode(await convert(CONTRACT, "markdown"));
    expect(md).toContain("Provider");
    expect(formatForPath("out.PDF")).toBe("pdf");
  });

  it("re-reads a changed file and keeps block ids", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    try {
      const [first] = doc.grep("Provider");
      const reloaded = await doc.reload(
        new Uint8Array(readFileSync(CONTRACT)),
        "contract.docx",
      );
      expect(reloaded.idsNew).toBe(0);
      expect(doc.grep("Provider")[0].id).toBe(first.id);
    } finally {
      doc.dispose();
    }
  });

  it("runs agent tools and reports expected failures as values", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    try {
      expect(
        doc.tools.definitions().some((tool) => tool.name === "replace"),
      ).toBe(true);
      const read = await doc.tools.run("read_blocks_from_index", { index: 0 });
      expect(read.ok).toBe(true);
      const missing = await doc.tools.run("replace", {
        id: "no-such-block",
        replacements: [{ find: "a", replace: "b" }],
      });
      expect(missing.ok).toBe(false);
      const unknown = await doc.tools.run("no_such_tool");
      expect(unknown).toMatchObject({
        ok: false,
        error: { code: "unknown_tool" },
      });
    } finally {
      doc.dispose();
    }
  });

  it("embeds images inserted by URL, without calling Revise's servers", async () => {
    const png = readFileSync(PNG);
    const requests: string[] = [];
    const server = createServer((request, response) => {
      requests.push(request.url ?? "");
      if (request.url === "/logo.png") {
        response.writeHead(200, { "content-type": "image/png" });
        response.end(png);
      } else {
        response.writeHead(200, { "content-type": "text/html" });
        response.end("<html>not an image</html>");
      }
    });
    await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
    const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
    const doc = await ReviseDocument.open(CONTRACT, {
      mode: "editing",
      // The test server is on loopback; real integrations keep the default.
      images: { allowPrivateNetwork: true },
    });
    try {
      const [anchor] = doc.grep("Provider");
      const inserted = await doc.tools.run("insert_block", {
        referenceId: anchor.id,
        position: "after",
        content: [`<img src="${base}/logo.png">`],
      });
      expect(inserted.ok).toBe(true);
      expect(requests).toEqual(["/logo.png"]);

      const out = (await doc.export({ format: "docx" })).bytes;
      const zip = await JSZip.loadAsync(out);
      const media = Object.values(zip.files).filter(
        (entry) => !entry.dir && entry.name.startsWith("word/media/"),
      );
      const embedded = await Promise.all(
        media.map((entry) => entry.async("nodebuffer")),
      );
      expect(embedded.some((bytes) => bytes.equals(png))).toBe(true);

      const notImage = await doc.tools.run("insert_block", {
        referenceId: anchor.id,
        position: "after",
        content: [`<img src="${base}/page.html">`],
      });
      expect(notImage.ok).toBe(false);
    } finally {
      doc.dispose();
      server.close();
    }
  });

  it("snapshots carry the Word package, mode, and author", async () => {
    const doc = await ReviseDocument.open(CONTRACT, { author: "Legal bot" });
    await doc.replaceAll("Provider", "Northstar");
    const snapshot = await doc.snapshot();
    doc.dispose();
    const restored = await ReviseDocument.restore(snapshot);
    try {
      expect(restored.mode).toBe("suggesting");
      expect(restored.author).toBe("Legal bot");
      expect(restored.package).not.toBeNull();
      const out = await restored.export({ format: "docx" });
      expect(out.fidelity?.method).toBe("graft");
    } finally {
      restored.dispose();
    }
    await expect(
      ReviseDocument.restore(snapshot.subarray(0, 40)),
    ).rejects.toMatchObject({ code: "snapshot_corrupt" });
  });

  it("finds sections by Word's automatic clause numbers, subsections included", async () => {
    const doc = await ReviseDocument.open({
      bytes: await numberedContract(),
      filename: "numbered.docx",
    });
    try {
      await doc.replaceAll("Provider", "Northstar");
      const all = await doc.changes();
      expect(all.map((c) => c.section)).toEqual([
        "1. Definitions",
        "2. Allocation of risk",
        "2.1 Cap on liability",
        "3. Termination",
      ]);
      const two = await doc.changes({ section: "2" });
      expect(two.map((c) => c.section)).toEqual([
        "2. Allocation of risk",
        "2.1 Cap on liability",
      ]);
      expect(
        (await doc.changes({ section: "Allocation of risk" })).length,
      ).toBe(2);
      expect((await doc.changes({ section: "2.1" })).length).toBe(1);
      expect((await doc.preview({ exceptSection: "2" })).length).toBe(2);
      await expect(doc.changes({ section: "9" })).rejects.toMatchObject({
        code: "section_not_found",
      });
      // A full title picks one heading; text in several headings is refused.
      expect((await doc.changes({ section: "termination" })).length).toBe(1);
      await expect(doc.changes({ section: "i" })).rejects.toMatchObject({
        code: "section_ambiguous",
      });
    } finally {
      doc.dispose();
    }
  });

  it("edits one paragraph without restyling it or its neighbors", async () => {
    const original = await numberedContract();
    const doc = await ReviseDocument.open(
      { bytes: original, filename: "numbered.docx" },
      { mode: "editing" },
    );
    try {
      await doc.replaceText("caps liability", "caps aggregate liability");
      const { bytes } = await doc.export({ format: "docx" });
      const before = await part(original, "word/document.xml");
      const after = await part(bytes, "word/document.xml");
      const paragraphs = (xml: string) =>
        xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];
      // Every paragraph but the edited one is byte-identical.
      const kept = new Set(paragraphs(after));
      const changed = paragraphs(before).filter((p) => !kept.has(p));
      expect(changed).toHaveLength(1);
      expect(changed[0]).toContain("caps liability");
      // The edited one gains no font or size its text didn't have.
      const edited = paragraphs(after).find((p) =>
        p.includes("caps aggregate liability"),
      )!;
      expect(edited).not.toMatch(/<w:rFonts|<w:sz\b/);
    } finally {
      doc.dispose();
    }
  });

  it("compares a one-word edit as an inline change", async () => {
    const original = await numberedContract();
    const doc = await ReviseDocument.open(
      { bytes: original, filename: "numbered.docx" },
      { mode: "editing" },
    );
    let edited: Uint8Array;
    try {
      await doc.replaceText("caps liability", "caps aggregate liability");
      edited = (await doc.export({ format: "docx" })).bytes;
    } finally {
      doc.dispose();
    }
    const { document, stats } = await compare(
      { bytes: original, filename: "a.docx" },
      { bytes: edited, filename: "b.docx" },
    );
    try {
      expect(stats).toMatchObject({
        inlineChanges: 1,
        insertedBlocks: 0,
        deletedBlocks: 0,
      });
    } finally {
      document.dispose();
    }
  });

  it("selects tracked changes by block ids and ranges", async () => {
    const doc = await ReviseDocument.open({
      bytes: await numberedContract(),
      filename: "numbered.docx",
    });
    try {
      await doc.replaceAll("Provider", "Northstar");
      await doc.replaceAll("Definitions", "Defined terms");
      const all = await doc.changes();
      const blockOf = (text: string) =>
        all.find((c) => c.inserted.includes(text) || c.deleted.includes(text))!;
      const body = all.filter((c) => c.inserted === "Northstar");
      // A top-level paragraph id.
      expect(
        (await doc.changes({ blocks: [body[1].blockId!] })).map((c) => c.id),
      ).toEqual([body[1].id]);
      // A range, first..last, in document order.
      const range = await doc.changes({
        blocks: [{ from: body[1].blockId!, to: body[2].blockId! }],
      });
      expect(range.map((c) => c.section)).toEqual([
        "2. Allocation of risk",
        "2.1 Cap on liability",
      ]);
      // A numbered heading: its list's id selects the change in it.
      const heading = blockOf("Defined terms");
      expect(
        (await doc.changes({ blocks: [heading.blockId!] })).map((c) => c.id),
      ).toEqual([heading.id]);
      await expect(
        doc.changes({
          blocks: [{ from: body[2].blockId!, to: body[1].blockId! }],
        }),
      ).rejects.toMatchObject({ code: "invalid_selection" });
      await expect(doc.changes({ blocks: ["nope"] })).rejects.toMatchObject({
        code: "block_not_found",
      });
      const result = await doc.reject({ blocks: [body[0].blockId!] });
      expect(result.resolved).toEqual([body[0].id]);
    } finally {
      doc.dispose();
    }
  });

  it("replaces everywhere: headers, footnotes, and the body", async () => {
    const doc = await ReviseDocument.open({
      bytes: await contractWithStories(),
      filename: "stories.docx",
    });
    try {
      const result = await doc.replaceAll("Provider", "Supplier");
      expect(result.regions).toEqual({ body: 2, header: 1, footnote: 1 });
      // Review lists them where they are, with their text.
      const where = (await doc.changes()).map((c) => [
        c.section ?? null,
        c.deleted,
        c.inserted,
      ]);
      expect(where).toContainEqual(["Header", "Provider", "Supplier"]);
      expect(where).toContainEqual(["Footnote 1", "Provider", "Supplier"]);
      expect(result.failures).toEqual([]);
      expect(result.nearMisses).toEqual([
        { text: "PROVIDER", count: 1, reason: "case" },
      ]);
      const { bytes } = await doc.export({ format: "docx" });
      const header = await part(bytes, "word/header1.xml");
      const notes = await part(bytes, "word/footnotes.xml");
      for (const xml of [header, notes]) {
        expect(xml).toMatch(/<w:delText[^>]*>Provider<\/w:delText>/);
        expect(xml).toMatch(/<w:ins\b[\s\S]*?Supplier/);
      }
      // Accepting all leaves no Provider anywhere but the capitalized one.
      await doc.accept({ all: true });
      const accepted = await doc.export({ format: "docx" });
      expect(await part(accepted.bytes, "word/header1.xml")).not.toMatch(
        /Provider/,
      );
    } finally {
      doc.dispose();
    }
  });

  it("matches the way Word's Find does: quotes, spaces, and case", async () => {
    const doc = await ReviseDocument.open(
      { bytes: await contractWithStories(), filename: "stories.docx" },
      { mode: "editing" },
    );
    try {
      // Straight quotes typed, curly in the document: matched, and the
      // replacement takes the curly style.
      const quoted = await doc.replaceAll('"Provider"', '"Supplier"');
      expect(quoted.replaced).toBe(1);
      expect(doc.grep("\u201cSupplier\u201d").length).toBe(1);
      // A plain space matches the no-break space.
      expect((await doc.replaceAll("net 30", "net 45")).replaced).toBe(1);
      // --exact: characters must match.
      const exact = await doc.replaceAll("net 45", "net 60", { exact: true });
      expect(exact.replaced).toBe(0);
      expect(exact.nearMisses).toEqual([
        { text: "net\u00a045", count: 1, reason: "characters" },
      ]);
      // --ignore-case.
      const anyCase = await doc.replaceAll("provider", "Vendor", {
        ignoreCase: true,
      });
      expect(anyCase.replaced).toBe(4);
      // The 3.0 name still works.
      expect((await doc.replaceText("Vendor", "Seller")).replaced).toBe(4);
    } finally {
      doc.dispose();
    }
  });

  it("guards its lifecycle and its inputs", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    await expect(doc.accept({ all: true, section: "8" })).rejects.toMatchObject(
      { code: "invalid_selection" },
    );
    // Operations queue: concurrent calls run one at a time, in order.
    const [renamed, changes] = await Promise.all([
      doc.replaceAll("Provider", "Northstar"),
      doc.changes(),
    ]);
    expect(changes.length).toBeGreaterThan(0);
    expect(renamed.replaced).toBeGreaterThan(0);
    doc.dispose();
    await expect(doc.changes()).rejects.toMatchObject({ code: "disposed" });

    const shared = await ReviseDocument.open(CONTRACT);
    const borrowed = await ReviseDocument.fromYDoc(shared.ydoc);
    await expect(
      borrowed.reload(new Uint8Array(readFileSync(CONTRACT)), "c.docx"),
    ).rejects.toMatchObject({ code: "borrowed_ydoc" });
    borrowed.dispose();
    expect(shared.grep("Provider").length).toBeGreaterThan(0);
    shared.dispose();
  });

  it("seeds a collaboration room once, and exports it from the room's original", async () => {
    // A sync server seeds its room on every connection; only the first
    // call writes anything.
    const room = new Y.Doc();
    expect(await seedYDoc(room, CONTRACT)).toBe(true);
    expect(await seedYDoc(room, CONTRACT)).toBe(false);

    // Two peers that both seeded an empty room before syncing converge on
    // one copy.
    const other = new Y.Doc();
    expect(
      await seedYDoc(other, {
        bytes: new Uint8Array(readFileSync(CONTRACT)),
        filename: "contract.docx",
      }),
    ).toBe(true);
    Y.applyUpdate(room, Y.encodeStateAsUpdate(other));
    expect(await seedYDoc(room, CONTRACT)).toBe(false);
    const solo = await ReviseDocument.open(CONTRACT);
    const doc = await ReviseDocument.fromYDoc(room);
    try {
      expect(doc.grep("Provider").length).toBe(solo.grep("Provider").length);
      // The room carries the original Word file, so a session over the
      // room keeps untouched content byte for byte without being handed
      // the package again.
      expect(doc.package).not.toBeNull();
      const exported = await doc.export({ format: "docx" });
      expect(exported.fidelity?.method).not.toBe("regenerated");
    } finally {
      doc.dispose();
      solo.dispose();
    }
  });

  it("refuses private-network image URLs by default", async () => {
    const doc = await ReviseDocument.open(CONTRACT, { mode: "editing" });
    try {
      const [anchor] = doc.grep("Provider");
      const blocked = await doc.tools.run("insert_block", {
        referenceId: anchor.id,
        position: "after",
        content: ['<img src="http://127.0.0.1:9/logo.png">'],
      });
      expect(blocked.ok).toBe(false);
      expect(JSON.stringify(blocked)).toMatch(/private or local address/);
    } finally {
      doc.dispose();
    }
  });

  it("reports what the .docx export actually did", async () => {
    const tracked = fixture(
      "legal-negotiation-after-tracked.docx",
    );
    const doc = await ReviseDocument.open(tracked);
    try {
      const final = await doc.export({ format: "docx", mode: "final" });
      expect(final.pendingTrackedChanges).toBe(0);
      expect(final.fidelity?.method).toBe("graft");
      // Accepting every change rewrites the changed paragraphs.
      expect(final.fidelity?.regeneratedBlocks).toBeGreaterThan(0);
    } finally {
      doc.dispose();
    }
    // A bare Y.js state has no original Word file to carry over.
    const opened = await ReviseDocument.open(CONTRACT);
    const bare = await ReviseDocument.restore(
      Y.encodeStateAsUpdate(opened.ydoc),
    );
    opened.dispose();
    try {
      const out = await bare.export({ format: "docx" });
      expect(out.fidelity?.method).toBe("regenerated");
      expect(out.warnings.join(" ")).toMatch(/original Word file/);
    } finally {
      bare.dispose();
    }
  });

  it("orders snapshots after queued edits", async () => {
    const doc = await ReviseDocument.open(CONTRACT);
    const edit = doc.replaceAll("Provider", "Northstar");
    const snapshot = doc.snapshot();
    await edit;
    const restored = await ReviseDocument.restore(await snapshot);
    try {
      expect(restored.grep("Northstar").length).toBeGreaterThan(0);
    } finally {
      restored.dispose();
      doc.dispose();
    }
  });

  it("never exports the original file when paragraphs were removed", async () => {
    // A Word file with one wholly inserted paragraph, then plain Tail.
    const original = await ReviseDocument.create({
      html: "<p>Inserted</p><p>Tail</p>",
    });
    const base = (await original.export({ format: "docx" })).bytes;
    original.dispose();
    const zip = await JSZip.loadAsync(base);
    const xml = await zip.file("word/document.xml")!.async("string");
    zip.file(
      "word/document.xml",
      xml.replace(
        /<w:p\b([^>]*)>((?:(?!<\/w:p>).)*?)<w:r>((?:(?!<\/w:p>).)*?Inserted(?:(?!<\/w:p>).)*?)<\/w:r><\/w:p>/s,
        (_m, attrs, pre, run) =>
          `<w:p${attrs}>${pre.replace(/<w:pPr>/, '<w:pPr><w:rPr><w:ins w:id="7" w:author="R"/></w:rPr>') || '<w:pPr><w:rPr><w:ins w:id="7" w:author="R"/></w:rPr></w:pPr>'}<w:ins w:id="8" w:author="R"><w:r>${run}</w:r></w:ins></w:p>`,
      ),
    );
    const redline = new Uint8Array(
      await zip.generateAsync({ type: "uint8array" }),
    );
    const doc = await ReviseDocument.open({
      bytes: redline,
      filename: "r.docx",
    });
    try {
      expect((await doc.changes()).length).toBeGreaterThan(0);
      await doc.reject({ all: true });
      expect(doc.grep("Inserted")).toEqual([]);
      const out = await doc.export({ format: "docx" });
      expect(Buffer.from(out.bytes).equals(Buffer.from(redline))).toBe(false);
      const reread = await ReviseDocument.open({
        bytes: out.bytes,
        filename: "o.docx",
      });
      try {
        expect(reread.grep("Inserted")).toEqual([]);
        expect(reread.grep("Tail").length).toBe(1);
        expect(await reread.changes()).toEqual([]);
      } finally {
        reread.dispose();
      }
    } finally {
      doc.dispose();
    }
  });

  describe("a replace in one table cell leaves the other cells' tracked changes alone", () => {
    // Section 7's security table holds 8 pending Word tracked changes in
    // cells of several rows. Editing one cell paragraph must not accept,
    // drop, or re-author any of them.
    const REDLINE = fixture(
      "legal-negotiation-after-tracked.docx",
    );
    const TARGET = "Provider security lead";

    // Block ids are minted per open, so compare blocks in document order.
    const texts = (doc: ReviseDocument) =>
      doc.grep(/[\s\S]/).map((hit) => ({ id: hit.id, text: hit.text }));

    const baselineAfterRejectAll = async () => {
      const doc = await ReviseDocument.open(REDLINE, { author: "Claude" });
      try {
        await doc.reject({ all: true });
        return texts(doc).map((block) => block.text);
      } finally {
        doc.dispose();
      }
    };

    for (const mode of ["direct", "tracked"] as const) {
      it(`in ${mode} mode`, async () => {
        const original = await baselineAfterRejectAll();
        const doc = await ReviseDocument.open(REDLINE, { author: "Claude" });
        try {
          const before = doc.session.listSuggestions();
          expect(before.length).toBe(170);
          const [hit] = doc.grep(TARGET);
          expect(hit).toBeDefined();
          const inTable = (await doc.changes()).filter(
            (change) => change.blockIndex === hit.blockIndex,
          );
          expect(inTable.length).toBe(8);

          const result = await doc.tools.call(
            "replace",
            {
              id: hit.id,
              replacements: [
                { find: TARGET, replace: "Security review owner" },
              ],
            },
            mode === "direct" ? { directMode: true } : undefined,
          );
          expect(result.ok).toBe(true);
          expect(doc.grep("Security review owner").map((h) => h.id)).toEqual([
            hit.id,
          ]);

          // Every change that was pending before is still pending, with its
          // id, author, label and date.
          const after = doc.session.listSuggestions();
          const afterById = new Map(after.map((s) => [s.id, s]));
          for (const change of before) {
            expect(afterById.get(change.id)).toEqual(change);
          }
          const added = (await doc.changes()).filter(
            (change) => !before.some((b) => b.id === change.id),
          );
          if (mode === "direct") {
            expect(added).toEqual([]);
          } else {
            // Only the edit itself.
            expect(added).toEqual([
              expect.objectContaining({
                author: "Claude",
                blockIndex: hit.blockIndex,
                deleted: TARGET,
                inserted: "Security review owner",
              }),
            ]);
          }

          // Rejecting everything still restores the original text of every
          // other block, the other cells of the table included.
          const targetIndex = texts(doc).findIndex((b) => b.id === hit.id);
          await doc.reject({ all: true });
          const restored = texts(doc).map((block) => block.text);
          expect(restored.length).toBe(original.length);
          restored.forEach((text, index) => {
            if (index === targetIndex && mode === "direct") {
              expect(text).toBe(
                original[index].replace(TARGET, "Security review owner"),
              );
            } else {
              expect(text, `block ${index}`).toBe(original[index]);
            }
          });
        } finally {
          doc.dispose();
        }
      });
    }
  });

  describe("table structure, row styling and edits inside a revised paragraph keep the security table's tracked changes", () => {
    // Section 7's security table holds 8 pending Word changes by "Revise
    // Agent" in its requirement column. Adding or removing rows and columns
    // that hold none of them, styling a row, or editing a paragraph that
    // already holds some must leave the others as they were: same type,
    // author, date and text, and the same w:id. Reject-all must restore the
    // original text, in the session and after the saved .docx is reopened.
    const REDLINE = fixture(
      "legal-negotiation-after-tracked.docx",
    );
    const PHRASES = [
      "72 hours after confirmed harm",
      "5 days; no objection right",
      "Indefinite retention",
    ];
    // The visible text of each cell of the security table, row by row.
    const securityTable = async (bytes: Uint8Array) => {
      const xml = await part(bytes, "word/document.xml");
      const at = ["Provider security lead", "Security review owner"]
        .map((anchor) => xml.indexOf(anchor))
        .find((index) => index >= 0)!;
      const start = xml.lastIndexOf("<w:tbl>", at);
      const table = xml.slice(start, xml.indexOf("</w:tbl>", at));
      return [...table.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((row) =>
        [...row[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cell) =>
          [...cell[0].matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
            .map((t) => t[1])
            .join(""),
        ),
      );
    };
    const ids = (doc: ReviseDocument) => {
      const root = ydocToDocument(doc.ydoc);
      const table = root.children.find(
        (block: any) =>
          block.type === "table" &&
          blockText(block).includes("Provider security lead"),
      ) as any;
      const paragraphs = table.children.flatMap((row: any) =>
        row.children.flatMap((cell: any) => cell.children),
      );
      return {
        tableId: table.id as string,
        rowIds: table.children.map((row: any) => row.id as string),
        // The requirement paragraph of the incident row: "Within 24 hours…"
        pendingParagraphId: paragraphs.find((p: any) =>
          blockText(p).includes("confirmed harm"),
        ).id as string,
        ownerParagraphId: paragraphs.find((p: any) =>
          blockText(p).includes("Provider security lead"),
        ).id as string,
      };
    };

    type Case = {
      tool: string;
      input: (found: ReturnType<typeof ids>) => ReviseJsonObject;
      // The w:ids this call may legitimately rewrite (it edits their text).
      touches?: string[];
    };
    const cases: Record<string, Case> = {
      insert_table_row: {
        tool: "insert_table_row",
        input: ({ tableId }) => ({
          tableId,
          referenceIndex: 0,
          position: "before",
          content: ["New A", "New B", "New C"],
        }),
      },
      remove_table_rows: {
        tool: "remove_table_rows",
        input: ({ tableId, rowIds }) => ({ tableId, rowIds: [rowIds[0]] }),
      },
      insert_table_column: {
        tool: "insert_table_column",
        input: ({ tableId }) => ({
          tableId,
          referenceIndex: 0,
          position: "before",
          content: ["New A", "New B", "New C", "New D"],
        }),
      },
      remove_table_column: {
        tool: "remove_table_column",
        input: ({ tableId }) => ({ tableId, columnIndex: 0 }),
      },
      "style_blocks on a row": {
        tool: "style_blocks",
        input: ({ rowIds }) => ({
          selectors: `#${rowIds[1]}`,
          attrs: [{ name: "backgroundColor", value: "yellow" }],
        }),
      },
      // A cell paragraph with no pending changes of its own: its new
      // revisions must carry the caller's author (they were "Anonymous").
      "replace_block on a cell paragraph": {
        tool: "replace_block",
        input: ({ ownerParagraphId }) => ({
          id: ownerParagraphId,
          content: ["<p>Security review owner</p>"],
        }),
      },
      "replace inside a revised paragraph": {
        tool: "replace",
        input: ({ pendingParagraphId }) => ({
          id: pendingParagraphId,
          replacements: [{ find: "24", replace: "12" }],
        }),
        // "Within 24 " is itself a pending insertion: the edit lands in it.
        touches: ["100086"],
      },
    };

    const unrelated = (tuples: string[], touches: string[] = []) =>
      tuples.filter((t) => !touches.includes(t.split("|")[1]));

    for (const mode of ["direct", "tracked"] as const) {
      for (const [label, spec] of Object.entries(cases)) {
        it(`${label} (${mode})`, async () => {
          const doc = await ReviseDocument.open(REDLINE, { author: "Claude" });
          try {
            const original = await revisions(await docxOf(doc));
            const securityRevisions = original.filter((t) =>
              /\|Revise Agent\|2026-09-14/.test(t),
            );
            expect(securityRevisions.length).toBeGreaterThan(8);
            const suggestionsBefore = doc.session.listSuggestions();

            const result = await doc.tools.run(
              spec.tool,
              spec.input(ids(doc)),
              { direct: mode === "direct" },
            );
            expect(result.ok, JSON.stringify(result)).toBe(true);

            // Every unrelated suggestion is still pending, untouched.
            const afterById = new Map(
              doc.session.listSuggestions().map((s) => [s.id, s]),
            );
            for (const s of suggestionsBefore) {
              if (spec.touches?.some((id) => s.id.endsWith(id))) continue;
              expect(afterById.get(s.id), s.id).toEqual(s);
            }

            const saved = await docxOf(doc);
            const now = await revisions(saved);
            expect(
              minus(unrelated(original, spec.touches), now),
              "revisions lost",
            ).toEqual([]);
            const added = minus(now, original);
            if (mode === "direct") {
              // Nothing new is tracked; an edit inside a pending insertion
              // only changes that insertion's text.
              for (const t of added) {
                expect(spec.touches ?? [], t).toContain(t.split("|")[1]);
              }
            } else {
              // An edit inside a pending insertion splits it around this
              // call's deletion; the pieces after the first take new w:ids.
              const touched = original
                .filter((t) => spec.touches?.includes(t.split("|")[1]))
                .map((t) => t.split("|").slice(2, 4).join("|"));
              for (const t of unrelated(added, spec.touches)) {
                if (touched.includes(t.split("|").slice(2, 4).join("|")))
                  continue;
                expect(t.split("|")[2], t).toBe("Claude");
              }
            }

            // Reject everything in the session…
            const pristine = await ReviseDocument.open(REDLINE);
            await pristine.reject({ all: true });
            const originalTable = await securityTable(await docxOf(pristine));
            pristine.dispose();
            for (const phrase of PHRASES)
              expect(originalTable.flat()).toContain(phrase);

            const reopened = await ReviseDocument.open({
              bytes: saved,
              filename: "saved.docx",
            });
            try {
              await doc.reject({ all: true });
              const inSession = await securityTable(await docxOf(doc));
              // …and in the saved file, opened fresh.
              await reopened.reject({ all: true });
              const cold = await securityTable(await docxOf(reopened));
              for (const [where, table] of [
                ["in session", inSession],
                ["after reopen", cold],
              ] as const) {
                for (const phrase of PHRASES)
                  expect(table.flat(), `${where}: ${phrase}`).toContain(phrase);
                // Tracked: rejecting undoes this call too.
                if (mode === "tracked")
                  expect(table, where).toEqual(originalTable);
              }
            } finally {
              reopened.dispose();
            }
          } finally {
            doc.dispose();
          }
        }, 60_000);
      }
    }
  });

  describe("a block edit on or beside a revised paragraph leaves its tracked changes alone", () => {
    // Clause 1.2 holds 8 pending Word changes by "Revise Agent"; clauses 1.1
    // and 1.3 around it hold none. Inserting, removing, splitting or joining
    // a neighbor, and styling, clearing, appending to, splitting, joining or
    // editing clause 1.2 itself, or inserting next to it, must leave clause
    // 1.2's changes as they were (same w:id, author, date and text). Direct
    // mode applies only the new edit; reject-all restores the original text,
    // in the session and after the saved .docx is reopened.
    const REDLINE = fixture(
      "legal-negotiation-after-tracked.docx",
    );
    const REVISED = "1.2 Subscription Fees.";
    // The visible text of the body paragraph that starts with `prefix`.
    const paragraph = async (bytes: Uint8Array, prefix: string) => {
      const xml = await part(bytes, "word/document.xml");
      for (const p of xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)) {
        const text = [...p[0].matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
          .map((t) => t[1])
          .join("");
        if (text.startsWith(prefix)) return text;
      }
      return null;
    };
    const neighbors = (doc: ReviseDocument) => {
      const blocks = ydocToDocument(doc.ydoc).children as any[];
      const at = blocks.findIndex((block) =>
        blockText(block).startsWith(REVISED),
      );
      expect(at).toBeGreaterThan(0);
      return {
        previous: blocks[at - 1].id as string,
        revised: blocks[at].id as string,
        next: blocks[at + 1].id as string,
      };
    };
    type Case = {
      tool: string;
      input: (found: ReturnType<typeof neighbors>) => ReviseJsonObject;
      // How a direct edit changes the revised paragraph's own text, which
      // reject-all then keeps: joined to a neighbor, text appended, split in
      // two, or one word replaced.
      changes?: "merged" | "appended" | "split" | "replaced";
    };
    const cases: Record<string, Case> = {
      "insert_block after the previous paragraph": {
        tool: "insert_block",
        input: ({ previous }) => ({
          referenceId: previous,
          position: "after",
          content: ["<p>Editor owner paragraph.</p>"],
        }),
      },
      "insert_block before the next paragraph": {
        tool: "insert_block",
        input: ({ next }) => ({
          referenceId: next,
          position: "before",
          content: ["<p>Editor owner paragraph.</p>"],
        }),
      },
      "remove_blocks on the previous paragraph": {
        tool: "remove_blocks",
        input: ({ previous }) => ({ id: previous }),
      },
      "remove_blocks on the next paragraph": {
        tool: "remove_blocks",
        input: ({ next }) => ({ id: next }),
      },
      "break_paragraph on the previous paragraph": {
        tool: "break_paragraph",
        input: ({ previous }) => ({ id: previous, substring: "between" }),
      },
      "break_paragraph on the next paragraph": {
        tool: "break_paragraph",
        input: ({ next }) => ({ id: next, substring: "Schedule B" }),
      },
      "join_paragraphs of the previous paragraph and the revised one": {
        tool: "join_paragraphs",
        input: ({ previous, revised }) => ({
          firstId: previous,
          secondId: revised,
        }),
        changes: "merged",
      },
      "join_paragraphs of the revised paragraph and the next one": {
        tool: "join_paragraphs",
        input: ({ revised, next }) => ({ firstId: revised, secondId: next }),
        changes: "merged",
      },
      "insert_block after the revised paragraph": {
        tool: "insert_block",
        input: ({ revised }) => ({
          referenceId: revised,
          position: "after",
          content: ["<p>Editor owner paragraph.</p>"],
        }),
      },
      "insert_block before the revised paragraph": {
        tool: "insert_block",
        input: ({ revised }) => ({
          referenceId: revised,
          position: "before",
          content: ["<p>Editor owner paragraph.</p>"],
        }),
      },
      "style_blocks text-align on the revised paragraph": {
        tool: "style_blocks",
        input: ({ revised }) => ({
          selectors: `#${revised}`,
          attrs: [{ name: "text-align", value: "center" }],
        }),
      },
      "clear_formatting on the revised paragraph": {
        tool: "clear_formatting",
        input: ({ revised }) => ({ id: revised }),
      },
      "append_to_paragraph on the revised paragraph": {
        tool: "append_to_paragraph",
        input: ({ revised }) => ({ id: revised, content: " Appended." }),
        changes: "appended",
      },
      "break_paragraph on the revised paragraph": {
        tool: "break_paragraph",
        input: ({ revised }) => ({ id: revised, substring: "An Authorized" }),
        changes: "split",
      },
      "replace in the revised paragraph, beside its changes": {
        tool: "replace",
        input: ({ revised }) => ({
          id: revised,
          replacements: [{ find: "individual", replace: "named person" }],
        }),
        changes: "replaced",
      },
    };

    for (const mode of ["direct", "tracked"] as const) {
      for (const [label, spec] of Object.entries(cases)) {
        it(`${label} (${mode})`, async () => {
          const pristine = await ReviseDocument.open(REDLINE);
          await pristine.reject({ all: true });
          const originalText = await paragraph(await docxOf(pristine), REVISED);
          pristine.dispose();
          expect(originalText).toBeTruthy();

          const doc = await ReviseDocument.open(REDLINE, { author: "X" });
          try {
            const original = await revisions(await docxOf(doc));
            const suggestionsBefore = doc.session.listSuggestions();
            expect(suggestionsBefore.length).toBe(170);
            // As an agent would: read the start of the document first.
            const read = await doc.tools.run("read_blocks_from_index", {
              start_index: 0,
            });
            expect(read.ok, JSON.stringify(read)).toBe(true);
            const result = await doc.tools.run(
              spec.tool,
              spec.input(neighbors(doc)),
              { direct: mode === "direct" },
            );
            expect(result.ok, JSON.stringify(result)).toBe(true);

            const afterById = new Map(
              doc.session.listSuggestions().map((s) => [s.id, s]),
            );
            for (const s of suggestionsBefore)
              expect(afterById.get(s.id), s.id).toEqual(s);

            const saved = await docxOf(doc);
            const now = await revisions(saved);
            expect(minus(original, now), "revisions lost").toEqual([]);
            const added = minus(now, original);
            if (mode === "direct") expect(added).toEqual([]);
            else for (const t of added) expect(t.split("|")[2], t).toBe("X");

            const reopened = await ReviseDocument.open({
              bytes: saved,
              filename: "saved.docx",
            });
            try {
              await doc.reject({ all: true });
              await reopened.reject({ all: true });
              for (const [where, bytes] of [
                ["in session", await docxOf(doc)],
                ["after reopen", await docxOf(reopened)],
              ] as const) {
                const xml = await part(bytes, "word/document.xml");
                const body = [...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
                  .map((t) => t[1])
                  .join("");
                const kept =
                  mode === "tracked" || !spec.changes
                    ? originalText!
                    : spec.changes === "replaced"
                      ? originalText!.replace("individual", "named person")
                      : spec.changes === "appended"
                        ? `${originalText} Appended.`
                        : originalText!;
                if (spec.changes === "split" && mode === "direct") {
                  const at = originalText!.indexOf("An Authorized");
                  expect(await paragraph(bytes, REVISED), where).toBe(
                    originalText!.slice(0, at),
                  );
                  expect(body, where).toContain(originalText!.slice(at));
                } else if (spec.changes === "merged" && mode === "direct") {
                  expect(body, where).toContain(originalText!);
                } else {
                  expect(await paragraph(bytes, REVISED), where).toBe(kept);
                }
              }
            } finally {
              reopened.dispose();
            }
          } finally {
            doc.dispose();
          }
        }, 60_000);
      }
    }
  });

  describe("an edit to one list item keeps every item's tracked changes", () => {
    // Both items of a Word list hold pending changes by "Reviewer". Editing
    // one item (or inserting/removing next to it) re-planned the whole list
    // from its rejected and accepted views: direct mode accepted the other
    // item's changes too, and tracked mode re-authored them.
    const ORIGINAL = ["Pay within 30 days of invoice", "Ship within 5 days"];
    const seed = async () => {
      const created = await ReviseDocument.create({
        html: `<ol>${ORIGINAL.map((t) => `<li>${t}</li>`).join("")}</ol><p>After the list.</p>`,
      });
      const plain = await docxOf(created);
      created.dispose();
      const doc = await ReviseDocument.open(
        { bytes: plain, filename: "list.docx" },
        { author: "Reviewer" },
      );
      try {
        const list = ydocToDocument(doc.ydoc).children[0] as any;
        for (const [index, find, replace] of [
          [0, "30", "45"],
          [1, "5", "10"],
        ] as const) {
          const result = await doc.tools.run("replace", {
            id: list.children[index].id,
            replacements: [{ find, replace }],
          });
          expect(result.ok, JSON.stringify(result)).toBe(true);
        }
        return await docxOf(doc);
      } finally {
        doc.dispose();
      }
    };
    const listItems = async (
      bytes: Uint8Array,
      decision: "accept" | "reject" = "reject",
    ) => {
      const doc = await ReviseDocument.open({ bytes, filename: "x.docx" });
      try {
        await doc[decision]({ all: true });
        const list = ydocToDocument(doc.ydoc).children[0] as any;
        return list.children.map((item: any) => blockText(item));
      } finally {
        doc.dispose();
      }
    };
    const cases: Record<
      string,
      {
        tool: string;
        input: (ids: string[]) => ReviseJsonObject;
        directRejected: string[];
        // Every item after an accept-all of the reopened file.
        accepted: string[];
        removesFirst?: boolean;
      }
    > = {
      replace: {
        tool: "replace",
        input: ([first]) => ({
          id: first,
          replacements: [{ find: "Pay", replace: "Remit" }],
        }),
        directRejected: ["Remit within 30 days of invoice", ORIGINAL[1]],
        accepted: ["Remit within 45 days of invoice", "Ship within 10 days"],
      },
      append_to_paragraph: {
        tool: "append_to_paragraph",
        input: ([first]) => ({ id: first, content: " net" }),
        directRejected: [`${ORIGINAL[0]} net`, ORIGINAL[1]],
        accepted: ["Pay within 45 days of invoice net", "Ship within 10 days"],
      },
      style_blocks: {
        tool: "style_blocks",
        input: ([first]) => ({
          selectors: `#${first}`,
          attrs: [{ name: "color", value: "#00aa00" }],
        }),
        directRejected: ORIGINAL,
        accepted: ["Pay within 45 days of invoice", "Ship within 10 days"],
      },
      insert_block: {
        tool: "insert_block",
        input: ([first]) => ({
          referenceId: first,
          position: "after",
          content: ["<li>Added item</li>"],
        }),
        directRejected: [ORIGINAL[0], "Added item", ORIGINAL[1]],
        accepted: [
          "Pay within 45 days of invoice",
          "Added item",
          "Ship within 10 days",
        ],
      },
      remove_blocks: {
        tool: "remove_blocks",
        input: ([first]) => ({ id: first }),
        directRejected: [ORIGINAL[1]],
        accepted: ["Ship within 10 days"],
        removesFirst: true,
      },
    };
    for (const mode of ["direct", "tracked"] as const) {
      for (const [label, spec] of Object.entries(cases)) {
        it(`${label} (${mode})`, async () => {
          const seeded = await seed();
          const original = await revisions(seeded);
          expect(original.length).toBeGreaterThanOrEqual(4);
          const doc = await ReviseDocument.open(
            { bytes: seeded, filename: "revised.docx" },
            { author: "X" },
          );
          let saved: Uint8Array;
          try {
            const list = ydocToDocument(doc.ydoc).children[0] as any;
            const ids = list.children.map((item: any) => item.id as string);
            const seededXml = await part(seeded, "word/document.xml");
            const firstItemRevisions = revisionsInXml(
              /<w:p\b(?:(?!<\/w:p>)[\s\S])*Pay within[\s\S]*?<\/w:p>/.exec(
                seededXml,
              )![0],
            );
            const result = await doc.tools.run(spec.tool, spec.input(ids), {
              direct: mode === "direct",
            });
            expect(result.ok, JSON.stringify(result)).toBe(true);
            saved = await docxOf(doc);
            const now = await revisions(saved);
            const expectedKept =
              mode === "direct" && spec.removesFirst
                ? minus(original, firstItemRevisions)
                : original;
            expect(minus(expectedKept, now), "revisions lost").toEqual([]);
            const added = minus(now, original);
            if (mode === "direct") expect(added).toEqual([]);
            else for (const t of added) expect(t.split("|")[2], t).toBe("X");
          } finally {
            doc.dispose();
          }
          expect(await listItems(saved)).toEqual(
            mode === "direct" ? spec.directRejected : ORIGINAL,
          );
          expect(await listItems(saved, "accept")).toEqual(spec.accepted);
        }, 60_000);
      }
    }
  });

  it.each(["direct", "tracked"] as const)(
    "insert_footnote lands in a paragraph that holds pending changes (%s)",
    async (mode) => {
      // The anchor offset is read from the paragraph with its pending
      // changes accepted, and was mapped onto the live text as an empty
      // range, which the mapper always refused: target_stale, every time.
      const doc = await ReviseDocument.open(
        fixture("legal-negotiation-after-tracked.docx"),
        { author: "X" },
      );
      try {
        const blocks = ydocToDocument(doc.ydoc).children as any[];
        const revised = blocks.find((block) =>
          blockText(block).startsWith("1.2 Subscription Fees."),
        );
        const before = doc.session.listSuggestions();
        const result = await doc.tools.run(
          "insert_footnote",
          { id: revised.id, find: "access to the Service", text: "A note." },
          { direct: mode === "direct" },
        );
        expect(result.ok, JSON.stringify(result)).toBe(true);
        const afterById = new Map(
          doc.session.listSuggestions().map((s) => [s.id, s]),
        );
        for (const s of before) expect(afterById.get(s.id), s.id).toEqual(s);
        const saved = await docxOf(doc);
        expect(await part(saved, "word/footnotes.xml")).toContain("A note.");
        const xml = await part(saved, "word/document.xml");
        const at = xml.indexOf("access to the Service");
        expect(xml.slice(at, xml.indexOf("</w:p>", at))).toMatch(
          /<w:footnoteReference /,
        );
      } finally {
        doc.dispose();
      }
    },
    60_000,
  );

  it.each(["accept", "reject"] as const)(
    "a tracked paragraph replacement %ss after a reopen without a stray space",
    async (decision) => {
      // The replacement shows old and new text side by side with a
      // display space between them. Written as plain text the space
      // survived reject; written inside the w:ins it survived accept.
      const created = await ReviseDocument.create({
        html: "<p>Provider security lead</p><p>Next paragraph.</p>",
      });
      const seeded = await docxOf(created);
      created.dispose();
      const doc = await ReviseDocument.open(
        { bytes: seeded, filename: "replace.docx" },
        { author: "Reviewer" },
      );
      let saved: Uint8Array;
      try {
        const id = (ydocToDocument(doc.ydoc).children[0] as any).id as string;
        const result = await doc.tools.run("replace_block", {
          id,
          content: ["<p>Security review owner</p>"],
        });
        expect(result.ok, JSON.stringify(result)).toBe(true);
        saved = await docxOf(doc);
        expect(
          (await revisions(saved)).filter((t) => /^(ins|del)\|/.test(t)).length,
        ).toBeGreaterThan(0);
      } finally {
        doc.dispose();
      }
      const reopened = await ReviseDocument.open({
        bytes: saved,
        filename: "saved.docx",
      });
      try {
        await reopened[decision]({ all: true });
        const out = await docxOf(reopened);
        expect(await revisions(out)).toEqual([]);
        const first = /<w:p\b[\s\S]*?<\/w:p>/.exec(
          (await part(out, "word/document.xml")).slice(
            (await part(out, "word/document.xml")).indexOf("<w:body>"),
          ),
        )![0];
        expect(
          [...first.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
            .map((t) => t[1])
            .join(""),
        ).toBe(
          decision === "accept"
            ? "Security review owner"
            : "Provider security lead",
        );
      } finally {
        reopened.dispose();
      }
    },
    60_000,
  );

  it.each(["accept", "reject"] as const)(
    "tracked paragraph spacing %ss to the right value after a reopen",
    async (decision) => {
      const created = await ReviseDocument.create({
        html: "<p>Spacing token.</p>",
      });
      const seeded = await docxOf(created);
      created.dispose();
      const after = async (bytes: Uint8Array) =>
        /<w:spacing\b[^>]*w:after="(\d+)"/.exec(
          await part(bytes, "word/document.xml"),
        )?.[1];
      const original = await after(seeded);
      expect(original).toBeDefined();
      const doc = await ReviseDocument.open(
        { bytes: seeded, filename: "spacing.docx" },
        { author: "Bob" },
      );
      let saved: Uint8Array;
      try {
        const id = (ydocToDocument(doc.ydoc).children[0] as any).id as string;
        const result = await doc.tools.run("style_blocks", {
          selectors: `#${id}`,
          attrs: [{ name: "margin-bottom", value: "36px" }],
        });
        expect(result.ok, JSON.stringify(result)).toBe(true);
        saved = await docxOf(doc);
      } finally {
        doc.dispose();
      }
      const reopened = await ReviseDocument.open({
        bytes: saved,
        filename: "saved.docx",
      });
      try {
        await reopened[decision]({ all: true });
        const out = await docxOf(reopened);
        expect(await revisions(out)).toEqual([]);
        expect(await after(out)).toBe(decision === "accept" ? "540" : original);
      } finally {
        reopened.dispose();
      }
    },
    60_000,
  );

  it.each([
    [
      "paragraph",
      "<table><tr><td><p>Spacing token.</p></td><td>Other.</td></tr></table>",
      '"type":"paragraph","id":"([^"]+)"',
    ],
    [
      "list item",
      "<table><tr><td><ul><li>Spacing token.</li></ul></td><td>Other.</td></tr></table>",
      '"type":"list-item","id":"([^"]+)"',
    ],
  ] as const)(
    "tracked spacing of a %s in a table cell rejects to its value after a reopen",
    async (_label, html, idPattern) => {
      // Paragraph property changes inside table cells were never written:
      // the body-index patcher does not reach cells.
      const created = await ReviseDocument.create({ html });
      const seeded = await docxOf(created);
      created.dispose();
      const spacing = async (bytes: Uint8Array) => {
        const xml = await part(bytes, "word/document.xml");
        const at = xml.indexOf("Spacing token.");
        const p = xml.slice(
          Math.max(xml.lastIndexOf("<w:p>", at), xml.lastIndexOf("<w:p ", at)),
          at,
        );
        return /<w:spacing\b[^>]*w:after="(\d+)"/.exec(
          p.replace(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/, ""),
        )?.[1];
      };
      const original = await spacing(seeded);
      const doc = await ReviseDocument.open(
        { bytes: seeded, filename: "cell.docx" },
        { author: "Bob" },
      );
      let saved: Uint8Array;
      try {
        const id = new RegExp(idPattern).exec(
          JSON.stringify(ydocToDocument(doc.ydoc)),
        )![1];
        const result = await doc.tools.run("style_blocks", {
          selectors: `#${id}`,
          attrs: [{ name: "margin-bottom", value: "36px" }],
        });
        expect(result.ok, JSON.stringify(result)).toBe(true);
        saved = await docxOf(doc);
      } finally {
        doc.dispose();
      }
      for (const decision of ["accept", "reject"] as const) {
        const reopened = await ReviseDocument.open({
          bytes: saved,
          filename: "saved.docx",
        });
        try {
          await reopened[decision]({ all: true });
          const out = await docxOf(reopened);
          expect(await revisions(out), decision).toEqual([]);
          expect(await spacing(out), decision).toBe(
            decision === "accept" ? "540" : original,
          );
        } finally {
          reopened.dispose();
        }
      }
    },
    60_000,
  );

  describe("tracked list level and numbering changes through a reopen", () => {
    const NESTED =
      '<ol start="3"><li>Parent<ol><li>Target</li><li>Sibling</li></ol></li><li>Last</li></ol>';
    const levels = (doc: ReviseDocument) => {
      const out: Array<[string, number]> = [];
      const walk = (node: any) => {
        if (node?.type === "list")
          for (const item of node.children)
            out.push([blockText(item), item.indentLevel ?? 0]);
        else for (const child of node?.children ?? []) walk(child);
      };
      walk(ydocToDocument(doc.ydoc));
      return out;
    };
    const itemId = (doc: ReviseDocument, text: string) => {
      let found = "";
      const walk = (node: any) => {
        if (node?.type === "list-item" && blockText(node) === text)
          found = node.id;
        for (const child of node?.children ?? []) walk(child);
      };
      walk(ydocToDocument(doc.ydoc));
      return found;
    };
    // Word writes a tracked level change as a w:pPrChange holding the
    // item's prior w:numPr. It was not written at all: after a reopen the
    // new level stuck, whichever way the change was decided.
    for (const [where, wrap] of [
      ["", (list: string) => list],
      [
        " in a table cell",
        (list: string) =>
          `<table><tr><td>${list}</td><td>Other.</td></tr></table>`,
      ],
    ] as const)
      for (const withSpacing of [false, true])
        it(`an item level change${withSpacing ? " with spacing" : ""}${where}`, async () => {
          const created = await ReviseDocument.create({ html: wrap(NESTED) });
          const seeded = await docxOf(created);
          created.dispose();
          const doc = await ReviseDocument.open(
            { bytes: seeded, filename: "nested.docx" },
            { author: "Bob" },
          );
          let saved: Uint8Array;
          let original: Array<[string, number]>;
          try {
            original = levels(doc);
            expect(original).toEqual([
              ["Parent", 0],
              ["Target", 1],
              ["Sibling", 1],
              ["Last", 0],
            ]);
            const result = await doc.tools.run("style_blocks", {
              selectors: `#${itemId(doc, "Target")}`,
              attrs: [
                { name: "indentLevel", value: 2 },
                ...(withSpacing
                  ? [{ name: "margin-bottom", value: "36px" }]
                  : []),
              ],
            });
            expect(result.ok, JSON.stringify(result)).toBe(true);
            saved = await docxOf(doc);
            expect((await revisions(saved)).length).toBeGreaterThan(0);
          } finally {
            doc.dispose();
          }
          for (const decision of ["accept", "reject"] as const) {
            const reopened = await ReviseDocument.open({
              bytes: saved,
              filename: "saved.docx",
            });
            try {
              await reopened[decision]({ all: true });
              expect(reopened.session.listSuggestions(), decision).toEqual([]);
              expect(levels(reopened), decision).toEqual(
                decision === "accept"
                  ? original.map(([text, level]) =>
                      text === "Target" ? [text, 2] : [text, level],
                    )
                  : original,
              );
            } finally {
              reopened.dispose();
            }
          }
        }, 60_000);

    // A restart or a list-type change is a new numbering definition in
    // Word; tracking it is not supported, so tracked mode refuses clearly
    // (direct mode applies it).
    for (const [label, tool, input] of [
      [
        "restarting the numbering",
        "style_blocks",
        (doc: ReviseDocument) => ({
          selectors: `#${listId(doc)}`,
          attrs: [{ name: "startNumber", value: 7 }],
        }),
      ],
      [
        "converting to bullets",
        "replace_block",
        (doc: ReviseDocument) => ({
          id: listId(doc),
          content: [`<ul><li>Alpha</li><li>Beta</li></ul>`],
        }),
      ],
    ] as const)
      for (const direct of [false, true])
        it(`${label} is ${direct ? "applied directly" : "refused when tracked"}`, async () => {
          const doc = await ReviseDocument.create(
            { html: "<ol><li>Alpha</li><li>Beta</li></ol><p>After.</p>" },
            { author: "Bob" },
          );
          try {
            await doc.setMode("suggesting");
            const before = await docxOf(doc);
            const result = await doc.tools.run(tool, input(doc), { direct });
            if (direct) {
              expect(result.ok, JSON.stringify(result)).toBe(true);
              expect(doc.session.listSuggestions()).toEqual([]);
            } else {
              expect(result.ok).toBe(false);
              expect(JSON.stringify(result)).toContain(
                "list_numbering_tracked_unsupported",
              );
              expect(await part(await docxOf(doc), "word/document.xml")).toBe(
                await part(before, "word/document.xml"),
              );
              expect(doc.session.listSuggestions()).toEqual([]);
            }
          } finally {
            doc.dispose();
          }
        }, 60_000);
    function listId(doc: ReviseDocument) {
      return (ydocToDocument(doc.ydoc).children[0] as any).id as string;
    }
  });

  describe("a reopened spacing change on a nested list item changes only its spacing", () => {
    // The imported item change carried only spacing in its snapshots, and
    // deciding it replaced the item with the snapshot (keeping its text):
    // the nested item lost its level and took the parent's numbering.
    const html = (wrap: (list: string) => string) =>
      wrap(
        '<ol start="3"><li>Parent<ol><li>Target</li><li>Sibling</li></ol></li><li>Last</li></ol>',
      );
    for (const [where, wrap] of [
      ["", (list: string) => list],
      [
        " in a table cell",
        (list: string) =>
          `<table><tr><td>${list}</td><td>Other.</td></tr></table>`,
      ],
    ] as const)
      for (const decision of ["accept", "reject"] as const)
        it(`${decision}-all${where}`, async () => {
          const created = await ReviseDocument.create({ html: html(wrap) });
          const seeded = await docxOf(created);
          created.dispose();
          const items = (doc: ReviseDocument) => {
            const lists: any[] = [];
            const walk = (node: any) => {
              if (node?.type === "list") lists.push(node);
              for (const child of node?.children ?? []) walk(child);
            };
            walk(ydocToDocument(doc.ydoc));
            return lists.flatMap((list) =>
              list.children.map((item: any) => ({
                text: blockText(item),
                level: item.indentLevel ?? 0,
                list: list.id,
              })),
            );
          };
          const doc = await ReviseDocument.open(
            { bytes: seeded, filename: "nested.docx" },
            { author: "Bob" },
          );
          let saved: Uint8Array;
          let shape: ReturnType<typeof items>;
          try {
            shape = items(doc);
            const id = (() => {
              let found = "";
              const walk = (node: any) => {
                if (node?.type === "list-item" && blockText(node) === "Target")
                  found = node.id;
                for (const child of node?.children ?? []) walk(child);
              };
              walk(ydocToDocument(doc.ydoc));
              return found;
            })();
            expect(id).toBeTruthy();
            const result = await doc.tools.run("style_blocks", {
              selectors: `#${id}`,
              attrs: [{ name: "margin-bottom", value: "36px" }],
            });
            expect(result.ok, JSON.stringify(result)).toBe(true);
            saved = await docxOf(doc);
          } finally {
            doc.dispose();
          }
          const reopened = await ReviseDocument.open({
            bytes: saved,
            filename: "saved.docx",
          });
          try {
            const levels = items(reopened).map(({ text, level }) => [
              text,
              level,
            ]);
            await reopened[decision]({ all: true });
            expect(reopened.session.listSuggestions()).toEqual([]);
            const after = items(reopened);
            expect(after.map(({ text, level }) => [text, level])).toEqual(
              levels,
            );
            expect(after.map(({ text, level }) => [text, level])).toEqual(
              shape.map(({ text, level }) => [text, level]),
            );
            // Same list (numbering) as its sibling.
            const target = after.find((item) => item.text === "Target")!;
            const sibling = after.find((item) => item.text === "Sibling")!;
            expect(target.list).toBe(sibling.list);
          } finally {
            reopened.dispose();
          }
        }, 60_000);
  });

  describe("a tracked link or spacing change in a list item round-trips through a reopen", () => {
    // List items exported their pending link changes with no revision
    // markup at all: after a reopen, reject-all was a no-op and the change
    // stuck. Paragraphs already wrote them as a deleted and an inserted run.
    const OLD = "https://example.com/old";
    const NEW = "https://example.com/new";
    type Case = {
      html: string;
      tool: string;
      input: (id: string) => ReviseJsonObject;
      // [text, target] of every link: before the change, and after it.
      before: string[];
      after: string[];
    };
    const list = (first: string) =>
      `<ul><li>${first}</li><li>Sibling.</li></ul>`;
    const linked = (href: string) =>
      `Before <a href="${href}">linked words</a> after.`;
    const cases: Record<string, Case> = {
      "replace_block removing the link": {
        html: list(linked(OLD)),
        tool: "replace_block",
        input: (id) => ({
          id,
          content: ["<li>Before linked words after.</li>"],
        }),
        before: [`linked words ${OLD}`],
        after: [],
      },
      "replace_block adding a link": {
        html: list("Before linked words after."),
        tool: "replace_block",
        input: (id) => ({ id, content: [`<li>${linked(NEW)}</li>`] }),
        before: [],
        after: [`linked words ${NEW}`],
      },
      "replace_block changing the target": {
        html: list(linked(OLD)),
        tool: "replace_block",
        input: (id) => ({ id, content: [`<li>${linked(NEW)}</li>`] }),
        before: [`linked words ${OLD}`],
        after: [`linked words ${NEW}`],
      },
      clear_formatting: {
        html: list(linked(OLD)),
        tool: "clear_formatting",
        input: (id) => ({ id }),
        before: [`linked words ${OLD}`],
        after: [],
      },
      "replace_block removing the link, in a table cell": {
        html: `<table><tr><td>${list(linked(OLD))}</td><td>Other.</td></tr></table>`,
        tool: "replace_block",
        input: (id) => ({
          id,
          content: ["<li>Before linked words after.</li>"],
        }),
        before: [`linked words ${OLD}`],
        after: [],
      },
    };
    const state = async (bytes: Uint8Array) => {
      const xml = await part(bytes, "word/document.xml");
      const rels = await part(bytes, "word/_rels/document.xml.rels");
      const target = (rid: string) =>
        new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ??
        new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels)?.[1];
      const texts = (fragment: string) =>
        [...fragment.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
          .map((t) => t[1])
          .join("");
      // Adjacent links to the same target read as one.
      const links: Array<[string, string]> = [];
      for (const m of xml.matchAll(
        /<w:hyperlink\b[^>]*r:id="([^"]+)"[^>]*>([\s\S]*?)<\/w:hyperlink>/g,
      )) {
        const href = target(m[1]) ?? "";
        const last = links[links.length - 1];
        if (last && last[1] === href) last[0] += texts(m[2]);
        else links.push([texts(m[2]), href]);
      }
      return {
        revisions: await revisions(bytes),
        text: texts(xml),
        links: links.map(([text, href]) => `${text} ${href}`),
      };
    };
    for (const [label, spec] of Object.entries(cases))
      for (const decision of ["accept", "reject"] as const)
        it(`${label}, then ${decision}-all`, async () => {
          const created = await ReviseDocument.create({ html: spec.html });
          const seeded = await docxOf(created);
          created.dispose();
          const doc = await ReviseDocument.open(
            { bytes: seeded, filename: "list.docx" },
            { author: "Bob" },
          );
          let saved: Uint8Array;
          try {
            const id = /"type":"list-item","id":"([^"]+)"/.exec(
              JSON.stringify(ydocToDocument(doc.ydoc)),
            )![1];
            const result = await doc.tools.run(spec.tool, spec.input(id));
            expect(result.ok, JSON.stringify(result)).toBe(true);
            saved = await docxOf(doc);
            expect((await revisions(saved)).length).toBeGreaterThan(0);
          } finally {
            doc.dispose();
          }
          const reopened = await ReviseDocument.open({
            bytes: saved,
            filename: "saved.docx",
          });
          try {
            await reopened[decision]({ all: true });
            const out = await state(await docxOf(reopened));
            expect(out.revisions).toEqual([]);
            expect(out.text).toContain("Before linked words after.");
            expect(out.links).toEqual(
              decision === "accept" ? spec.after : spec.before,
            );
          } finally {
            reopened.dispose();
          }
        }, 60_000);

    it.each(["accept", "reject"] as const)(
      "tracked item spacing %ss to the right value after a reopen",
      async (decision) => {
        const created = await ReviseDocument.create({ html: list("Spacing.") });
        const seeded = await docxOf(created);
        created.dispose();
        const after = async (bytes: Uint8Array) =>
          /<w:spacing\b[^>]*w:after="(\d+)"/.exec(
            await part(bytes, "word/document.xml"),
          )?.[1];
        const original = await after(seeded);
        const doc = await ReviseDocument.open(
          { bytes: seeded, filename: "list.docx" },
          { author: "Bob" },
        );
        let saved: Uint8Array;
        try {
          const id = /"type":"list-item","id":"([^"]+)"/.exec(
            JSON.stringify(ydocToDocument(doc.ydoc)),
          )![1];
          const result = await doc.tools.run("style_blocks", {
            selectors: `#${id}`,
            attrs: [{ name: "margin-bottom", value: "36px" }],
          });
          expect(result.ok, JSON.stringify(result)).toBe(true);
          saved = await docxOf(doc);
        } finally {
          doc.dispose();
        }
        const reopened = await ReviseDocument.open({
          bytes: saved,
          filename: "saved.docx",
        });
        try {
          await reopened[decision]({ all: true });
          const out = await docxOf(reopened);
          expect(await revisions(out)).toEqual([]);
          expect(await after(out)).toBe(
            decision === "accept" ? "540" : original,
          );
        } finally {
          reopened.dispose();
        }
      },
      60_000,
    );
  });

  it("a tracked removal of a cell list item that opens with another author's insertion rejects without it after a reopen", async () => {
    // Alice's pending "not" opens the only item of a list in a table cell;
    // Bob removes the item, tracked. The cell's paragraph-mark deletion is
    // a self-closing w:del in the mark's rPr, and the export took it for
    // the start of the first w:del run, so "not" lost its w:ins and
    // reject-all after a reopen restored it: "not liable" in place of
    // " liable".
    const created = await ReviseDocument.create({
      html: "<table><tr><td><ul><li>not liable for consequential damages.</li></ul></td><td>Other column.</td></tr></table>",
    });
    const plain = await docxOf(created);
    created.dispose();
    const zip = await JSZip.loadAsync(plain);
    const xml = await zip.file("word/document.xml")!.async("string");
    const seededXml = xml.replace(
      /(<w:r>(?:(?!<\/w:r>)[\s\S])*?)<w:t( xml:space="preserve")?>not liable for consequential damages\.<\/w:t><\/w:r>/,
      (_whole, open) =>
        `<w:ins w:id="777" w:author="Alice" w:date="2025-01-02T03:04:05Z">${open}<w:t>not</w:t></w:r></w:ins>${open}<w:t xml:space="preserve"> liable for consequential damages.</w:t></w:r>`,
    );
    expect(seededXml).not.toBe(xml);
    zip.file("word/document.xml", seededXml);
    const seeded = await zip.generateAsync({ type: "uint8array" });

    const doc = await ReviseDocument.open(
      { bytes: seeded, filename: "seeded.docx" },
      { author: "Bob" },
    );
    let saved: Uint8Array;
    try {
      const id = /"type":"list-item","id":"([^"]+)"/.exec(
        JSON.stringify(ydocToDocument(doc.ydoc)),
      )![1];
      const result = await doc.tools.run("remove_blocks", { id });
      expect(result.ok, JSON.stringify(result)).toBe(true);
      saved = await docxOf(doc);
    } finally {
      doc.dispose();
    }
    const savedXml = await part(saved, "word/document.xml");
    expect(savedXml).toMatch(
      /<w:ins [^>]*w:author="Alice"[^>]*><w:del [^>]*w:author="Bob"[^>]*>[\s\S]*?not[\s\S]*?<\/w:del><\/w:ins>/,
    );
    const cellText = async (bytes: Uint8Array) => {
      const body = await part(bytes, "word/document.xml");
      const cell = /<w:tc>[\s\S]*?<\/w:tc>/.exec(body)![0];
      return [...cell.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
        .map((t) => t[1])
        .join("");
    };
    for (const [decision, expected] of [
      ["reject", " liable for consequential damages."],
      ["accept", ""],
    ] as const) {
      const reopened = await ReviseDocument.open({
        bytes: saved,
        filename: "saved.docx",
      });
      try {
        await reopened[decision]({ all: true });
        const out = await docxOf(reopened);
        expect(await revisions(out), decision).toEqual([]);
        expect(await cellText(out), decision).toBe(expected);
      } finally {
        reopened.dispose();
      }
    }
  }, 60_000);

  describe("links in revised list items survive a reopen", () => {
    const LINK = "https://example.com/link";
    // A file whose first list item's linked text "OLD linked" has Alice's
    // pending OLD -> NEW (and the second item OLD_B -> NEW_B).
    const seeded = async (html: string) => {
      const created = await ReviseDocument.create({ html });
      const plain = await docxOf(created);
      created.dispose();
      const doc = await ReviseDocument.open(
        { bytes: plain, filename: "seed.docx" },
        { author: "Alice" },
      );
      try {
        for (const [find, replace] of [
          ["OLD linked", "NEW linked"],
          ["OLD_B", "NEW_B"],
        ]) {
          const hit = doc.grep(find.split(" ")[0])[0];
          if (!hit) continue;
          const result = await doc.tools.run("replace", {
            id: hit.id,
            replacements: [
              { find: find.split(" ")[0], replace: replace.split(" ")[0] },
            ],
          });
          expect(result.ok, JSON.stringify(result)).toBe(true);
        }
        return await docxOf(doc);
      } finally {
        doc.dispose();
      }
    };
    const firstItemId = (doc: ReviseDocument) => {
      const found = JSON.stringify(ydocToDocument(doc.ydoc));
      return /"type":"list-item","id":"([^"]+)"/.exec(found)![1];
    };
    const edit = async (
      bytes: Uint8Array,
      tool: string,
      input: (id: string) => ReviseJsonObject,
      direct: boolean,
    ) => {
      const doc = await ReviseDocument.open(
        { bytes, filename: "revised.docx" },
        { author: "Bob" },
      );
      try {
        const result = await doc.tools.run(tool, input(firstItemId(doc)), {
          direct,
        });
        expect(result.ok, JSON.stringify(result)).toBe(true);
        return await docxOf(doc);
      } finally {
        doc.dispose();
      }
    };
    const decided = async (
      bytes: Uint8Array,
      decision: "accept" | "reject",
    ) => {
      const doc = await ReviseDocument.open({ bytes, filename: "saved.docx" });
      try {
        await doc[decision]({ all: true });
        const out = await docxOf(doc);
        const xml = await part(out, "word/document.xml");
        const texts = (fragment: string) =>
          [...fragment.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
            .map((t) => t[1])
            .join("");
        return {
          revisions: await revisions(out),
          paragraphs: [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
            .map((p) => texts(p[0]))
            .filter(Boolean),
          listItems: (xml.match(/<w:numPr>/g) ?? []).length,
          links: [
            ...xml.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/g),
          ].map((m) => texts(m[1])),
        };
      } finally {
        doc.dispose();
      }
    };

    it("tracked removal of a revised linked item in a table cell", async () => {
      const bytes = await seeded(
        `<table><tr><td><ul><li>Before <a href="${LINK}">OLD linked</a> after.</li><li>Sibling item.</li></ul></td><td>Other cell.</td></tr></table>`,
      );
      const saved = await edit(bytes, "remove_blocks", (id) => ({ id }), false);
      const rejected = await decided(saved, "reject");
      expect(rejected.revisions).toEqual([]);
      expect(rejected.paragraphs).toContain("Before OLD linked after.");
      expect(rejected.links.join("")).toBe("OLD linked");
      const accepted = await decided(saved, "accept");
      expect(accepted.revisions).toEqual([]);
      expect(accepted.paragraphs.join("|")).not.toMatch(/linked/);
      expect(accepted.links).toEqual([]);
      // The removed item is gone, not left behind as an empty item.
      expect(accepted.listItems).toBe(1);
      expect(rejected.listItems).toBe(2);
    }, 60_000);

    for (const tool of ["replace_block", "clear_formatting"] as const)
      it(`direct ${tool} that drops the link of a revised item`, async () => {
        const bytes = await seeded(
          `<ul><li>Before <a href="${LINK}">OLD linked</a> after.</li><li>Second OLD_B item.</li></ul>`,
        );
        const saved = await edit(
          bytes,
          tool,
          (id) =>
            tool === "replace_block"
              ? { id, content: ["<li>Before NEW linked after.</li>"] }
              : { id },
          true,
        );
        for (const decision of ["accept", "reject"] as const) {
          const out = await decided(saved, decision);
          expect(out.revisions, decision).toEqual([]);
          // Direct: the link is gone whichever way Alice's changes go.
          expect(out.links, decision).toEqual([]);
          expect(out.paragraphs, decision).toEqual(
            decision === "accept"
              ? ["Before NEW linked after.", "Second NEW_B item."]
              : ["Before OLD linked after.", "Second OLD_B item."],
          );
        }
      }, 60_000);
  });

  it.each(["accept", "reject"] as const)(
    "a tracked replace_block that drops a link %ss correctly after a reopen",
    async (decision) => {
      const TARGET = "https://example.com/probe";
      const created = await ReviseDocument.create({
        html: `<p>Before <a href="${TARGET}">linked words</a> after.</p>`,
      });
      const seeded = await docxOf(created);
      created.dispose();
      const doc = await ReviseDocument.open(
        { bytes: seeded, filename: "link.docx" },
        { author: "Reviewer" },
      );
      let saved: Uint8Array;
      try {
        const id = (ydocToDocument(doc.ydoc).children[0] as any).id as string;
        const result = await doc.tools.run("replace_block", {
          id,
          content: ["<p>Before linked words after.</p>"],
        });
        expect(result.ok, JSON.stringify(result)).toBe(true);
        // In the session, rejecting restores the link.
        saved = await docxOf(doc);
      } finally {
        doc.dispose();
      }
      const reopened = await ReviseDocument.open({
        bytes: saved,
        filename: "saved.docx",
      });
      try {
        await reopened[decision]({ all: true });
        const out = await docxOf(reopened);
        expect(await revisions(out)).toEqual([]);
        const xml = await part(out, "word/document.xml");
        const text = [...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
          .map((t) => t[1])
          .join("");
        expect(text).toBe("Before linked words after.");
        const linked =
          /<w:hyperlink\b[^>]*r:id="([^"]+)"[^>]*>([\s\S]*?)<\/w:hyperlink>/.exec(
            xml,
          );
        if (decision === "accept") {
          expect(linked).toBeNull();
        } else {
          expect(linked).not.toBeNull();
          expect(
            [...linked![2].matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
              .map((t) => t[1])
              .join(""),
          ).toBe("linked words");
          expect(await part(out, "word/_rels/document.xml.rels")).toContain(
            `Id="${linked![1]}"`,
          );
          expect(await part(out, "word/_rels/document.xml.rels")).toContain(
            TARGET,
          );
        }
      } finally {
        reopened.dispose();
      }
    },
    60_000,
  );

  describe("tracked formatting over a hyperlink keeps the link through a saved accept- or reject-all", () => {
    // A hyperlink is a w:hyperlink element in Word, never a run property,
    // so a w:rPrChange's prior properties never name it. Rejecting the
    // reopened change must not take the link away; accepting keeps it too.
    const TARGET = "https://example.com/terms";
    const linkRuns = (xml: string) =>
      [
        ...xml.matchAll(
          /<w:hyperlink\b[^>]*r:id="([^"]+)"[^>]*>([\s\S]*?)<\/w:hyperlink>/g,
        ),
      ].map((m) => ({
        rid: m[1],
        text: [...m[2].matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
          .map((t) => t[1])
          .join(""),
      }));
    const linkSurvives = async (bytes: Uint8Array) => {
      const runs = linkRuns(await part(bytes, "word/document.xml"));
      expect(runs.map((r) => r.text).join("")).toBe("the terms");
      const rels = await part(bytes, "word/_rels/document.xml.rels");
      for (const { rid } of runs)
        expect(rels).toMatch(
          new RegExp(
            `Id="${rid}"[^>]*Target="${TARGET.replace(/[./]/g, "\\$&")}"`,
          ),
        );
    };
    for (const tool of ["style_blocks", "clear_formatting"] as const)
      for (const decision of ["reject", "accept"] as const)
        it(`${tool}, then ${decision}-all`, async () => {
          const created = await ReviseDocument.create({
            html: `<p>Read <b>all of</b> <a href="${TARGET}">the terms</a> now.</p>`,
          });
          const seeded = await docxOf(created);
          created.dispose();
          await linkSurvives(seeded);
          const doc = await ReviseDocument.open(
            { bytes: seeded, filename: "link.docx" },
            { author: "Reviewer" },
          );
          let saved: Uint8Array;
          try {
            const id = (ydocToDocument(doc.ydoc).children[0] as any)
              .id as string;
            const result = await doc.tools.run(
              tool,
              tool === "style_blocks"
                ? {
                    selectors: `#${id}`,
                    attrs: [{ name: "color", value: "#ff0000" }],
                  }
                : { id },
            );
            expect(result.ok, JSON.stringify(result)).toBe(true);
            saved = await docxOf(doc);
            expect(
              (await revisions(saved)).some((t) => t.startsWith("rPrChange")),
            ).toBe(true);
          } finally {
            doc.dispose();
          }
          const reopened = await ReviseDocument.open({
            bytes: saved,
            filename: "saved.docx",
          });
          try {
            await reopened[decision]({ all: true });
            const out = await docxOf(reopened);
            expect(await revisions(out)).toEqual([]);
            if (tool === "clear_formatting" && decision === "accept") {
              // clear_formatting removes links: accepting it after a reopen
              // removes this one (the change is written as the linked run
              // deleted and the plain run inserted).
              const xml = await part(out, "word/document.xml");
              expect(linkRuns(xml)).toEqual([]);
              expect(
                [...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
                  .map((t) => t[1])
                  .join(""),
              ).toBe("Read all of the terms now.");
            } else {
              await linkSurvives(out);
            }
          } finally {
            reopened.dispose();
          }
        }, 60_000);
  });

  it("a tracked join of footnote paragraphs, one holding another author's insertion, accepts as one paragraph after a reopen", async () => {
    // The joined paragraph's mark deletion was written for body and cell
    // paragraphs only; in a note it was dropped, so accept-all on the
    // reopened file left the two paragraphs apart.
    const created = await ReviseDocument.create({
      html: "<p>Body text here.</p>",
    });
    const pid = (ydocToDocument(created.ydoc).children[0] as any).id as string;
    const note = await created.tools.run("insert_footnote", {
      id: pid,
      find: "here",
      text: "Alpha base tail",
    });
    expect(note.ok, JSON.stringify(note)).toBe(true);
    const plain = await docxOf(created);
    created.dispose();
    const zip = await JSZip.loadAsync(plain);
    const notesXml = await zip.file("word/footnotes.xml")!.async("string");
    const seededXml = notesXml.replace(
      /(<w:r>(?:(?!<\/w:r>)[\s\S])*?)<w:t( xml:space="preserve")?>Alpha base tail<\/w:t><\/w:r>([\s\S]*?<\/w:p>)/,
      (_whole, open, _space, rest) =>
        `<w:ins w:id="777" w:author="Alice" w:date="2025-01-02T03:04:05Z">${open}<w:t>Alpha base</w:t></w:r></w:ins>${open}<w:t xml:space="preserve"> tail</w:t></w:r>${rest}<w:p>${open}<w:t>Beta base tail</w:t></w:r></w:p>`,
    );
    expect(seededXml).not.toBe(notesXml);
    zip.file("word/footnotes.xml", seededXml);
    const seeded = await zip.generateAsync({ type: "uint8array" });

    const doc = await ReviseDocument.open(
      { bytes: seeded, filename: "notes.docx" },
      { author: "Bob" },
    );
    let saved: Uint8Array;
    try {
      const paragraphs = (ydocToDocument(doc.ydoc) as any).notes[0].children;
      expect(paragraphs.map(blockText)).toEqual([
        expect.stringContaining("Alpha base tail"),
        "Beta base tail",
      ]);
      const result = await doc.tools.run("join_paragraphs", {
        firstId: paragraphs[0].id,
        secondId: paragraphs[1].id,
        separator: " ",
      });
      expect(result.ok, JSON.stringify(result)).toBe(true);
      saved = await docxOf(doc);
    } finally {
      doc.dispose();
    }
    const noteParagraphs = async (bytes: Uint8Array) => {
      const xml = await part(bytes, "word/footnotes.xml");
      const body = [
        ...xml.matchAll(/<w:footnote\b[^>]*>([\s\S]*?)<\/w:footnote>/g),
      ]
        .map((m) => m[1])
        .find((n) => n.includes("Beta"))!;
      return [...body.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((p) =>
        [...p[0].matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
          .map((t) => t[1])
          .join(""),
      );
    };
    for (const [decision, expected] of [
      ["accept", ["Alpha base tail Beta base tail"]],
      ["reject", [" tail", "Beta base tail"]],
    ] as const) {
      const reopened = await ReviseDocument.open({
        bytes: saved,
        filename: "saved.docx",
      });
      try {
        await reopened[decision]({ all: true });
        const out = await docxOf(reopened);
        expect(await revisions(out), decision).toEqual([]);
        expect(await noteParagraphs(out), decision).toEqual(expected);
      } finally {
        reopened.dispose();
      }
    }
  }, 60_000);

  it("refuses a footnote inside a footnote and changes nothing", async () => {
    // Word has no note inside a note: the reference read back as a
    // superscript "0" with its body gone.
    const doc = await ReviseDocument.create({ html: "<p>Body text here.</p>" });
    try {
      const pid = (ydocToDocument(doc.ydoc).children[0] as any).id as string;
      const note = await doc.tools.run("insert_footnote", {
        id: pid,
        find: "here",
        text: "Alpha base tail",
      });
      expect(note.ok, JSON.stringify(note)).toBe(true);
      const before = JSON.stringify(ydocToDocument(doc.ydoc));
      const noteParagraph = (ydocToDocument(doc.ydoc) as any).notes[0]
        .children[0].id as string;
      const nested = await doc.tools.run("insert_footnote", {
        id: noteParagraph,
        find: "tail",
        text: "Nested note",
      });
      expect(nested.ok).toBe(false);
      expect(JSON.stringify(nested)).toContain("footnote_in_note_unsupported");
      expect(JSON.stringify(ydocToDocument(doc.ydoc))).toBe(before);
    } finally {
      doc.dispose();
    }
  }, 60_000);

  describe("tracked paragraph insertions and splits in cells and notes decide alike after a reopen", () => {
    // Rejecting them on the reopened file left an extra empty paragraph.
    const seeded = async (where: "cell" | "note") => {
      const created = await ReviseDocument.create({
        html:
          where === "cell"
            ? "<table><tr><td><p>Alpha base tail</p></td><td>Other.</td></tr></table>"
            : "<p>Body text here.</p>",
      });
      try {
        if (where === "note") {
          const pid = (ydocToDocument(created.ydoc).children[0] as any)
            .id as string;
          const note = await created.tools.run("insert_footnote", {
            id: pid,
            find: "here",
            text: "Alpha base tail",
          });
          expect(note.ok, JSON.stringify(note)).toBe(true);
        }
        return await docxOf(created);
      } finally {
        created.dispose();
      }
    };
    const target = (doc: ReviseDocument, where: "cell" | "note") => {
      const root = ydocToDocument(doc.ydoc) as any;
      return where === "note"
        ? root.notes[0].children[0].id
        : root.children[0].children[0].children[0].children[0].id;
    };
    const story = (doc: ReviseDocument, where: "cell" | "note") => {
      const root = ydocToDocument(doc.ydoc) as any;
      const blocks =
        where === "note"
          ? root.notes[0].children
          : root.children[0].children[0].children[0].children;
      return blocks.map(blockText);
    };
    it.each([
      ["cell", "insert_block"],
      ["cell", "break_paragraph"],
      ["note", "insert_block"],
      ["note", "break_paragraph"],
    ] as const)(
      "%s %s",
      async (where, tool) => {
        const doc = await ReviseDocument.open(
          { bytes: await seeded(where), filename: "seed.docx" },
          { author: "Bob" },
        );
        let saved: Uint8Array;
        let snapshot: Awaited<ReturnType<ReviseDocument["snapshot"]>>;
        try {
          const id = target(doc, where);
          const result = await doc.tools.run(
            tool,
            tool === "insert_block"
              ? {
                  referenceId: id,
                  position: "after",
                  content: ["<p>Added clause</p>"],
                }
              : { id, substring: "tail" },
          );
          expect(result.ok, JSON.stringify(result)).toBe(true);
          saved = await docxOf(doc);
          snapshot = await doc.snapshot();
        } finally {
          doc.dispose();
        }
        for (const decision of ["accept", "reject"] as const) {
          const live = await ReviseDocument.restore(snapshot);
          const fresh = await ReviseDocument.open({
            bytes: saved,
            filename: "saved.docx",
          });
          try {
            await live[decision]({ all: true });
            await fresh[decision]({ all: true });
            expect(story(fresh, where), decision).toEqual(story(live, where));
          } finally {
            live.dispose();
            fresh.dispose();
          }
        }
      },
      60_000,
    );
  });

  describe("tracked formatting over a footnote reference keeps the note through a saved reject-all", () => {
    // The note reference run takes a w:rPrChange like the rest of the
    // paragraph. Its prior run properties in the .docx never name the
    // reference (that is the w:footnoteReference element, not a property),
    // so rejecting the reopened change stripped the reference mark: the
    // paragraph kept a literal "1" and the note body was dropped.
    for (const tool of ["style_blocks", "clear_formatting"] as const) {
      it(
        tool,
        async () => {
          const created = await ReviseDocument.create({
            html: "<p>Alpha bold <b>bravo</b> tail.</p><p>Second paragraph.</p>",
          });
          const pid = (ydocToDocument(created.ydoc).children[0] as any)
            .id as string;
          const note = await created.tools.run("insert_footnote", {
            id: pid,
            find: "Alpha",
            text: "Footnote marker.",
          });
          expect(note.ok, JSON.stringify(note)).toBe(true);
          const seeded = await docxOf(created);
          created.dispose();
          expect(await part(seeded, "word/footnotes.xml")).toContain(
            "Footnote marker.",
          );

          const doc = await ReviseDocument.open(
            { bytes: seeded, filename: "notes.docx" },
            { author: "Reviewer" },
          );
          let saved: Uint8Array;
          try {
            const id = (ydocToDocument(doc.ydoc).children[0] as any)
              .id as string;
            const result = await doc.tools.run(
              tool,
              tool === "style_blocks"
                ? {
                    selectors: `#${id}`,
                    attrs: [{ name: "color", value: "#ff0000" }],
                  }
                : { id },
            );
            expect(result.ok, JSON.stringify(result)).toBe(true);
            saved = await docxOf(doc);
            expect(
              (await revisions(saved)).filter((t) => t.startsWith("rPrChange"))
                .length,
              JSON.stringify(result),
            ).toBeGreaterThan(0);
          } finally {
            doc.dispose();
          }
          const reopened = await ReviseDocument.open({
            bytes: saved,
            filename: "saved.docx",
          });
          try {
            await reopened.reject({ all: true });
            const out = await docxOf(reopened);
            expect(await part(out, "word/footnotes.xml")).toContain(
              "Footnote marker.",
            );
            const body = await part(out, "word/document.xml");
            expect(body).toMatch(/<w:footnoteReference w:id="\d+"\/>/);
            expect(body).not.toContain(">Alpha1<");
            expect(await revisions(out)).toEqual([]);
          } finally {
            reopened.dispose();
          }
        },
        60_000,
      );
    }
  });

  it("names the problem with a code", async () => {
    await expect(ReviseDocument.open("missing.docx")).rejects.toMatchObject({
      code: "not_found",
    });
    await expect(
      ReviseDocument.open({ bytes: new Uint8Array(), filename: "x.xyz" }),
    ).rejects.toMatchObject({ code: "unsupported_format" });
  });
});
