// @vitest-environment node
// Ported from Revise's main repository on each release; changes land there.

import { describe, expect, it } from "vitest";
import { ReviseDocument, convert } from "../src/lib";

// Pending tracked changes are what a reviewer is there to see. Reading,
// grepping, and searching show them by default (the tracked view), and
// accepted/original show the text with every change accepted or rejected.

async function redline() {
  const doc = await ReviseDocument.create({
    html: "<p>Fees are $100 per hour.</p><p>Payment is due in 30 days.</p><p>Late fees apply.</p>",
  });
  await doc.setMode("suggesting");
  const jane = await ReviseDocument.restore(await doc.snapshot(), {
    author: "Jane Smith",
  });
  doc.dispose();
  await jane.replaceAll("$100", "$120");
  return jane;
}

async function readAll(doc: ReviseDocument) {
  const result = await doc.tools.run("read_blocks_from_index", { index: 0 });
  if (!result.ok) throw new Error(result.error.message);
  return result.value.context?.html ?? "";
}

describe("tracked view", () => {
  it("marks pending changes inline, with their id and author", async () => {
    const doc = await redline();
    try {
      const [change] = await doc.changes();
      const html = await doc.viewHtml(await readAll(doc), "tracked");
      expect(html).toContain(
        `<del data-change-id="${change.id}" data-author="Jane Smith">100</del>`,
      );
      expect(html).toContain(
        `<ins data-change-id="${change.id}" data-author="Jane Smith">120</ins>`,
      );
      // Block ids stay what read shows, so edits can still target them.
      expect(html).toMatch(/<p id="[^"]+"[^>]*>Fees are /);
    } finally {
      doc.dispose();
    }
  });

  it("shows the original text, and accepted is what the tools read", async () => {
    const doc = await redline();
    try {
      const accepted = await readAll(doc);
      expect(await doc.viewHtml(accepted, "accepted")).toBe(accepted);
      expect(accepted).toContain("$120");
      expect(accepted).not.toContain("$100");
      const original = await doc.viewHtml(accepted, "original");
      expect(original).toContain("Fees are $100 per hour.");
      expect(original).not.toContain("$120");
      expect(original).not.toMatch(/<ins|<del/);
    } finally {
      doc.dispose();
    }
  });

  it("shows a deleted paragraph where it was", async () => {
    const doc = await redline();
    try {
      const html = await readAll(doc);
      const id = /<p id="([^"]+)"[^>]*>Payment/.exec(
        await doc.viewHtml(html, "original"),
      )![1];
      const removed = await doc.tools.run("remove_blocks", { id });
      expect(removed.ok).toBe(true);
      const after = await readAll(doc);
      expect(after).not.toContain("Payment");
      const tracked = await doc.viewHtml(after, "tracked");
      expect(tracked).toMatch(
        /<del data-change-id="[^"]+" data-author="Jane Smith">Payment is due in 30 days\.<\/del>/,
      );
      expect(tracked.indexOf("Payment")).toBeLessThan(tracked.indexOf("Late"));
    } finally {
      doc.dispose();
    }
  });

  it("greps each view, labeling matches inside changes", async () => {
    const doc = await redline();
    try {
      expect(doc.grep("$100", { view: "accepted" })).toEqual([]);
      expect(doc.grep("$100", { view: "original" })[0].text).toBe(
        "Fees are $100 per hour.",
      );
      const [hit] = doc.grep("$1", { view: "tracked" });
      expect(hit.text).toBe("Fees are $[-100-]{+120+} per hour.");
      expect(hit.count).toBe(2);
      expect(hit.matches.map((m) => [m.text, m.change])).toEqual([
        ["$1", "deleted"],
        ["$1", "inserted"],
      ]);
      // Text no change touches counts once, unlabeled.
      const [fees] = doc.grep("Fees are");
      expect(fees.count).toBe(1);
      expect(fees.matches[0]).toMatchObject({ change: null, changeId: null });
    } finally {
      doc.dispose();
    }
  });
});

describe("tracked markup pasted back into a tool", () => {
  it("keeps the inserted text and drops the deleted text", async () => {
    const doc = await redline();
    try {
      const tracked = await doc.viewHtml(await readAll(doc), "tracked");
      const block = /<p id="[^"]+"[^>]*>Fees are[\s\S]*?<\/p>/.exec(
        tracked,
      )![0];
      const id = /id="([^"]+)"/.exec(block)![1];
      await doc.setMode("editing");
      const result = await doc.tools.run("replace_block", {
        id,
        content: [block.replace("per hour", "per day")],
      });
      expect(result.ok).toBe(true);
      const [hit] = doc.grep("Fees", { view: "accepted" });
      expect(hit.text).toBe("Fees are $120 per day.");
    } finally {
      doc.dispose();
    }
  });
});

describe("views in export and convert", () => {
  it("exports the original, accepted, and tracked text", async () => {
    const doc = await redline();
    try {
      const text = async (view: "tracked" | "accepted" | "original") => {
        const result = await doc.export({ format: "txt", view });
        return {
          text: new TextDecoder().decode(result.bytes),
          view: result.view,
        };
      };
      expect(await text("original")).toMatchObject({ view: "original" });
      expect((await text("original")).text).toContain("$100");
      expect((await text("accepted")).text).toContain("$120");
      // A .txt can't hold changes: tracked comes out accepted, and says so.
      expect(await text("tracked")).toMatchObject({ view: "accepted" });
      // The document keeps its changes.
      expect((await doc.changes()).length).toBe(1);
      const docx = await doc.export({ format: "docx", view: "original" });
      const reopened = await ReviseDocument.open({
        bytes: docx.bytes,
        filename: "original.docx",
      });
      try {
        expect(await reopened.changes()).toEqual([]);
        expect(reopened.grep("$100")).toHaveLength(1);
      } finally {
        reopened.dispose();
      }
      await expect(
        doc.export({ format: "txt", mode: "review", view: "original" }),
      ).rejects.toMatchObject({ code: "usage" });
    } finally {
      doc.dispose();
    }
  });

  it("converts the original text of a redlined file", async () => {
    const doc = await redline();
    const docx = (await doc.export({ format: "docx" })).bytes;
    doc.dispose();
    const source = { bytes: docx, filename: "redline.docx" };
    const text = async (view: "accepted" | "original") =>
      new TextDecoder().decode(await convert(source, "txt", { view }));
    expect(await text("original")).toContain("$100");
    expect(await text("accepted")).toContain("$120");
  });
});
