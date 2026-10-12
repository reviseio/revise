/**
 * Word files in and out of a live document: graft export (untouched content
 * byte for byte from the original package) and re-reading a .docx that
 * changed elsewhere while keeping block ids.
 */
import { engineModules } from "@reviseio/engine";
import * as Y from "yjs";
import type { BlockNode, DocRoot } from "@reviseio/engine";
import {
  createReviseYDoc,
  documentToDocx,
  parseDocument,
  ydocSourceFidelity,
  ydocToDocument,
} from "@reviseio/engine";
import { createHash } from "node:crypto";
import { seedDocRootWithClaim } from "@reviseio/engine";
import { ReviseError } from "./errors";

/** What a .docx export actually did. */
export interface DocxExportOutcome {
  /** "graft": untouched content copied byte for byte from the original;
   * "regenerated": the whole file written fresh (no original, or graft
   * refused it, see fallbackCode). */
  method: "graft" | "regenerated";
  exact: boolean;
  regeneratedBlocks: number;
  fallbackCode?: string;
}

/**
 * The session as a .docx: the model's changes grafted onto the session's
 * package. Without a package (a document not opened from .docx) it is the
 * ordinary export. Paragraphs and rows the export writes fresh get a
 * w14:paraId if they lack one (see stampParaIds); untouched ones never do.
 * Reports what the converter actually did, not what it was asked to do.
 */
export async function exportSessionDocx(
  ydoc: Y.Doc,
  pkg: Uint8Array | null,
  mode: "review" | "final" = "review",
): Promise<{ bytes: Uint8Array; outcome: DocxExportOutcome }> {
  let doc = ydocToDocument(ydoc);
  const grafting = Boolean(pkg && doc.docxSource);
  const regenerated = grafting
    ? new Set((await ydocSourceFidelity(ydoc)).regeneratedBlockIds)
    : null;
  if (mode === "final") {
    const { scrubDocument } = await engineModules.scrubber();
    doc = scrubDocument(doc, {
      removeComments: true,
      resolveSuggestions: "accept",
    }).doc;
  }
  doc = stampParaIds(
    doc,
    regenerated,
    grafting ? graftKeptParts(doc, await engineModules.docxGraft()) : undefined,
  );
  if (!grafting || !pkg || !doc.docxSource) {
    return {
      bytes: await documentToDocx(doc),
      outcome: {
        method: "regenerated",
        exact: false,
        regeneratedBlocks: doc.children.length,
        fallbackCode: doc.docxSource ? "no-original-package" : "not-from-docx",
      },
    };
  }
  const { hashBytes } = await engineModules.docxGraft();
  // The package is the capture's own descendant (patches keep its body
  // structure), so pair the capture with it.
  doc = {
    ...doc,
    docxSource: { ...doc.docxSource, sourceHash: await hashBytes(pkg) },
  };
  let reported: DocxExportOutcome = {
    method: "regenerated",
    exact: false,
    regeneratedBlocks: doc.children.length,
    fallbackCode: "no-outcome",
  };
  const bytes = await documentToDocx(doc, {
    originalDocx: pkg,
    onGraftOutcome: (outcome) => {
      reported =
        outcome.mode === "regenerated"
          ? {
              method: "regenerated",
              exact: false,
              regeneratedBlocks: doc.children.length,
              fallbackCode: outcome.fallbackCode,
            }
          : {
              method: "graft",
              exact: outcome.mode === "exact",
              regeneratedBlocks: outcome.regeneratedBlockCount ?? 0,
            };
    },
  });
  return { bytes, outcome: reported };
}

// ── block identity across re-reads ──────────────────────────────────────────
//
// When a file changes outside revise it is re-imported, and import mints
// fresh ids. Old ids carry over only where identity is known: Word's own
// w14:paraId (or the one revise wrote for a paragraph it created), or an
// identical block. Never by position: an id an agent read must never land
// on a different paragraph. Anything unmatched gets a fresh id.

interface IdNode {
  id?: string;
  type?: string;
  oxmlAttrs?: string;
  children?: unknown[];
}

const PARA_ID_KEY = "w14:paraId";

function paraIdOf(node: IdNode): string | undefined {
  if (!node.oxmlAttrs) return undefined;
  try {
    const attrs = JSON.parse(node.oxmlAttrs) as Record<string, unknown>;
    const value = attrs[PARA_ID_KEY];
    return typeof value === "string" ? value.toUpperCase() : undefined;
  } catch {
    return undefined;
  }
}

/** The w14:paraId revise writes for a block it created: derived from the
 * block id, so it is the same on every save and maps back on re-read.
 * Word's rules: 8 hex digits, below 0x80000000. */
export function derivedParaId(blockId: string, salt = 0): string {
  const value =
    createHash("sha1")
      .update(salt ? `${blockId}:${salt}` : blockId)
      .digest()
      .readUInt32BE(0) & 0x7fffffff;
  return (value || 1).toString(16).toUpperCase().padStart(8, "0");
}

function childrenOf(node: IdNode): IdNode[] {
  return (Array.isArray(node.children) ? node.children : []).filter(
    (child): child is IdNode =>
      child !== null && typeof child === "object" && "type" in child,
  );
}

/** Content identity: the node's kind, its text, and its children's, in
 * order. Formatting is left out: the model normalizes it (default fonts,
 * spacing) differently from a fresh import, and a formatting-only change
 * does not make an agent's reading of the text stale. */
function contentSignature(node: unknown): string {
  const n = node as { type?: string; text?: string; children?: unknown[] };
  if (!n || typeof n !== "object") return "";
  if (n.type === "text") return JSON.stringify(n.text ?? "");
  const children = Array.isArray(n.children) ? n.children : [];
  return `${n.type ?? "?"}(${children.map(contentSignature).join(",")})`;
}

function walkIds(node: IdNode, visit: (node: IdNode) => void): void {
  visit(node);
  for (const child of childrenOf(node)) walkIds(child, visit);
}

/**
 * Give paragraphs and table rows a w14:paraId where they lack one, but only
 * in blocks the export writes fresh (`regenerated`; null = every block):
 * untouched content keeps its exact bytes. Returns a copy.
 */
/** What graft export copies from the original: whole parts, and changed
 * rows whose own properties (their head) are unchanged. */
interface GraftKept {
  whole: Set<string>;
  head: Set<string>;
}

/**
 * Rows, cells, and list items the graft will copy from the original bytes:
 * those whose signature still matches the capture. A changed table or list
 * is regenerated part by part, so stamping a paraId on an untouched row
 * would change its signature and force the whole table to be rewritten.
 * Mirrors graftExport's own checks (emitParts, cellLevelRow).
 */
function graftKeptParts(
  doc: DocRoot,
  graft: Awaited<ReturnType<typeof engineModules.docxGraft>>,
): GraftKept {
  const kept = { whole: new Set<string>(), head: new Set<string>() };
  const source = doc.docxSource as
    | {
        signatureVersion?: number;
        parts?: Record<
          string,
          {
            parts: Array<{
              id: string;
              signature: string;
              head?: string;
              cells?: Array<{ id: string; signature: string }>;
            }>;
          }
        >;
      }
    | undefined;
  if (!source?.parts) return kept;
  const version =
    source.signatureVersion === graft.CURRENT_GRAFT_SIGNATURE_VERSION
      ? graft.CURRENT_GRAFT_SIGNATURE_VERSION
      : 1;
  for (const block of doc.children) {
    const captured = block.id ? source.parts[block.id] : undefined;
    if (!captured || (block.type !== "table" && block.type !== "list")) {
      continue;
    }
    for (const part of block.children as unknown as IdNode[]) {
      const entry = captured.parts.find((p) => p.id === part.id);
      if (!entry || !part.id) continue;
      if (entry.signature === graft.partSignature(part, version)) {
        kept.whole.add(part.id);
        continue;
      }
      // A changed row with unchanged properties keeps its untouched cells.
      if (
        block.type === "table" &&
        entry.cells &&
        entry.head === graft.partHeadSignature(part, version)
      ) {
        // The row's own opening bytes come from the original too.
        kept.head.add(part.id);
        (part.children as IdNode[] | undefined)?.forEach((cell, c) => {
          const cellEntry = entry.cells![c];
          if (
            cell.id &&
            cellEntry?.id === cell.id &&
            cellEntry.signature === graft.partSignature(cell, version)
          ) {
            kept.whole.add(cell.id);
          }
        });
      }
    }
  }
  return kept;
}

export function stampParaIds(
  doc: DocRoot,
  regenerated: ReadonlySet<string> | null,
  /** Parts the graft copies from the original: never stamped. */
  kept: GraftKept = { whole: new Set(), head: new Set() },
): DocRoot {
  const used = new Set<string>();
  for (const block of doc.children) {
    walkIds(block as IdNode, (node) => {
      const id = paraIdOf(node);
      if (id) used.add(id);
    });
  }
  const stamp = (node: IdNode): IdNode => {
    if (node.id && kept.whole.has(node.id)) return node;
    const copy: IdNode = { ...node };
    if (
      (copy.type === "paragraph" || copy.type === "table-row") &&
      copy.id &&
      !kept.head.has(copy.id) &&
      !paraIdOf(copy)
    ) {
      let salt = 0;
      let candidate = derivedParaId(copy.id);
      while (used.has(candidate)) candidate = derivedParaId(copy.id, ++salt);
      used.add(candidate);
      let attrs: Record<string, unknown> = {};
      try {
        attrs = copy.oxmlAttrs ? JSON.parse(copy.oxmlAttrs) : {};
      } catch {
        attrs = {};
      }
      copy.oxmlAttrs = JSON.stringify({ ...attrs, [PARA_ID_KEY]: candidate });
    }
    if (Array.isArray(node.children)) {
      copy.children = node.children.map((child) =>
        child && typeof child === "object" && "type" in child
          ? stamp(child as IdNode)
          : child,
      );
    }
    return copy;
  };
  return {
    ...doc,
    children: doc.children.map((block) =>
      !regenerated || (block.id && regenerated.has(block.id))
        ? (stamp(block as IdNode) as BlockNode)
        : block,
    ),
  };
}

export interface IdMatchReport {
  /** Ids carried over to the re-read document. */
  kept: number;
  /** New blocks (or ones with no provable identity): fresh ids. */
  fresh: number;
  /** Kept ids whose content changed: earlier reads of them are stale. */
  changedIds: string[];
  /** Old ids that no longer exist. */
  removedIds: string[];
}

/** Longest common subsequence of two signature lists (index pairs). */
function lcsIndexPairs(a: string[], b: string[]): Array<[number, number]> {
  const n = a.length;
  const m = b.length;
  if (n === 0 || m === 0) return [];
  if (n * m > 4_000_000) {
    // Too large for a table: match the common prefix and suffix only.
    const pairs: Array<[number, number]> = [];
    let i = 0;
    while (i < n && i < m && a[i] === b[i]) pairs.push([i, i++]);
    let k = 1;
    while (k <= n - i && k <= m - i && a[n - k] === b[m - k]) {
      pairs.push([n - k, m - k]);
      k += 1;
    }
    return pairs.sort((x, y) => x[0] - y[0]);
  }
  const table = new Uint32Array((n + 1) * (m + 1));
  for (let i = n - 1; i >= 0; i -= 1) {
    for (let j = m - 1; j >= 0; j -= 1) {
      table[i * (m + 1) + j] =
        a[i] === b[j]
          ? table[(i + 1) * (m + 1) + j + 1] + 1
          : Math.max(table[(i + 1) * (m + 1) + j], table[i * (m + 1) + j + 1]);
    }
  }
  const pairs: Array<[number, number]> = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      pairs.push([i, j]);
      i += 1;
      j += 1;
    } else if (table[(i + 1) * (m + 1) + j] >= table[i * (m + 1) + j + 1]) {
      i += 1;
    } else {
      j += 1;
    }
  }
  return pairs;
}

/** Carry ids from `previous` onto the re-imported `next` (mutated). */
export function matchIds(previous: DocRoot, next: DocRoot): IdMatchReport {
  const renamed = new Map<string, string>();
  const changed = new Set<string>();
  // Every id in `next` that is not carried over must not collide with an
  // old id that now means something else.
  const oldIds = new Set<string>();
  for (const block of previous.children) {
    walkIds(block as IdNode, (node) => node.id && oldIds.add(node.id));
  }
  // paraId → old node, for Word's ids and the ones revise derived.
  const byParaId = new Map<string, IdNode>();
  for (const block of previous.children) {
    walkIds(block as IdNode, (node) => {
      const explicit = paraIdOf(node);
      if (explicit && !byParaId.has(explicit)) byParaId.set(explicit, node);
    });
  }
  for (const block of previous.children) {
    walkIds(block as IdNode, (node) => {
      if (!node.id) return;
      const derived = derivedParaId(node.id);
      if (!byParaId.has(derived)) byParaId.set(derived, node);
    });
  }
  // Containers (tables, lists) have no paraId of their own: they are found
  // through their rows' and paragraphs' ids.
  const oldParent = new Map<IdNode, IdNode>();
  const indexParents = (node: IdNode) => {
    for (const child of childrenOf(node)) {
      oldParent.set(child, node);
      indexParents(child);
    }
  };
  previous.children.forEach((block) => indexParents(block as IdNode));
  const seenParaIds = new Set<string>();
  const duplicated = new Set<string>();
  for (const block of next.children) {
    walkIds(block as IdNode, (node) => {
      const pid = paraIdOf(node);
      if (!pid) return;
      if (seenParaIds.has(pid)) duplicated.add(pid);
      seenParaIds.add(pid);
    });
  }
  const claimed = new Set<IdNode>();
  let fresh = 0;

  const carrySubtree = (from: IdNode, to: IdNode) => {
    // Identical content: the structures match node for node.
    if (from.id && to.id) {
      renamed.set(to.id, from.id);
      to.id = from.id;
    }
    claimed.add(from);
    const a = childrenOf(from);
    const b = childrenOf(to);
    b.forEach((child, i) => {
      if (a[i] && a[i].type === child.type) carrySubtree(a[i], child);
    });
  };

  const matchLists = (olds: IdNode[], news: IdNode[]) => {
    const pairs = new Map<IdNode, IdNode>();
    // 1. Word's (or revise's) paragraph and row ids. An id that appears
    // twice in the new file (a copied paragraph) proves nothing.
    for (const node of news) {
      const pid = paraIdOf(node);
      if (pid && duplicated.has(pid)) continue;
      const old = pid ? byParaId.get(pid) : undefined;
      if (old && !claimed.has(old) && old.type === node.type) {
        pairs.set(node, old);
        claimed.add(old);
      }
    }
    // 1b. Containers, through a descendant's paraId.
    const oldSet = new Set(olds);
    for (const node of news) {
      if (pairs.has(node) || paraIdOf(node)) continue;
      let found: IdNode | undefined;
      walkIds(node, (inner) => {
        if (found || inner === node) return;
        const pid = paraIdOf(inner);
        let at = pid && !duplicated.has(pid) ? byParaId.get(pid) : undefined;
        while (at && !oldSet.has(at)) at = oldParent.get(at);
        if (at && !claimed.has(at) && at.type === node.type) found = at;
      });
      if (found) {
        pairs.set(node, found);
        claimed.add(found);
      }
    }
    // 2. Identical blocks, in order, among the rest.
    const restOld = olds.filter((node) => !claimed.has(node));
    const restNew = news.filter((node) => !pairs.has(node));
    const lcs = lcsIndexPairs(
      restOld.map(contentSignature),
      restNew.map(contentSignature),
    );
    for (const [i, j] of lcs) {
      if (restOld[i].type === restNew[j].type) {
        pairs.set(restNew[j], restOld[i]);
        claimed.add(restOld[i]);
      }
    }
    for (const node of news) {
      const old = pairs.get(node);
      if (!old) {
        walkIds(node, (n) => {
          if (n.id && n.type !== "text") fresh += 1;
          // A fresh id must never reuse an old id with another meaning.
          if (n.id && oldIds.has(n.id)) n.id = undefined;
        });
        continue;
      }
      if (contentSignature(old) === contentSignature(node)) {
        carrySubtree(old, node);
      } else {
        if (old.id && node.id) {
          renamed.set(node.id, old.id);
          node.id = old.id;
          changed.add(old.id);
        }
        matchLists(childrenOf(old), childrenOf(node));
      }
    }
  };
  matchLists(previous.children as IdNode[], next.children as IdNode[]);

  // Blocks that lost their id above get a fresh one.
  for (const block of next.children) {
    walkIds(block as IdNode, (node) => {
      if (node.type !== "text" && node.id === undefined && "id" in node) {
        node.id = freshId(oldIds);
      }
    });
  }
  // The graft capture is keyed by block id: follow the renames.
  const source = next.docxSource;
  if (source) {
    source.spans = source.spans.map((span) => ({
      ...span,
      blockIds: span.blockIds.map((id) => renamed.get(id) ?? id),
    }));
  }
  const nextIds = new Set<string>();
  for (const block of next.children) {
    walkIds(block as IdNode, (node) => node.id && nextIds.add(node.id));
  }
  return {
    kept: renamed.size,
    fresh,
    changedIds: [...changed],
    removedIds: [...oldIds].filter((id) => !nextIds.has(id)),
  };
}

function freshId(avoid: ReadonlySet<string>): string {
  const alphabet =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  for (;;) {
    let id = "";
    for (let i = 0; i < 6; i += 1) {
      id += alphabet[Math.floor(Math.random() * alphabet.length)];
    }
    if (!avoid.has(id)) return id;
  }
}

export interface SyncResult {
  ydoc: Y.Doc;
  pkg: Uint8Array;
  idsKept: number;
  idsNew: number;
  /** Kept ids whose content changed: earlier reads of them are stale. */
  changedIds: string[];
  /** Ids that no longer exist. */
  removedIds: string[];
  changedBlocks: number;
  changedStyles: string[];
}

/**
 * Make `bytes` (any supported format) the session's model: parse them, carry
 * block ids over where identity is provable (see matchIds), and
 * keep a .docx as the graft package. Used after an xml patch and when the
 * file changed outside revise.
 */
export async function relearn(
  ydoc: Y.Doc,
  bytes: Uint8Array,
  filename: string,
): Promise<SyncResult> {
  let next: DocRoot;
  try {
    next = await parseDocument(bytes, filename);
  } catch (error) {
    throw new ReviseError(
      `${filename} does not open: ${(error as Error).message}. Nothing was changed.`,
      "import_failed",
    );
  }
  const previous = ydocToDocument(ydoc);
  const ids = matchIds(previous, next);
  const { refreshDocxSourceSignatures, storeDocxCarrier } =
    await engineModules.docxGraft();
  refreshDocxSourceSignatures(next);
  const synced = createReviseYDoc();
  seedDocRootWithClaim(synced, next);
  if (next.docxSource?.sourceHash) {
    await storeDocxCarrier(synced, bytes, {
      hash: next.docxSource.sourceHash,
    });
  }
  const styleMap = (doc: DocRoot) =>
    new Map(
      (doc.stylesheet?.paragraphStyles ?? []).map((style) => [
        style.styleId,
        JSON.stringify(style),
      ]),
    );
  const stylesBefore = styleMap(previous);
  const changedStyles = [...styleMap(next)]
    .filter(([id, json]) => stylesBefore.get(id) !== json)
    .map(([id]) => id);
  return {
    ydoc: synced,
    pkg: bytes,
    idsKept: ids.kept,
    idsNew: ids.fresh,
    changedIds: ids.changedIds,
    removedIds: ids.removedIds,
    changedBlocks: ids.changedIds.length,
    changedStyles,
  };
}
