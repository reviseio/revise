/**
 * The Revise document library: open a document, read it, edit it with the
 * agent tools, review tracked changes, render pages, and write it back out,
 * all in process, with nothing saved unless asked. The `revise` CLI is a
 * thin shell over this module.
 *
 * One `ReviseDocument` runs one operation at a time: calls made while
 * another is in flight wait their turn, so edits, reviews, exports, and
 * reloads never interleave.
 */
import { engineModules } from "@reviseio/engine";
import * as Y from "yjs";
import {
  createReviseYDoc,
  createServerDocumentSession,
  decodeYDoc,
  encodeYDoc,
  exportDocument,
  fileToYDoc,
  installDomShims,
  parseDocument,
  ydocSourceFidelity,
  ydocToDocument,
  type ReviseServerExportFormat,
} from "@reviseio/engine";
import { seedDocRootWithClaim, seedYDocFromFile } from "@reviseio/engine";
import type {
  BlockNode,
  ImageFetchPolicy,
  ReviseDocumentMeasurementData,
  ReviseToolDefinition,
  ReviseJsonObject,
  ReviseToolArguments,
  ReviseToolName,
  ReviseToolResult,
  ServerDocumentSession,
} from "@reviseio/engine";
import { ReviseError } from "./errors";

/** How faithfully the document exports back to its original Word file. */
export type SourceFidelity = Awaited<ReturnType<typeof ydocSourceFidelity>>;
import { exportSessionDocx, relearn, type DocxExportOutcome } from "./docx";
import type { RenderOptions, RenderResult } from "./render";

export type DocumentMode = "suggesting" | "editing";
export type ExportFormat = ReviseServerExportFormat | "pdf";

/** Formats the library reads. */
export const READABLE = /\.(docx|md|markdown|html?|odt|rtf|txt)$/i;
/** Formats with no tracked changes: edits there apply directly. */
const PLAIN_FORMATS = /\.(md|markdown|html?|txt)$/i;

const EXTENSION_FORMATS: Record<string, ExportFormat> = {
  ".docx": "docx",
  ".pdf": "pdf",
  ".md": "markdown",
  ".markdown": "markdown",
  ".html": "html",
  ".htm": "html",
  ".txt": "txt",
};

/** The export format a file name implies (`out.pdf` → "pdf"). */
export function formatForPath(file: string): ExportFormat | undefined {
  const match = /\.[^./\\]+$/.exec(file);
  return match ? EXTENSION_FORMATS[match[0].toLowerCase()] : undefined;
}

/** A document to read: a path, or bytes with the file name they came from. */
export type DocumentSource = string | { bytes: Uint8Array; filename: string };

async function readSource(
  source: DocumentSource,
): Promise<{ bytes: Uint8Array; filename: string }> {
  if (typeof source !== "string") return source;
  const fs = await import("node:fs");
  if (!fs.existsSync(source)) {
    throw new ReviseError(`No such file: ${source}`, "not_found");
  }
  try {
    return { bytes: new Uint8Array(fs.readFileSync(source)), filename: source };
  } catch (error) {
    throw new ReviseError(
      `Could not read ${baseName(source)}: ${(error as Error).message}`,
      "read_failed",
      { cause: error },
    );
  }
}

async function importYDoc(bytes: Uint8Array, filename: string): Promise<Y.Doc> {
  try {
    return await fileToYDoc(bytes, filename);
  } catch (error) {
    throw new ReviseError(
      `${baseName(filename)} does not open: ${(error as Error).message}`,
      "import_failed",
      { cause: error },
    );
  }
}

function assertReadable(filename: string): void {
  if (!READABLE.test(filename)) {
    throw new ReviseError(
      `Cannot read ${baseName(filename)}. Supported: .docx, .md, .html, .odt, .rtf, .txt.`,
      "unsupported_format",
    );
  }
}

function baseName(file: string): string {
  return file.split(/[/\\]/).pop() ?? file;
}

function stem(file: string): string {
  return baseName(file).replace(/\.[^.]+$/, "");
}

export interface OpenOptions {
  /** Who edits and comments: shown as the author of tracked changes. */
  author?: string;
  /** "suggesting" records edits as tracked changes (the default for Word
   * files); "editing" applies them directly (the default for plain text). */
  mode?: DocumentMode;
  /** A stable id for this document (tool sessions, collaboration). */
  documentId?: string;
  /** How images inserted by URL are fetched. By default only public http(s)
   * hosts are reachable; see ImageFetchPolicy. */
  images?: ImageFetchPolicy;
}

export interface ChangeEntry {
  id: string;
  type: string;
  author: string;
  blockId?: string;
  blockIndex?: number;
  /** The heading of the section the change is in. */
  section?: string;
  deleted: string;
  inserted: string;
}

/** Narrows tracked changes by where they are. Filters combine: a change
 * must pass every one given. */
export interface ChangeFilter {
  /** A heading and everything under it, subsections included: its number
   * ("8", "8.", "8.1", including Word's automatic numbering) or its title
   * ("Allocation of risk", or part of it). */
  section?: string;
  /** Everything outside that section. */
  exceptSection?: string;
  /** Block ids, and inclusive `{ from, to }` ranges in document order. Ids
   * of list items, table cells, and blocks inside them work. */
  blocks?: Array<string | { from: string; to: string }>;
}

/** Which tracked changes to settle: exactly one of `ids`, `all`, or
 * filters (`section`, `exceptSection`, `blocks`). */
export interface ChangeSelection extends ChangeFilter {
  ids?: string[];
  all?: boolean;
}

export interface SettleResult {
  resolved: string[];
  /** Requested ids that are not pending tracked changes. */
  missing: string[];
  /** Pending changes that could not be settled. */
  unresolved: string[];
}

/** How pending tracked changes read: "tracked" shows them (deleted and
 * inserted text, marked), "accepted" as if every one were accepted (what
 * the editing tools see), "original" as if every one were rejected. */
export type ReadView = "tracked" | "accepted" | "original";

export interface GrepMatch {
  text: string;
  /** Inside text a pending change inserted or deleted; null for text no
   * change touches. */
  change: "inserted" | "deleted" | null;
  changeId: string | null;
}

export interface GrepHit {
  id: string;
  blockIndex: number;
  count: number;
  /** The block's text in the view. Tracked marks changes inline:
   * "Fees are [-$100-]{+$120+} per hour." */
  text: string;
  matches: GrepMatch[];
}

/** Where a block is: the body, a page header or footer, or a note. */
export type TextRegion = "body" | "header" | "footer" | "footnote" | "endnote";

export interface ReplaceAllOptions {
  /** Only in these blocks (ids from `read`, `grep`, `search`). */
  blocks?: string[];
  /** @deprecated The older name for `blocks`. */
  ids?: string[];
  /** Match regardless of case: "Provider" also finds "PROVIDER". */
  ignoreCase?: boolean;
  /** Match characters exactly. By default curly and straight quotes,
   * non-breaking and ordinary spaces, and non-breaking and ordinary
   * hyphens match each other, the way Word's Find treats quotes. */
  exact?: boolean;
}

export interface ReplaceTextResult {
  /** Occurrences replaced. */
  replaced: number;
  /** Ids of the blocks that contained the text. */
  blocks: string[];
  /** Occurrences replaced in each part of the document. */
  regions: Partial<Record<TextRegion, number>>;
  /** Tracked changes the replacement created (none in editing mode). */
  changeIds: string[];
  /** Blocks that contained the text but could not be edited. */
  failures: Array<{ blockId: string; code: string; message: string }>;
  /** Text that would match with looser options but wasn't replaced: other
   * capitalization (`ignoreCase`) or other quotes and spaces (`exact`). */
  nearMisses: Array<{
    text: string;
    count: number;
    reason: "case" | "characters";
  }>;
}

export interface TextDiff {
  /** Paragraph-level diff with word detail: `~ a [-old-]{+new+} b`. */
  text: string;
  changed: number;
}

export interface DocumentSummary {
  text: string;
  data: {
    pages?: number;
    words: number;
    pageSize: string;
    orientation: string;
    pageInches: [number, number];
    marginsInches: number[];
    title?: string;
    headings: Record<string, number>;
    lists: number[];
    tables: string[];
    images: number;
    headersFooters: Array<{ side: string; slot: string; text: string }>;
    pendingTrackedChanges: number;
  };
}

export interface ExportOptions {
  format: ExportFormat;
  /** "review" keeps tracked changes and comments; "final" accepts them.
   * Only .docx can hold tracked changes: PDF, Markdown, HTML, and text
   * always show them accepted. */
  mode?: "review" | "final";
  /** The same choice in read's words, and one more: "tracked" is
   * "review", "accepted" is "final", and "original" writes the document as
   * it was before every pending change. Give `mode` or `view`, not both. */
  view?: ReadView;
  title?: string;
}

export interface ExportResult {
  bytes: Uint8Array;
  format: ExportFormat;
  /** The mode the output actually has: only .docx can be "review". */
  mode: "review" | "final";
  /** The view the output shows: "tracked" only for a .docx. */
  view: ReadView;
  /** Tracked changes still pending in the output. */
  pendingTrackedChanges: number;
  /** For .docx, what the export actually did: "graft" carried untouched
   * content over byte for byte from the original file; "regenerated" wrote
   * the whole file fresh (fallbackCode says why). */
  fidelity: DocxExportOutcome | null;
  warnings: string[];
}

/** Bytes of a document in a format, through graft export when the original
 * Word package is available, with what the .docx export actually did. */
export async function exportYDoc(
  ydoc: Y.Doc,
  pkg: Uint8Array | null,
  { format, mode = "review", title = "Document" }: ExportOptions,
): Promise<{ bytes: Uint8Array; outcome: DocxExportOutcome | null }> {
  if (format === "pdf")
    return { bytes: await exportPdf(ydoc, title), outcome: null };
  if (format === "docx") return exportSessionDocx(ydoc, pkg, mode);
  return {
    bytes: await exportDocument(ydoc, { format, mode, title }),
    outcome: null,
  };
}

/** PDF through the editor's own PDF converter: the same pagination as the
 * editor and page rendering. Pending tracked changes are accepted, as in the
 * editor's PDF export. */
async function exportPdf(ydoc: Y.Doc, title: string): Promise<Uint8Array> {
  const [{ serializeDocumentFormat }, { DocumentMutator }] = await Promise.all([
    engineModules.documentFormats(),
    engineModules.mutator(),
  ]);
  const doc = DocumentMutator.materializeAcceptedSuggestions(
    ydocToDocument(ydoc),
  );
  const output = await serializeDocumentFormat(doc, "pdf", { title });
  return typeof output === "string"
    ? new TextEncoder().encode(output)
    : new Uint8Array(await output.arrayBuffer());
}

// ── snapshots ───────────────────────────────────────────────────────────────
//
// A snapshot is the complete document for storage or hand-off to the
// editor: a small versioned header (schema, identity, mode, author, title),
// the Y.js state, and the original Word package when there is one, so a
// restored document exports exactly as the original would have.

const SNAPSHOT_MAGIC = "RVSNAP";
const SNAPSHOT_SCHEMA = 1;

interface SnapshotHeader {
  schema: number;
  documentId: string;
  mode: DocumentMode;
  author?: string;
  title: string;
  stateBytes: number;
  packageBytes: number;
}

function encodeSnapshot(
  header: Omit<SnapshotHeader, "schema" | "stateBytes" | "packageBytes">,
  state: Uint8Array,
  pkg: Uint8Array | null,
): Uint8Array {
  const json = new TextEncoder().encode(
    JSON.stringify({
      ...header,
      schema: SNAPSHOT_SCHEMA,
      stateBytes: state.length,
      packageBytes: pkg?.length ?? 0,
    } satisfies SnapshotHeader),
  );
  const magic = new TextEncoder().encode(SNAPSHOT_MAGIC);
  const out = new Uint8Array(
    magic.length + 4 + json.length + state.length + (pkg?.length ?? 0),
  );
  out.set(magic, 0);
  new DataView(out.buffer).setUint32(magic.length, json.length);
  out.set(json, magic.length + 4);
  out.set(state, magic.length + 4 + json.length);
  if (pkg) out.set(pkg, magic.length + 4 + json.length + state.length);
  return out;
}

function decodeSnapshot(bytes: Uint8Array): {
  header: SnapshotHeader | null;
  state: Uint8Array;
  pkg: Uint8Array | null;
} {
  const magic = new TextDecoder().decode(
    bytes.subarray(0, SNAPSHOT_MAGIC.length),
  );
  if (magic !== SNAPSHOT_MAGIC) {
    // A bare Y.js update (Y.encodeStateAsUpdate): no header.
    return { header: null, state: bytes, pkg: null };
  }
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const jsonLength = view.getUint32(SNAPSHOT_MAGIC.length);
    const start = SNAPSHOT_MAGIC.length + 4;
    const header = JSON.parse(
      new TextDecoder().decode(bytes.subarray(start, start + jsonLength)),
    ) as SnapshotHeader;
    const valid =
      Number.isInteger(header.schema) &&
      header.schema >= 1 &&
      typeof header.documentId === "string" &&
      (header.mode === "suggesting" || header.mode === "editing") &&
      typeof header.title === "string" &&
      (header.author === undefined || typeof header.author === "string") &&
      Number.isInteger(header.stateBytes) &&
      header.stateBytes > 0 &&
      Number.isInteger(header.packageBytes) &&
      header.packageBytes >= 0 &&
      start + jsonLength + header.stateBytes + header.packageBytes ===
        bytes.length;
    if (!valid && !(header.schema > SNAPSHOT_SCHEMA)) throw new Error("header");
    if (header.schema > SNAPSHOT_SCHEMA) {
      throw new ReviseError(
        `This snapshot uses schema ${header.schema}; this version of revise reads up to ${SNAPSHOT_SCHEMA}. Upgrade revise.`,
        "snapshot_too_new",
      );
    }
    const stateStart = start + jsonLength;
    const state = bytes.subarray(stateStart, stateStart + header.stateBytes);
    const pkg = header.packageBytes
      ? bytes.subarray(
          stateStart + header.stateBytes,
          stateStart + header.stateBytes + header.packageBytes,
        )
      : null;
    if (
      state.length !== header.stateBytes ||
      (pkg && pkg.length !== header.packageBytes)
    ) {
      throw new Error("truncated");
    }
    return { header, state, pkg };
  } catch (error) {
    if (error instanceof ReviseError) throw error;
    throw new ReviseError(
      "This snapshot is damaged or incomplete.",
      "snapshot_corrupt",
      {
        cause: error,
      },
    );
  }
}

let documentCounter = 0;

/**
 * One open document: a live Revise document model plus the agent tool
 * session that edits it. Nothing touches disk unless you export it.
 */
export class ReviseDocument {
  private disposed = false;
  private queue: Promise<unknown> = Promise.resolve();

  private constructor(
    private currentYDoc: Y.Doc,
    private currentSession: ServerDocumentSession,
    private readonly settings: OpenOptions & {
      documentId: string;
      mode: DocumentMode;
      title: string;
    },
    private originalPackage: Uint8Array | null,
    /** False when the Y.Doc came from the caller (fromYDoc): it may be
     * shared with an editor or a provider, so it is never replaced. */
    private readonly ownsYDoc: boolean,
  ) {}

  /** Open a file (path) or bytes. Word files keep their original package,
   * so untouched content exports byte for byte. */
  static async open(
    source: DocumentSource,
    options: OpenOptions = {},
  ): Promise<ReviseDocument> {
    const { bytes, filename } = await readSource(source);
    assertReadable(filename);
    const ydoc = await importYDoc(bytes, filename);
    return ReviseDocument.build(ydoc, {
      mode: PLAIN_FORMATS.test(filename) ? "editing" : "suggesting",
      ...options,
      title: stem(filename),
      package: /\.docx$/i.test(filename) ? bytes : null,
      owned: true,
    });
  }

  /** Start a document from the HTML the tools read and write (headings,
   * `data-style-id="Title"`, lists, tables), or from Markdown. */
  static async create(
    content: { html?: string; markdown?: string; title?: string } = {},
    options: OpenOptions = {},
  ): Promise<ReviseDocument> {
    let ydoc: Y.Doc;
    if (content.html !== undefined) {
      await installDomShims();
      const { AgentHtmlConverter } = await engineModules.agentHtml();
      const body =
        /<body[^>]*>([\s\S]*)<\/body>/i.exec(content.html)?.[1] ?? content.html;
      ydoc = createReviseYDoc();
      seedDocRootWithClaim(ydoc, AgentHtmlConverter.fromAgentHtml(body));
    } else {
      const markdown =
        content.markdown ?? (content.title ? `# ${content.title}\n` : "\n");
      ydoc = await importYDoc(new TextEncoder().encode(markdown), "new.md");
    }
    return ReviseDocument.build(ydoc, {
      mode: "editing",
      ...options,
      title: content.title ?? "Document",
      package: null,
      owned: true,
    });
  }

  /** Restore a document from `snapshot()` (or a bare Y.js update). Options
   * override what the snapshot recorded. */
  static async restore(
    snapshot: Uint8Array,
    options: OpenOptions = {},
  ): Promise<ReviseDocument> {
    const { header, state, pkg } = decodeSnapshot(snapshot);
    let ydoc: Y.Doc;
    try {
      ydoc = decodeYDoc(state);
    } catch (error) {
      throw new ReviseError("This snapshot is damaged.", "snapshot_corrupt", {
        cause: error,
      });
    }
    return ReviseDocument.build(ydoc, {
      ...(header
        ? {
            documentId: header.documentId,
            mode: header.mode,
            title: header.title,
            ...(header.author ? { author: header.author } : {}),
          }
        : {}),
      ...options,
      package: pkg,
      owned: true,
    });
  }

  /** Wrap an existing Revise Y.Doc, e.g. one shared with an editor. The
   * Y.Doc stays yours: `dispose()` leaves it intact, and `reload()` is not
   * available (it would have to replace it). */
  static async fromYDoc(
    ydoc: Y.Doc,
    options: OpenOptions & { package?: Uint8Array | null; title?: string } = {},
  ): Promise<ReviseDocument> {
    // A room seeded from a Word file carries that file (seedYDoc stores it),
    // so untouched content exports byte for byte without passing it again.
    let pkg = options.package;
    if (pkg === undefined) {
      const { loadDocxCarrier } = await engineModules.docxGraft();
      pkg = loadDocxCarrier(ydoc)?.bytes ?? null;
    }
    return ReviseDocument.build(ydoc, {
      ...options,
      package: pkg,
      owned: false,
    });
  }

  private static async build(
    ydoc: Y.Doc,
    options: OpenOptions & {
      package?: Uint8Array | null;
      title?: string;
      owned: boolean;
    },
  ): Promise<ReviseDocument> {
    const { package: pkg, owned, ...rest } = options;
    const settings = {
      ...rest,
      documentId:
        rest.documentId ?? `revise-${Date.now()}-${++documentCounter}`,
      mode: rest.mode ?? "suggesting",
      title: rest.title ?? "Document",
    };
    return new ReviseDocument(
      ydoc,
      await ReviseDocument.startSession(ydoc, settings),
      settings,
      pkg ?? null,
      owned,
    );
  }

  private static async startSession(
    ydoc: Y.Doc,
    settings: OpenOptions & { documentId: string; mode: DocumentMode },
  ): Promise<ServerDocumentSession> {
    // A file is the source of truth: refuse tracked edits a .docx cannot
    // carry as revisions.
    (await engineModules.serverSession()).requireDocxRoundTrip(ydoc);
    return createServerDocumentSession(ydoc, {
      documentId: settings.documentId,
      mode: settings.mode,
      ...(settings.author ? { agentName: settings.author } : {}),
      ...(settings.images ? { images: settings.images } : {}),
    });
  }

  private assertOpen(): void {
    if (this.disposed) {
      throw new ReviseError("This document was disposed.", "disposed");
    }
  }

  /** Run `work` after every operation already queued on this document. */
  private serial<T>(work: () => Promise<T> | T): Promise<T> {
    if (this.disposed) {
      return Promise.reject(
        new ReviseError("This document was disposed.", "disposed"),
      );
    }
    const run = this.queue.then(() => {
      this.assertOpen();
      return work();
    });
    this.queue = run.catch(() => undefined);
    return run;
  }

  /** The live document model. Shared with `@reviseio/editor`. */
  get ydoc(): Y.Doc {
    return this.currentYDoc;
  }

  /** The low-level tool session behind `tools`: an escape hatch for what
   * the methods here don't cover. */
  get session(): ServerDocumentSession {
    return this.currentSession;
  }

  get mode(): DocumentMode {
    return this.currentSession.getMode();
  }

  /** Switch between tracked changes ("suggesting") and direct edits. */
  setMode(mode: DocumentMode): Promise<void> {
    return this.serial(() => {
      if (mode === "suggesting") this.currentSession.setSuggestingMode();
      else this.currentSession.setEditingMode();
      this.settings.mode = mode;
    });
  }

  get author(): string | undefined {
    return this.settings.author;
  }

  get documentId(): string {
    return this.settings.documentId;
  }

  /** The original Word package that export grafts onto, if any. */
  get package(): Uint8Array | null {
    return this.originalPackage;
  }

  /** The complete document (versioned; includes the original Word package)
   * to store or hand to `@reviseio/editor`, as of every operation queued
   * before this call. Restore with `restore()`. */
  snapshot(): Promise<Uint8Array> {
    return this.serial(() =>
      encodeSnapshot(
        {
          documentId: this.settings.documentId,
          mode: this.mode,
          ...(this.settings.author ? { author: this.settings.author } : {}),
          title: this.settings.title,
        },
        encodeYDoc(this.currentYDoc),
        this.originalPackage,
      ),
    );
  }

  /** An opaque value that changes whenever the document does (including
   * edits that only delete, which a state vector misses). */
  version(): Uint8Array {
    this.assertOpen();
    return Y.encodeSnapshot(Y.snapshot(this.currentYDoc));
  }

  /**
   * Re-read the document from new bytes (it changed elsewhere). Block ids
   * carry over only where identity is provable (Word's paragraph ids, or
   * identical content); `changedIds` lists kept ids whose content changed,
   * so earlier reads of them are stale. Replaces the Y.Doc, so it is not
   * available on a document made with `fromYDoc`.
   */
  reload(
    bytes: Uint8Array,
    filename: string,
  ): Promise<{
    idsKept: number;
    idsNew: number;
    changedIds: string[];
    removedIds: string[];
  }> {
    return this.serial(async () => {
      if (!this.ownsYDoc) {
        throw new ReviseError(
          "reload() would replace a Y.Doc you own (fromYDoc). Open a new document instead.",
          "borrowed_ydoc",
        );
      }
      const fresh = await relearn(this.currentYDoc, bytes, filename);
      const session = await ReviseDocument.startSession(
        fresh.ydoc,
        this.settings,
      );
      this.currentSession.dispose();
      this.currentYDoc.destroy();
      this.currentYDoc = fresh.ydoc;
      this.currentSession = session;
      if (/\.docx$/i.test(filename)) this.originalPackage = bytes;
      return {
        idsKept: fresh.idsKept,
        idsNew: fresh.idsNew,
        changedIds: fresh.changedIds,
        removedIds: fresh.removedIds,
      };
    });
  }

  /** Start the tool session over: forget what the agent has read and
   * searched (e.g. after the host compacts its transcript). */
  resetToolContext(): Promise<void> {
    return this.serial(async () => {
      const session = await ReviseDocument.startSession(
        this.currentYDoc,
        this.settings,
      );
      this.currentSession.dispose();
      this.currentSession = session;
    });
  }

  /** Release the tool session. The document can't be used afterwards; a
   * Y.Doc it created is destroyed, one passed to fromYDoc is left alone. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.currentSession.dispose();
    if (this.ownsYDoc) this.currentYDoc.destroy();
  }

  // ── tools ────────────────────────────────────────────────────────────────

  readonly tools = {
    /** Every editing tool's name, description, and JSON input schema. */
    definitions: (): ReviseToolDefinition[] => {
      this.assertOpen();
      return this.currentSession.tools.getDefinitions();
    },
    /** Run a known tool with typed input. Expected failures come back as
     * `{ ok: false, error }`; nothing throws for a bad edit. */
    call: <Name extends ReviseToolName>(
      name: Name,
      ...args: ReviseToolArguments<Name>
    ): Promise<ReviseToolResult<Name>> =>
      this.serial(async () =>
        this.afterTool(
          name,
          await this.currentSession.tools.execute(name, ...args),
        ),
      ),
    /** Run a tool by name with untyped input (agent tool calls). Unknown
     * tools and bad input come back as `{ ok: false, error }` too. */
    run: (
      name: string,
      input: ReviseJsonObject = {},
      options: { direct?: boolean } = {},
    ): Promise<ReviseToolResult<string>> =>
      this.serial(async () => {
        const definition = this.currentSession.tools
          .getDefinitions()
          .find((tool) => tool.name === name);
        if (!definition) {
          return {
            ok: false as const,
            error: {
              code: "unknown_tool",
              message: `Unknown tool "${name}".`,
            },
          } as unknown as ReviseToolResult<string>;
        }
        const required = (definition.inputSchema.required as string[]) ?? [];
        const filled =
          required.includes("context_notes") &&
          input.context_notes === undefined
            ? { ...input, context_notes: "N/A" }
            : input;
        return this.afterTool(
          name,
          await this.currentSession.tools.executeDynamic(
            name,
            filled,
            options.direct ? { directMode: true } : undefined,
          ),
        );
      }),
  };

  /** The document's title for exports: its saved title property (dc:title
   * in a .docx, or set_title's), else the name it was opened under. */
  private title(): string {
    const core = this.currentYDoc.getMap("metadata").get("coreProperties") as
      | { title?: unknown }
      | undefined;
    return typeof core?.title === "string" && core.title.trim()
      ? core.title
      : this.settings.title;
  }

  /** set_title renames the document in the editor. In a file, the title is
   * the document property (docProps/core.xml in a .docx, the PDF title), so
   * it is saved there too; without this, the tool reported success and the
   * saved file kept its old title. */
  private afterTool<Result extends { ok: boolean }>(
    name: string,
    result: Result,
  ): Result {
    if (name !== "set_title" || !result.ok) return result;
    const metadata = this.currentYDoc.getMap("metadata");
    const title = metadata.get("title");
    if (typeof title !== "string") return result;
    const core = metadata.get("coreProperties");
    this.currentYDoc.transact(() => {
      metadata.set("coreProperties", {
        ...(core && typeof core === "object" ? core : {}),
        title,
      });
    });
    this.settings.title = title;
    return result;
  }

  // ── reading ──────────────────────────────────────────────────────────────

  /** Sections with their block ranges and word counts. */
  outline(
    options: { limit?: number; offset?: number } = {},
  ): Promise<ReviseDocumentMeasurementData> {
    return this.serial(async () => {
      const measured = await this.currentSession.tools.call(
        "measure_document",
        {
          include_sections: true,
          section_limit: options.limit ?? 200,
          section_offset: options.offset ?? 0,
        },
      );
      return measured.data;
    });
  }

  /** Blocks, words, pending changes, and how faithfully it exports to Word. */
  status(): Promise<{
    blocks: number;
    words: number;
    pendingSuggestions: number;
    fidelity: SourceFidelity | null;
  }> {
    return this.serial(async () => {
      const measured = await this.currentSession.tools.call(
        "measure_document",
        {},
      );
      return {
        blocks: measured.data.totalBlocks,
        words: measured.data.wordCount,
        pendingSuggestions: this.currentSession.listSuggestions().length,
        fidelity: this.originalPackage
          ? await ydocSourceFidelity(this.currentYDoc)
          : null,
      };
    });
  }

  /** Text-bearing blocks matching a pattern (paragraphs, headings, list
   * items, table-cell paragraphs), with their ids. */
  grep(
    pattern: string | RegExp,
    options: { ignoreCase?: boolean; view?: ReadView } = {},
  ): GrepHit[] {
    this.assertOpen();
    const matcher =
      typeof pattern === "string"
        ? new RegExp(
            pattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
            options.ignoreCase ? "gi" : "g",
          )
        : new RegExp(
            pattern.source,
            pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`,
          );
    const view = options.view ?? "tracked";
    return textBlocks(ydocToDocument(this.currentYDoc))
      .map((block) => {
        const matches = matchesIn(block.segments, matcher, view);
        return {
          id: block.id,
          blockIndex: block.topIndex,
          count: matches.length,
          text: segmentsText(block.segments, view),
          matches,
        };
      })
      .filter((hit) => hit.count > 0);
  }

  /**
   * HTML the read tools returned (`read_blocks_from_index`,
   * `read_specific_blocks`), in a view. The tools read the accepted view;
   * "tracked" marks pending changes as `<ins>`/`<del>` with
   * `data-change-id` and `data-author`, and shows whole blocks a change
   * deletes where they were; "original" shows the text before every change.
   * Block ids are unchanged. The editing tools read `<ins>` as its text and
   * ignore `<del>`, so tracked HTML pasted back edits the accepted text.
   */
  viewHtml(html: string, view: ReadView): Promise<string> {
    if (view === "accepted") return Promise.resolve(html);
    return this.serial(async () => {
      await installDomShims();
      const { AgentHtmlConverter } = await engineModules.agentHtml();
      const raw = ydocToDocument(this.currentYDoc);
      const authors = new Map(
        this.currentSession
          .listSuggestions()
          .map((r) => [r.id, r.agentName ?? r.authorType ?? "unknown"]),
      );
      return renderView(html, raw, view, authors, AgentHtmlConverter);
    });
  }

  /** What the document holds, in one glance: page setup, header and footer
   * text with live fields, title, headings, lists, tables, review state. */
  summary(): Promise<DocumentSummary> {
    return this.serial(() =>
      documentSummary(this.currentYDoc, this.currentSession),
    );
  }

  // ── editing ──────────────────────────────────────────────────────────────

  /**
   * Replace every occurrence everywhere: the body, tables, lists, page
   * headers and footers, and footnotes and endnotes. Tracked changes in
   * suggesting mode. Curly and straight quotes and non-breaking spaces match
   * each other unless `exact`, and the replacement keeps the document's
   * characters (see matchDocumentCharacters). Replacing text with itself
   * changes nothing and is not a failure.
   */
  replaceAll(
    find: string,
    replace: string,
    options: ReplaceAllOptions = {},
  ): Promise<ReplaceTextResult> {
    return this.serial(async () => {
      if (!find) {
        throw new ReviseError("Nothing to find: give some text.", "usage");
      }
      const only = options.blocks ?? options.ids;
      const all = replaceTargets(ydocToDocument(this.currentYDoc)).filter(
        (block) => !only || only.includes(block.id),
      );
      const exactMatch = textMatcher(find, {
        exact: options.exact,
        ignoreCase: options.ignoreCase,
      });
      const loosest = textMatcher(find, { ignoreCase: true });
      const caseOnly = textMatcher(find, {
        exact: options.exact,
        ignoreCase: true,
      });
      const targets: Array<(typeof all)[number] & { matches: string[] }> = [];
      const missed = new Map<
        string,
        { count: number; reason: "case" | "characters" }
      >();
      for (const block of all) {
        const matches = exactMatch(block.text);
        if (matches.length) targets.push({ ...block, matches });
        // Near misses: what the loosest matching finds beyond these.
        const left = [...matches];
        for (const text of loosest(block.text)) {
          const at = left.indexOf(text);
          if (at >= 0) {
            left.splice(at, 1);
            continue;
          }
          // Missed only because of capitalization, or because of quotes
          // and spaces too.
          const reason = caseOnly(text).length ? "case" : "characters";
          const entry = missed.get(text) ?? { count: 0, reason };
          entry.count += 1;
          missed.set(text, entry);
        }
      }
      const nearMisses = [...missed].map(([text, m]) => ({ text, ...m }));
      const regions: ReplaceTextResult["regions"] = {};
      if (find === replace) {
        return {
          replaced: 0,
          blocks: targets.map((t) => t.id),
          regions,
          changeIds: [],
          failures: [],
          nearMisses,
        };
      }
      const failures: ReplaceTextResult["failures"] = [];
      const changeIds: string[] = [];
      let replaced = 0;
      for (const target of targets) {
        // The tools want a block read before it's edited. Header and note
        // paragraphs aren't readable that way but can be edited.
        await this.currentSession.tools.executeDynamic("read_specific_blocks", {
          block_ids: [target.id],
          context_notes: "N/A",
        });
        const distinct = [...new Set(target.matches)];
        const result = await this.currentSession.tools.executeDynamic(
          "replace",
          {
            id: target.id,
            replacements: distinct.map((text) => ({
              find: text,
              replace: matchDocumentCharacters(find, text, replace),
              occurrence: "all",
            })),
            context_notes: "N/A",
          },
        );
        if ("error" in result) {
          if (!/^No changes detected/i.test(result.error.message)) {
            failures.push({
              blockId: target.id,
              code: result.error.code,
              message: result.error.message,
            });
          }
        } else {
          replaced += target.matches.length;
          regions[target.region] =
            (regions[target.region] ?? 0) + target.matches.length;
          changeIds.push(...(result.value.suggestionIds ?? []));
        }
      }
      return {
        replaced,
        blocks: targets.map((t) => t.id),
        regions,
        changeIds,
        failures,
        nearMisses,
      };
    });
  }

  /** @deprecated The 3.0 name of `replaceAll`. */
  replaceText(
    find: string,
    replace: string,
    options: ReplaceAllOptions = {},
  ): Promise<ReplaceTextResult> {
    return this.replaceAll(find, replace, options);
  }

  /** The document as it was before every pending change, exported from
   * a copy: this document keeps its changes. */
  private async exportOriginal(options: ExportOptions): Promise<ExportResult> {
    const copy = await ReviseDocument.restore(await this.snapshot(), {
      author: this.settings.author,
    });
    try {
      await copy.reject({ all: true });
      const result = await copy.export({
        ...options,
        view: undefined,
        mode: "final",
      });
      return { ...result, view: "original" };
    } finally {
      copy.dispose();
    }
  }

  // ── review ───────────────────────────────────────────────────────────────

  /** Pending tracked changes with where they are and what they do. */
  changes(filter: ChangeFilter = {}): Promise<ChangeEntry[]> {
    return this.serial(() => this.filteredChanges(filter));
  }

  private async filteredChanges(
    filter: ChangeFilter = {},
  ): Promise<ChangeEntry[]> {
    return applyChangeFilter(
      await describeChanges(this.currentYDoc, this.currentSession),
      filter,
    );
  }

  /** The tracked changes a selection would settle, without settling them. */
  preview(selection: ChangeSelection): Promise<ChangeEntry[]> {
    return this.serial(async () => {
      const changes = await this.filteredChanges();
      const ids = await this.selectedIds(selection, changes);
      return ids === "all"
        ? changes
        : changes.filter((c) => ids.includes(c.id));
    });
  }

  accept(selection: ChangeSelection): Promise<SettleResult> {
    return this.serial(() => this.settle(selection, true));
  }

  reject(selection: ChangeSelection): Promise<SettleResult> {
    return this.serial(() => this.settle(selection, false));
  }

  private async selectedIds(
    selection: ChangeSelection,
    known?: ChangeEntry[],
  ): Promise<string[] | "all"> {
    const { section, exceptSection, blocks, all } = selection;
    const ids = selection.ids ?? [];
    const filtered = Boolean(section || exceptSection || blocks?.length);
    const kinds = [ids.length > 0, Boolean(all), filtered];
    if (kinds.filter(Boolean).length !== 1) {
      throw new ReviseError(
        "Select tracked changes one way: ids, all, or filters (section, exceptSection, blocks).",
        "invalid_selection",
      );
    }
    if (filtered) {
      return applyChangeFilter(
        known ?? (await describeChanges(this.currentYDoc, this.currentSession)),
        { section, exceptSection, blocks },
      ).map((c) => c.id);
    }
    return all ? "all" : ids;
  }

  private async settle(
    selection: ChangeSelection,
    accept: boolean,
  ): Promise<SettleResult> {
    const ids = await this.selectedIds(selection);
    const session = this.currentSession;
    if (ids === "all") {
      return accept
        ? session.acceptAllSuggestions()
        : session.rejectAllSuggestions();
    }
    if (!ids.length) return { resolved: [], missing: [], unresolved: [] };
    return accept
      ? session.acceptSuggestions(ids)
      : session.rejectSuggestions(ids);
  }

  /**
   * What the pending tracked changes do to the text (before → after), or
   * what changed since `baseline` (a snapshot).
   */
  diff(baseline?: Uint8Array): Promise<TextDiff> {
    return this.serial(async () => {
      let before: Y.Doc;
      if (baseline) {
        before = decodeYDoc(decodeSnapshot(baseline).state);
      } else {
        before = decodeYDoc(encodeYDoc(this.currentYDoc));
        const copy = await createServerDocumentSession(before, {
          documentId: `${this.settings.documentId}-diff`,
        });
        copy.rejectAllSuggestions();
        copy.dispose();
      }
      return textDiff(
        await markdownOf(before),
        await markdownOf(this.currentYDoc),
      );
    });
  }

  // ── output ───────────────────────────────────────────────────────────────

  /** Write the document in a format. Reports what the output actually
   * contains: its pending changes, and for .docx how much of the original
   * file was carried over byte for byte. */
  async export(options: ExportOptions): Promise<ExportResult> {
    const view = viewOfMode(options);
    if (view === "original") return this.exportOriginal(options);
    return this.serial(async () => {
      options = {
        ...options,
        format: exportFormat(options.format),
        mode: view === "accepted" ? "final" : options.mode,
      };
      // Only a .docx can hold tracked changes; every other format shows
      // them accepted.
      const mode =
        options.format === "docx" ? (options.mode ?? "review") : "final";
      const pending = this.currentSession.listSuggestions().length;
      const { bytes, outcome } = await exportYDoc(
        this.currentYDoc,
        this.originalPackage,
        { title: this.title(), ...options, mode },
      );
      const warnings: string[] = [];
      if (
        options.format !== "docx" &&
        (options.mode ?? "review") === "review" &&
        pending
      ) {
        warnings.push(
          `A .${options.format === "markdown" ? "md" : options.format} file can't hold tracked changes: it shows the ${pending} pending change${pending === 1 ? "" : "s"} accepted.`,
        );
      }
      if (outcome?.method === "regenerated") {
        if (outcome.fallbackCode === "no-original-package") {
          warnings.push(
            "The original Word file isn't available (e.g. restored from a bare Y.js state), so the whole .docx was written fresh.",
          );
        } else if (outcome.fallbackCode !== "not-from-docx") {
          warnings.push(
            `Couldn't carry untouched content over from the original Word file (${outcome.fallbackCode}); the whole .docx was written fresh.`,
          );
        }
      }
      return {
        bytes,
        format: options.format,
        mode,
        view: mode === "final" ? ("accepted" as const) : ("tracked" as const),
        pendingTrackedChanges: mode === "final" ? 0 : pending,
        fidelity: outcome,
        warnings,
      };
    });
  }

  /** Pages as images, laid out exactly as the editor lays them out. */
  renderPages(options: RenderOptions = {}): Promise<RenderResult> {
    return this.serial(async () => {
      try {
        const { renderPages } = await import("./render");
        return await renderPages(
          this.currentYDoc,
          this.settings.documentId,
          options,
        );
      } catch (error) {
        if (error instanceof ReviseError) throw error;
        throw new ReviseError(
          `Rendering failed: ${(error as Error).message}`,
          "render_failed",
          { cause: error },
        );
      }
    });
  }
}

// ── compare and convert ────────────────────────────────────────────────────

export interface CompareResult {
  /** The revised document with every difference as a tracked change. */
  document: ReviseDocument;
  stats: {
    inlineChanges: number;
    insertedBlocks: number;
    deletedBlocks: number;
    unmarkedBlocks: number;
  };
  untrackedRegions: string[];
  /** Accepting every change in the exported redline (read back from the
   * .docx) gives the revised text; rejecting every change gives the
   * original. Text only: formatting and comments are not compared. */
  verified: { acceptAllGivesRevised: boolean; rejectAllGivesOriginal: boolean };
}

/** A redline: `revised` with every difference from `original` tracked. */
/**
 * Seed a collaboration room (a Y.Doc your sync server owns) from a file, and
 * say whether it wrote anything. A room that already holds a document is
 * left alone, so call it on every connection:
 *
 *   onLoadDocument: async ({ document }) => {
 *     await seedYDoc(document, "contract.docx");
 *   }
 *
 * Two peers that both seeded an empty room before syncing are repaired to a
 * single copy. A Word file is stored in the room too, so any session over
 * it (`ReviseDocument.fromYDoc`, or the editor) exports untouched content
 * byte for byte.
 */
export async function seedYDoc(
  ydoc: Y.Doc,
  source: DocumentSource,
): Promise<boolean> {
  const { bytes, filename } = await readSource(source);
  assertReadable(filename);
  return seedYDocFromFile(ydoc, bytes, filename);
}

export async function compare(
  originalSource: DocumentSource,
  revisedSource: DocumentSource,
  options: OpenOptions = {},
): Promise<CompareResult> {
  const [originalFile, revisedFile] = await Promise.all([
    readSource(originalSource),
    readSource(revisedSource),
  ]);
  assertReadable(originalFile.filename);
  assertReadable(revisedFile.filename);
  const [original, revised] = await Promise.all([
    parseDocument(originalFile.bytes, originalFile.filename),
    parseDocument(revisedFile.bytes, revisedFile.filename),
  ]);
  const author = options.author ?? baseName(revisedFile.filename);
  const { compareDocuments } = await engineModules.compare();
  const comparison = compareDocuments(original, revised, {
    authorName: author,
  });
  if (!comparison) {
    throw new ReviseError(
      "These documents are too large to compare.",
      "too_large",
    );
  }
  // The comparison is the revised document with tracked changes, so the
  // revised file's package is the graft carrier: untouched content exports
  // as is.
  const ydoc = createReviseYDoc();
  seedDocRootWithClaim(ydoc, comparison.doc);
  const isDocx = /\.docx$/i.test(revisedFile.filename);
  if (comparison.doc.docxSource?.sourceHash && isDocx) {
    const { storeDocxCarrier } = await engineModules.docxGraft();
    await storeDocxCarrier(ydoc, revisedFile.bytes, {
      hash: comparison.doc.docxSource.sourceHash,
    });
  }
  const document = await ReviseDocument.restore(
    encodeSnapshot(
      {
        documentId: options.documentId ?? `revise-compare-${Date.now()}`,
        mode: options.mode ?? "suggesting",
        author,
        title: stem(revisedFile.filename),
      },
      encodeYDoc(ydoc),
      isDocx ? revisedFile.bytes : null,
    ),
    options.images ? { images: options.images } : {},
  );

  const plainOf = async (y: Y.Doc) =>
    new TextDecoder()
      .decode(await exportDocument(y, { format: "txt", mode: "final" }))
      .replace(/\s+/g, " ")
      .trim();
  // Check the file the user will get, not the model: export the redline,
  // read it back, and accept/reject everything in what came back.
  const redline = isDocx
    ? await fileToYDoc(
        (await exportSessionDocx(document.ydoc, document.package)).bytes,
        "redline.docx",
      )
    : decodeYDoc(encodeYDoc(document.ydoc));
  const rejected = decodeYDoc(encodeYDoc(redline));
  const rejecter = await createServerDocumentSession(rejected, {
    documentId: `${document.documentId}-verify`,
  });
  rejecter.rejectAllSuggestions();
  rejecter.dispose();
  const [acceptedText, rejectedText, revisedText, originalText] =
    await Promise.all([
      plainOf(redline),
      plainOf(rejected),
      plainOf(await fileToYDoc(revisedFile.bytes, revisedFile.filename)),
      plainOf(await fileToYDoc(originalFile.bytes, originalFile.filename)),
    ]);
  return {
    document,
    stats: comparison.stats,
    untrackedRegions: comparison.untrackedRegions,
    verified: {
      acceptAllGivesRevised: acceptedText === revisedText,
      rejectAllGivesOriginal: rejectedText === originalText,
    },
  };
}

/** Convert a document to another format in one step. */
export async function convert(
  source: DocumentSource,
  format: ExportFormat,
  options: { mode?: "review" | "final"; view?: ReadView } = {},
): Promise<Uint8Array> {
  format = exportFormat(format);
  const view = viewOfMode(options);
  if (view === "original") {
    // Reject every change in a document, then write what is left.
    const doc = await ReviseDocument.open(source);
    try {
      return (await doc.export({ format, view })).bytes;
    } finally {
      doc.dispose();
    }
  }
  if (view === "accepted") options = { mode: "final" };
  const { bytes, filename } = await readSource(source);
  assertReadable(filename);
  const ydoc = await importYDoc(bytes, filename);
  return format === "pdf"
    ? exportPdf(ydoc, stem(filename))
    : exportDocument(ydoc, {
        format,
        mode: options.mode ?? "review",
        title: stem(filename),
      });
}

// ── internals ──────────────────────────────────────────────────────────────

/** The view an export asks for, from `view` or the older `mode`. */
function viewOfMode(options: {
  mode?: "review" | "final";
  view?: ReadView;
}): ReadView | undefined {
  const { mode, view } = options;
  if (
    view !== undefined &&
    !["tracked", "accepted", "original"].includes(view)
  ) {
    throw new ReviseError(
      `Unknown view "${view}". Use tracked, accepted, or original.`,
      "usage",
    );
  }
  if (view && mode && (mode === "final") !== (view !== "tracked")) {
    throw new ReviseError(
      `mode "${mode}" and view "${view}" disagree; give one of them.`,
      "usage",
    );
  }
  return view;
}

const EXPORT_FORMATS = ["docx", "pdf", "markdown", "html", "txt"];

/** An export format by name; "md" is markdown. An unknown name is refused
 * by name (the converter's own message called markdown "import-only"). */
function exportFormat(format: ExportFormat): ExportFormat {
  const name = String(format).toLowerCase();
  if (name === "md") return "markdown" as ExportFormat;
  if (!EXPORT_FORMATS.includes(name)) {
    throw new ReviseError(
      `Unknown export format "${format}". Use one of: ${EXPORT_FORMATS.join(", ")}.`,
      "unsupported_format",
    );
  }
  return name as ExportFormat;
}

/** Text-bearing blocks a tool can target: paragraphs and headings, list
 * items, and paragraphs inside table cells, each with its own id. */
/** A run of a block's text: unchanged, or inside a pending change. */
interface Segment {
  text: string;
  change: "inserted" | "deleted" | null;
  changeId: string | null;
}

function textBlocks(
  doc: ReturnType<typeof ydocToDocument>,
): Array<{ id: string; segments: Segment[]; topIndex: number }> {
  const out: Array<{ id: string; segments: Segment[]; topIndex: number }> = [];
  const segmentsOf = (node: { children?: unknown[] }) => {
    const segments: Segment[] = [];
    for (const child of node.children ?? []) {
      const c = child as {
        type?: string;
        text?: string;
        marks?: {
          deletion?: boolean;
          insertion?: boolean;
          suggestionId?: string;
        };
      };
      if (c.type !== "text" || !c.text) continue;
      const change = c.marks?.deletion
        ? "deleted"
        : c.marks?.insertion
          ? "inserted"
          : null;
      const changeId = change ? (c.marks?.suggestionId ?? null) : null;
      const last = segments[segments.length - 1];
      if (last && last.change === change && last.changeId === changeId) {
        last.text += c.text;
      } else {
        segments.push({ text: c.text, change, changeId });
      }
    }
    return segments;
  };
  doc.children.forEach((block, topIndex) => {
    const visit = (node: unknown) => {
      const n = node as { type?: string; id?: string; children?: unknown[] };
      if (!n || typeof n !== "object") return;
      if ((n.type === "paragraph" || n.type === "list-item") && n.id) {
        out.push({ id: n.id, segments: segmentsOf(n), topIndex });
        return;
      }
      n.children?.forEach(visit);
    };
    visit(block);
  });
  return out;
}

/** Every paragraph replace-all can edit, with its text (pending deletions
 * left out) and where it is: body (tables and lists included), page
 * headers and footers, then footnotes and endnotes. */
function replaceTargets(
  doc: ReturnType<typeof ydocToDocument>,
): Array<{ id: string; text: string; region: TextRegion }> {
  const out: Array<{ id: string; text: string; region: TextRegion }> = [];
  const textOf = (node: { children?: unknown[] }) =>
    (node.children ?? [])
      .map((child) => {
        const c = child as {
          type?: string;
          text?: string;
          marks?: { deletion?: boolean };
        };
        return c.type === "text" && !c.marks?.deletion ? (c.text ?? "") : "";
      })
      .join("");
  const visit = (node: unknown, region: TextRegion) => {
    const n = node as { type?: string; id?: string; children?: unknown[] };
    if (!n || typeof n !== "object") return;
    if ((n.type === "paragraph" || n.type === "list-item") && n.id) {
      out.push({ id: n.id, text: textOf(n), region });
      return;
    }
    n.children?.forEach((child) => visit(child, region));
  };
  doc.children.forEach((block) => visit(block, "body"));
  for (const block of doc.chrome ?? []) {
    block.children.forEach((column) => visit(column, block.side));
  }
  for (const note of doc.notes ?? []) {
    const region = note.noteKind === "endnote" ? "endnote" : "footnote";
    note.children.forEach((child) => visit(child, region));
  }
  return out;
}

/** Characters that match each other unless matching is exact. */
const LOOSE_CHARACTERS: Record<string, string> = {
  "\u2018": "'", // ‘
  "\u2019": "'", // ’
  "\u201a": "'", // ‚
  "\u201b": "'", // ‛
  "\u2032": "'", // ′
  "\u201c": '"', // “
  "\u201d": '"', // ”
  "\u201e": '"', // „
  "\u201f": '"', // ‟
  "\u2033": '"', // ″
  "\u00a0": " ", // no-break space
  "\u202f": " ", // narrow no-break space
  "\u2007": " ", // figure space
  "\u2009": " ", // thin space
  "\u2011": "-", // non-breaking hyphen
  "\u2010": "-", // hyphen
};

/** A one-to-one character folding, so matches map back to the original
 * text at the same offsets. */
function fold(
  text: string,
  options: { exact?: boolean; ignoreCase?: boolean },
) {
  return Array.from(text, (char) => {
    let c = options.exact ? char : (LOOSE_CHARACTERS[char] ?? char);
    if (options.ignoreCase) {
      const lower = c.toLowerCase();
      if (lower.length === c.length) c = lower;
    }
    return c;
  });
}

/** Finds `find` in a text: each match as the text actually there. */
function textMatcher(
  find: string,
  options: { exact?: boolean; ignoreCase?: boolean },
): (text: string) => string[] {
  const needle = fold(find, options).join("");
  return (text) => {
    const chars = Array.from(text);
    const hay = fold(text, options).join("");
    const matches: string[] = [];
    // fold keeps one entry per code point; map string offsets to them.
    const starts: number[] = [];
    let offset = 0;
    for (const c of fold(text, options)) {
      starts.push(offset);
      offset += c.length;
    }
    let from = 0;
    while (needle) {
      const at = hay.indexOf(needle, from);
      if (at < 0) break;
      const first = starts.indexOf(at);
      let last = first;
      while (last < starts.length && starts[last] < at + needle.length) last++;
      if (first >= 0) matches.push(chars.slice(first, last).join(""));
      from = at + Math.max(needle.length, 1);
    }
    return matches;
  };
}

/**
 * The replacement in the document's characters: where `find` had a plain
 * quote, space, or hyphen and the document a curly quote, no-break space,
 * or no-break hyphen, the replacement's plain ones of that kind follow the
 * document's, in order, when it has as many. `"net 30"` → `"net 45"` keeps
 * the no-break space; `'"Acme"'` → `'"Beta"'` keeps the curly quotes.
 */
function matchDocumentCharacters(
  find: string,
  matched: string,
  replace: string,
): string {
  const findChars = Array.from(find);
  const matchedChars = Array.from(matched);
  const out = Array.from(replace);
  for (const plain of ["'", '"', " ", "-"]) {
    const documentChars = findChars
      .map((c, i) => (c === plain ? matchedChars[i] : undefined))
      .filter((c): c is string => c !== undefined);
    if (documentChars.every((c) => c === plain)) continue;
    const slots = out.flatMap((c, i) => (c === plain ? [i] : []));
    if (slots.length !== documentChars.length) continue;
    slots.forEach((slot, n) => (out[slot] = documentChars[n]));
  }
  return out.join("");
}

const shows = (segment: Segment, view: ReadView) =>
  view === "tracked" ||
  segment.change === null ||
  (segment.change === "inserted") === (view === "accepted");

function segmentsText(segments: Segment[], view: ReadView): string {
  return segments
    .filter((segment) => shows(segment, view))
    .map((segment) =>
      view !== "tracked" || segment.change === null
        ? segment.text
        : segment.change === "deleted"
          ? `[-${segment.text}-]`
          : `{+${segment.text}+}`,
    )
    .join("");
}

/**
 * Matches in a block's text, in document order. A view's text is its
 * segments joined; tracked matches against both the accepted and the
 * original text, so a match is found whichever side of a change it is on,
 * and text no change touches counts once.
 */
function matchesIn(
  segments: Segment[],
  matcher: RegExp,
  view: ReadView,
): GrepMatch[] {
  const found: Array<GrepMatch & { at: number }> = [];
  const scan = (
    side: "accepted" | "original",
    keep: (change: Segment["change"]) => boolean,
  ) => {
    // Each character of the side's text, by its place in all the segments.
    const places: number[] = [];
    const owners: Segment[] = [];
    let text = "";
    let at = 0;
    for (const segment of segments) {
      if (shows(segment, side)) {
        for (let i = 0; i < segment.text.length; i++) {
          places.push(at + i);
          owners.push(segment);
        }
        text += segment.text;
      }
      at += segment.text.length;
    }
    for (const match of text.matchAll(matcher)) {
      if (!match[0]) continue;
      const start = match.index ?? 0;
      const touched = owners
        .slice(start, start + match[0].length)
        .find((owner) => owner.change !== null);
      const change = touched?.change ?? null;
      if (!keep(change)) continue;
      found.push({
        text: match[0],
        change,
        changeId: touched?.changeId ?? null,
        at: places[start],
      });
    }
  };
  if (view === "tracked") {
    // Unchanged text is on both sides: take it from the accepted one.
    scan("accepted", (change) => change !== "deleted");
    scan("original", (change) => change === "deleted");
  } else {
    scan(view, () => true);
  }
  return (
    found
      // At the same place, deleted text reads first, as in a redline.
      .sort(
        (a, b) =>
          a.at - b.at ||
          Number(b.change === "deleted") - Number(a.change === "deleted"),
      )
      .map(({ text, change, changeId }) => ({ text, change, changeId }))
  );
}

// Sentinels around changed text while rendering: private-use characters
// pass through the HTML serializer untouched.
const OPEN = "";
const DELETED = "";
const INSERTED = "";
const CLOSE = "";
const CHANGE_BASE = 0xe100;
const SENTINEL_RUN = /([])([-])([\s\S]*?)/g;

type AgentHtmlModule = Awaited<ReturnType<typeof engineModules.agentHtml>>;

function renderView(
  html: string,
  raw: ReturnType<typeof ydocToDocument>,
  view: "tracked" | "original",
  authors: Map<string, string>,
  AgentHtmlConverter: AgentHtmlModule["AgentHtmlConverter"],
): string {
  const ids = topLevelIds(html);
  const blocks = [...raw.children, ...(raw.notes ?? [])] as BlockNode[];
  const position = new Map(blocks.map((block, i) => [block.id, i] as const));
  const known = ids
    .map((id) => position.get(id))
    .filter((at): at is number => at !== undefined);
  if (!known.length) return html;
  // Whole blocks a pending change deletes are missing from the tools'
  // (accepted) view; show them where they were.
  const deleted = (block: BlockNode | undefined) =>
    block !== undefined && wholeBlockDeletion(block) !== null;
  let first = Math.min(...known);
  let last = Math.max(...known);
  if (blocks.slice(0, first).every((block) => deleted(block))) first = 0;
  while (last + 1 < blocks.length && deleted(blocks[last + 1])) last += 1;

  const changes: string[] = [];
  const indexOf = (id: string) => {
    let i = changes.indexOf(id);
    if (i < 0) i = changes.push(id) - 1;
    return String.fromCharCode(CHANGE_BASE + Math.min(i, 0xf8ff - CHANGE_BASE));
  };
  // Text in a block a change deletes is deleted text, marked or not.
  const marked = (node: unknown, deletedBy?: string): unknown => {
    if (!node || typeof node !== "object") return node;
    const n = node as {
      type?: string;
      text?: string;
      marks?: {
        deletion?: boolean;
        insertion?: boolean;
        suggestionId?: string;
      };
      children?: unknown[];
    };
    if (n.type === "text" && n.text) {
      const deletion = n.marks?.deletion || (deletedBy && !n.marks?.insertion);
      if (!deletion && !n.marks?.insertion) return node;
      const id = n.marks?.suggestionId ?? deletedBy ?? "";
      const kind = deletion ? DELETED : INSERTED;
      return { ...n, text: `${OPEN}${kind}${indexOf(id)}${n.text}${CLOSE}` };
    }
    const structural = node as {
      structuralSuggestionId?: string;
      structuralSuggestionKind?: string;
    };
    const by =
      structural.structuralSuggestionKind === "delete"
        ? structural.structuralSuggestionId
        : deletedBy;
    return Array.isArray(n.children)
      ? { ...n, children: n.children.map((child) => marked(child, by)) }
      : node;
  };
  const range = blocks
    .slice(first, last + 1)
    .map((block) => marked(block)) as BlockNode[];
  const doc = { ...raw, children: range, notes: undefined } as typeof raw;
  const rendered = AgentHtmlConverter.extractSpecificBlocks(
    doc,
    range.map((block) => block.id),
    undefined,
    { scanDocumentFontFamily: false },
  ).foundBlocks.map((block, i) => {
    const whole = view === "tracked" ? wholeBlockDeletion(range[i]) : null;
    // A deleted block with no text (an image, an empty paragraph) still
    // says so.
    return whole && !block.html.includes(OPEN)
      ? block.html.replace(
          /^<([a-z][\w-]*)/,
          `<$1 data-change="deleted" data-change-id="${escapeAttr(whole)}" data-author="${escapeAttr(authors.get(whole) ?? "unknown")}"`,
        )
      : block.html;
  });
  // The tools head their HTML with "Showing N block(s):"; keep it true.
  const header = /^\s*Showing \d+ block\(s\):\s*/.exec(html)
    ? `Showing ${rendered.length} block(s):\n\n`
    : "";
  return (header + rendered.join("\n")).replace(
    SENTINEL_RUN,
    (_, kind: string, index: string, text: string) => {
      const deletion = kind === DELETED;
      if (view === "original") return deletion ? text : "";
      const id = changes[index.charCodeAt(0) - CHANGE_BASE] ?? "";
      const tag = deletion ? "del" : "ins";
      return `<${tag} data-change-id="${escapeAttr(id)}" data-author="${escapeAttr(authors.get(id) ?? "unknown")}">${text}</${tag}>`;
    },
  );
}

/** The change deleting a whole block, if one does. */
function wholeBlockDeletion(block: BlockNode): string | null {
  const b = block as {
    structuralSuggestionId?: string;
    structuralSuggestionKind?: string;
  };
  if (b.structuralSuggestionKind === "delete" && b.structuralSuggestionId) {
    return b.structuralSuggestionId;
  }
  // Or every character in it is deleted text.
  let id: string | null = null;
  let all = true;
  let any = false;
  const visit = (node: unknown) => {
    if (!all || !node || typeof node !== "object") return;
    const n = node as {
      type?: string;
      text?: string;
      marks?: { deletion?: boolean; suggestionId?: string };
      children?: unknown[];
    };
    if (n.type === "text") {
      if (!n.text) return;
      any = true;
      if (!n.marks?.deletion) all = false;
      else id ??= n.marks.suggestionId ?? null;
      return;
    }
    n.children?.forEach(visit);
  };
  visit(block);
  return any && all ? id : null;
}

/** Ids of the top-level elements in a run of block HTML. */
function topLevelIds(html: string): string[] {
  const body = new DOMParser().parseFromString(
    `<!doctype html><body>${html}</body>`,
    "text/html",
  ).body;
  return Array.from(body.children)
    .map((el) => el.getAttribute("id"))
    .filter((id): id is string => Boolean(id));
}

const escapeAttr = (value: string) =>
  value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/** Where a change sits: its innermost block's place in document order. */
interface ChangePlace {
  /** Pre-order position of the innermost block holding the change. */
  at: number;
  /** Ids of the blocks holding it, outermost first. */
  path: string[];
}

/** Every block (list items, table cells and their paragraphs included) in
 * document order: its pre-order position and the last position inside it. */
interface DocumentPlaces {
  span: Map<string, { start: number; end: number }>;
  sections: Array<{ title: string; level: number; start: number; end: number }>;
  changes: WeakMap<ChangeEntry, ChangePlace>;
}

const placesByEntries = new WeakMap<ChangeEntry[], DocumentPlaces>();

/**
 * Every pending tracked change with where it is and what it does: the
 * top-level block, the enclosing section heading, and the deleted and
 * inserted text, so changes can be picked by meaning instead of id.
 */
async function describeChanges(
  ydoc: Y.Doc,
  session: ServerDocumentSession,
): Promise<ChangeEntry[]> {
  const records = session.listSuggestions();
  const doc = ydocToDocument(ydoc);
  const { documentHeadingSections } = await engineModules.outline();

  // Positions come from the model itself, pending deletions included:
  // measurement indexes skip those, so they don't line up with it.
  const span = new Map<string, { start: number; end: number }>();
  let next = 0;
  const number = (node: unknown) => {
    if (!node || typeof node !== "object") return;
    const n = node as { id?: unknown; type?: string; children?: unknown[] };
    const isBlock = typeof n.id === "string" && n.type !== "text";
    const start = next;
    if (isBlock) next += 1;
    n.children?.forEach(number);
    if (isBlock) span.set(n.id as string, { start, end: next - 1 });
  };
  doc.children.forEach(number);
  // Headers, footers, and notes come after the body, outside every section.
  const bodyEnd = next;
  (doc.chrome ?? []).forEach(number);
  (doc.notes ?? []).forEach(number);

  const headings = documentHeadingSections(doc, { maxDepth: 3 })
    .flatMap((h) => {
      const start = span.get(h.itemId ?? h.blockId)?.start;
      return start === undefined
        ? []
        : [{ title: h.title, level: h.level as number, start }];
    })
    .sort((a, b) => a.start - b.start);
  // A section runs to the next heading at its level or above, so it
  // includes its subsections.
  const sections = headings.map((h, i) => ({
    ...h,
    end:
      (headings.slice(i + 1).find((later) => later.level <= h.level)?.start ??
        bodyEnd) - 1,
  }));
  const sectionAt = (at: number) => {
    let title: string | undefined;
    for (const section of sections) {
      if (section.start > at) break;
      title = section.title;
    }
    return title;
  };

  const positionOf = new Map(
    doc.children.map((block, index) => [block.id, index] as const),
  );
  const found = new Map<
    string,
    {
      blockIndex?: number;
      blockId: string;
      /** "Header", "Footnote 2": where a change outside the body is. */
      label?: string;
      path: string[];
      deleted: string;
      inserted: string;
    }
  >();
  const collect = (
    block: { id?: string },
    blockIndex: number | undefined,
    label: string | undefined,
  ) => {
    const visit = (node: unknown, path: string[]) => {
      if (!node || typeof node !== "object") return;
      const n = node as {
        id?: unknown;
        type?: string;
        text?: string;
        marks?: {
          suggestionId?: string;
          insertion?: boolean;
          deletion?: boolean;
        };
        suggestionId?: string;
        structuralSuggestionId?: string;
        children?: unknown[];
      };
      const here =
        typeof n.id === "string" && n.type !== "text" ? [...path, n.id] : path;
      const id =
        n.marks?.suggestionId ?? n.suggestionId ?? n.structuralSuggestionId;
      if (id) {
        const entry = found.get(id) ?? {
          blockIndex,
          blockId: block.id ?? "",
          label,
          path: here,
          deleted: "",
          inserted: "",
        };
        if (n.type === "text") {
          if (n.marks?.deletion) entry.deleted += n.text ?? "";
          else if (n.marks?.insertion) entry.inserted += n.text ?? "";
        }
        found.set(id, entry);
      }
      n.children?.forEach((child) => visit(child, here));
    };
    visit(block, []);
  };
  doc.children.forEach((block, blockIndex) =>
    collect(block, blockIndex, undefined),
  );
  for (const block of doc.chrome ?? []) {
    collect(block, undefined, block.side === "header" ? "Header" : "Footer");
  }
  const noteCount = { footnote: 0, endnote: 0 };
  for (const note of doc.notes ?? []) {
    const kind = note.noteKind === "endnote" ? "endnote" : "footnote";
    noteCount[kind] += 1;
    collect(
      note,
      undefined,
      `${kind === "endnote" ? "Endnote" : "Footnote"} ${noteCount[kind]}`,
    );
  }

  const changes = new WeakMap<ChangeEntry, ChangePlace>();
  const entries = records.map((record) => {
    const at = found.get(record.id);
    const blockIndex =
      at?.blockIndex ??
      (record.blockId ? positionOf.get(record.blockId) : undefined) ??
      (record.tableId ? positionOf.get(record.tableId) : undefined);
    const path =
      at?.path ??
      [record.tableId, record.blockId].filter((id): id is string =>
        Boolean(id && span.has(id)),
      );
    const place = path.length ? span.get(path[path.length - 1]) : undefined;
    const entry: ChangeEntry = {
      id: record.id,
      type: record.type ?? "text",
      author: record.agentName ?? record.authorType ?? "unknown",
      ...(at?.blockId || record.blockId
        ? { blockId: at?.blockId || record.blockId }
        : {}),
      ...(blockIndex !== undefined ? { blockIndex } : {}),
      ...(at?.label
        ? { section: at.label }
        : place && place.start < bodyEnd
          ? { section: sectionAt(place.start) }
          : {}),
      deleted: at?.deleted ?? "",
      inserted: at?.inserted ?? "",
    };
    if (place) changes.set(entry, { at: place.start, path });
    return entry;
  });
  placesByEntries.set(entries, { span, sections, changes });
  return entries;
}

/** Whether a heading is the one a section query names: "8", "8.", "8.1",
 * "Allocation of risk", or "8. Allocation". */
export function headingMatches(title: string, query: string): boolean {
  const heading = title.trim().toLowerCase();
  const q = query.trim().toLowerCase().replace(/\.$/, "");
  if (/^[\d.]+$/.test(q)) {
    return new RegExp(`^${q.replace(/\./g, "\\.")}(?:[.\\s)]|$)`).test(heading);
  }
  return heading.includes(q);
}

/** A --blocks query: block ids and `first..last` ranges ("…" and "..."
 * work too), e.g. "a1,b2..c3". */
export function parseBlockQuery(
  query: string,
): Array<string | { from: string; to: string }> {
  return query
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const [from, to, ...rest] = part
        .split(/\.\.\.?|…/)
        .map((id) => id.trim());
      if (to === undefined) return from;
      if (!from || !to || rest.length) {
        throw new ReviseError(
          `"${part}" is not a block range; write first..last.`,
          "invalid_selection",
        );
      }
      return { from, to };
    });
}

/**
 * The changes a filter keeps. A section is its heading's range, subsections
 * included; blocks are ids (a change inside a list item, table cell, or any
 * block nested in one counts) and inclusive ranges in document order.
 * Naming a section or block that doesn't exist is an error, not an empty
 * result.
 */
function applyChangeFilter(
  entries: ChangeEntry[],
  filter: ChangeFilter,
): ChangeEntry[] {
  const { section, exceptSection, blocks } = filter;
  if (!section && !exceptSection && !blocks?.length) return entries;
  const places = placesByEntries.get(entries);
  if (!places) return entries;
  const sectionRanges = (query: string) => {
    const q = query.trim().toLowerCase();
    // A title typed in full wins over headings that merely contain it.
    const exact = places.sections.filter(
      (s) =>
        s.title.toLowerCase() === q ||
        s.title.toLowerCase().replace(/^[\d.]+[.)]?\s+/, "") === q,
    );
    const candidates = exact.length
      ? exact
      : places.sections.filter((s) => headingMatches(s.title, query));
    // Subsections of a match are already inside it.
    const matched = candidates.filter(
      (s) =>
        !candidates.some(
          (outer) => outer !== s && outer.start < s.start && s.end <= outer.end,
        ),
    );
    if (matched.length > 1) {
      throw new ReviseError(
        `"${query}" matches ${matched.length} headings: ${matched
          .map((s) => `"${s.title}"`)
          .join(", ")}. Name one by its number or full title.`,
        "section_ambiguous",
      );
    }
    if (!matched.length) {
      throw new ReviseError(
        places.sections.length
          ? `No heading matches "${query}". Headings: ${places.sections
              .slice(0, 12)
              .map((s) => `"${s.title}"`)
              .join(", ")}${places.sections.length > 12 ? ", …" : ""}.`
          : `No heading matches "${query}": the document has no headings. Select by blocks or change ids instead.`,
        "section_not_found",
      );
    }
    return matched;
  };
  const spanOf = (id: string) => {
    const found = places.span.get(id);
    if (!found) {
      throw new ReviseError(`No block with id "${id}".`, "block_not_found");
    }
    return found;
  };
  const inside = sectionRanges;
  const keepSections = section ? inside(section) : undefined;
  const dropSections = exceptSection ? inside(exceptSection) : undefined;
  const ids = new Set<string>();
  const ranges: Array<{ start: number; end: number }> = [];
  for (const item of blocks ?? []) {
    if (typeof item === "string") {
      spanOf(item);
      ids.add(item);
      continue;
    }
    const from = spanOf(item.from);
    const to = spanOf(item.to);
    if (from.start > to.start) {
      throw new ReviseError(
        `Block "${item.from}" comes after "${item.to}"; write the range first..last.`,
        "invalid_selection",
      );
    }
    ranges.push({ start: from.start, end: to.end });
  }
  const within = (at: number, r: { start: number; end: number }) =>
    at >= r.start && at <= r.end;
  return entries.filter((entry) => {
    const place = places.changes.get(entry);
    if (!place) return false;
    if (keepSections && !keepSections.some((r) => within(place.at, r))) {
      return false;
    }
    if (dropSections?.some((r) => within(place.at, r))) return false;
    if (
      blocks?.length &&
      !place.path.some((id) => ids.has(id)) &&
      !ranges.some((r) => within(place.at, r))
    ) {
      return false;
    }
    return true;
  });
}

async function markdownOf(ydoc: Y.Doc): Promise<string> {
  const bytes = await exportDocument(ydoc, {
    format: "markdown",
    mode: "final",
  });
  return new TextDecoder().decode(bytes);
}

function lcsPairs<T>(a: T[], b: T[]): Array<[number, number]> {
  // Matching ends first: most real edits leave both untouched, and this
  // keeps the table below small even for book-length documents.
  let head = 0;
  while (head < a.length && head < b.length && a[head] === b[head]) head += 1;
  let tail = 0;
  while (
    tail < a.length - head &&
    tail < b.length - head &&
    a[a.length - 1 - tail] === b[b.length - 1 - tail]
  ) {
    tail += 1;
  }
  const pairs: Array<[number, number]> = [];
  for (let k = 0; k < head; k += 1) pairs.push([k, k]);
  const midA = a.slice(head, a.length - tail);
  const midB = b.slice(head, b.length - tail);
  // Past ~4M cells the table costs more than it is worth: report the
  // middle as replaced instead.
  if (midA.length * midB.length <= 4_000_000) {
    for (const [i, j] of lcsTable(midA, midB)) pairs.push([i + head, j + head]);
  }
  for (let k = tail; k > 0; k -= 1) {
    pairs.push([a.length - k, b.length - k]);
  }
  return pairs;
}

function lcsTable<T>(a: T[], b: T[]): Array<[number, number]> {
  const n = a.length;
  const m = b.length;
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

function wordDiff(before: string, after: string): string {
  const a = before.split(/(\s+)/);
  const b = after.split(/(\s+)/);
  if (a.length * b.length > 4_000_000) return `[-${before}-]{+${after}+}`;
  const pairs = lcsPairs(a, b);
  let i = 0;
  let j = 0;
  let output = "";
  const flush = (ai: number, bj: number) => {
    const removed = a.slice(i, ai).join("");
    const added = b.slice(j, bj).join("");
    if (removed) output += `[-${removed}-]`;
    if (added) output += `{+${added}+}`;
  };
  for (const [ai, bj] of pairs) {
    flush(ai, bj);
    output += a[ai];
    i = ai + 1;
    j = bj + 1;
  }
  flush(a.length, b.length);
  return output;
}

function textDiff(before: string, after: string): TextDiff {
  const split = (s: string) => s.split(/\n+/).filter((line) => line.trim());
  const a = split(before);
  const b = split(after);
  const pairs = lcsPairs(a, b);
  const hunks: string[] = [];
  let changed = 0;
  let i = 0;
  let j = 0;
  const flush = (ai: number, bj: number) => {
    const removed = a.slice(i, ai);
    const added = b.slice(j, bj);
    const paired = Math.min(removed.length, added.length);
    for (let k = 0; k < paired; k += 1) {
      hunks.push(`~ ${wordDiff(removed[k], added[k])}`);
      changed += 1;
    }
    for (const line of removed.slice(paired)) {
      hunks.push(`- ${line}`);
      changed += 1;
    }
    for (const line of added.slice(paired)) {
      hunks.push(`+ ${line}`);
      changed += 1;
    }
  };
  for (const [ai, bj] of pairs) {
    flush(ai, bj);
    i = ai + 1;
    j = bj + 1;
  }
  flush(a.length, b.length);
  return { text: hunks.join("\n"), changed };
}

async function documentSummary(
  ydoc: Y.Doc,
  session: ServerDocumentSession,
): Promise<DocumentSummary> {
  const doc = ydocToDocument(ydoc);
  const { resolvePageLayoutSettings, getPageDimensionsPx } =
    await engineModules.pageLayout();
  const layout = resolvePageLayoutSettings(doc.pageLayout);
  const { pageWidth, pageHeight } = getPageDimensionsPx(layout);
  const inches = (px: number) => Number((px / 96).toFixed(2));
  const margins = [
    layout.marginTopPx,
    layout.marginRightPx,
    layout.marginBottomPx,
    layout.marginLeftPx,
  ].map(inches);
  const measured = await session.tools.call("measure_document", {});
  const textWithFields = (node: unknown): string => {
    const n = node as {
      type?: string;
      text?: string;
      marks?: { field?: string };
      children?: unknown[];
    };
    if (n?.type === "text") {
      // Same tokens the header/footer tools accept: {PAGE}, {PAGES}.
      const field = n.marks?.field?.toLowerCase();
      if (field) {
        return field === "numpages" || field === "pages"
          ? "{PAGES}"
          : `{${field.toUpperCase()}}`;
      }
      return n.text ?? "";
    }
    return (n?.children ?? [])
      .map(textWithFields)
      .join(n?.type === "table-row" ? " | " : "");
  };
  const chrome = (doc.chrome ?? [])
    .filter((block) => block.type === "header-footer")
    .map((block) => {
      const b = block as unknown as { side: string; slot?: string };
      return {
        side: b.side,
        slot: b.slot ?? "default",
        text: textWithFields(block).replace(/\s+/g, " ").trim(),
      };
    })
    .filter((c) => c.text);
  let title: string | undefined;
  const headings: Record<string, number> = {};
  const lists: number[] = [];
  const tables: string[] = [];
  let images = 0;
  for (const block of doc.children) {
    if (block.type === "paragraph") {
      if (block.styleId === "Title" && title === undefined) {
        title = textWithFields(block);
      }
      if (block.heading) {
        headings[`H${block.heading}`] =
          (headings[`H${block.heading}`] ?? 0) + 1;
      }
    } else if (block.type === "list") {
      lists.push(block.children.length);
    } else if (block.type === "table") {
      tables.push(
        `${block.children.length}×${block.children[0]?.children.length ?? 0}`,
      );
    } else if (block.type === "image") {
      images += 1;
    }
  }
  const pending = session.listSuggestions().length;
  const data: DocumentSummary["data"] = {
    pages: measured.data.pageCount,
    words: measured.data.wordCount,
    pageSize: layout.pageSize,
    orientation: layout.orientation,
    pageInches: [inches(pageWidth), inches(pageHeight)],
    marginsInches: margins,
    title,
    headings,
    lists,
    tables,
    images,
    headersFooters: chrome,
    pendingTrackedChanges: pending,
  };
  const marginText = margins.every((m) => m === margins[0])
    ? `${margins[0]} in`
    : `${margins.join("/")} in (top/right/bottom/left)`;
  const parts = [
    title ? `title "${title}"` : "",
    Object.keys(headings).length
      ? `headings ${Object.entries(headings)
          .map(([level, n]) => `${n}×${level}`)
          .join(", ")}`
      : "",
    lists.length
      ? `${lists.length} list${lists.length > 1 ? "s" : ""} (${lists.join(", ")} items)`
      : "",
    tables.length
      ? `${tables.length} table${tables.length > 1 ? "s" : ""} (${tables.join(", ")})`
      : "",
    images ? `${images} image${images > 1 ? "s" : ""}` : "",
  ].filter(Boolean);
  const text = [
    `${data.pages ?? "?"} page${data.pages === 1 ? "" : "s"}, ${measured.data.wordCount} words; ${layout.pageSize} ${layout.orientation} (${inches(pageWidth)} × ${inches(pageHeight)} in), margins ${marginText}.`,
    ...(parts.length ? [`Contents: ${parts.join("; ")}.`] : []),
    ...chrome.map(
      (c) =>
        `${c.side[0].toUpperCase()}${c.side.slice(1)}${c.slot === "default" ? "" : ` (${c.slot})`}: ${c.text}`,
    ),
    ...(pending ? [`Pending tracked changes: ${pending}.`] : []),
  ].join("\n");
  return { text, data };
}
