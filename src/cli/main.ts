/**
 * `revise` — the word processor as a command line, for agents.
 *
 * A thin layer over `revise/server`: the same converters, semantic tools,
 * tracked changes, and graft export the editor uses, with document state kept
 * in `.revise/` between invocations (see store.ts). Output is plain text an
 * agent can read; `--json` switches every command to machine-readable JSON.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import * as Y from "yjs";
import {
  ReviseDocument,
  compare,
  convert,
  formatForPath,
  type DocumentSummary,
  type ReviseJsonObject,
  type ReviseToolResult,
} from "../lib";
import {
  createServerDocumentSession,
  decodeYDoc,
  encodeYDoc,
  fileToYDoc,
} from "../lib/internal";
import {
  READABLE,
  exportYDoc,
  parseBlockQuery,
  type ReadView,
} from "../lib/document";
import { relearn } from "../lib/docx";
import {
  CliError,
  amendLatestCheckpoint,
  canonicalPath,
  abandonPending,
  checkpointAt,
  commitState,
  createDocumentRecord,
  documentDir,
  hashFileBytes,
  readCheckpoint,
  readMeta,
  readPackage,
  readState,
  removeDocumentRecord,
  sessionExists,
  sessionNameFor,
  stageCommit,
  stageTruncate,
  finishPending,
  isCommitted,
  recoverPending,
  withFileLocks,
  writeFileAtomic,
  writeMeta,
  type CliDocumentMeta,
} from "./store";
import { SKILL_MARKDOWN, SKILL_REFERENCES } from "./skill";
import {
  INPUT_FORMATS,
  inputFormatFor,
  sniffFormat,
  UnreadableInput,
  utf8Of,
  type InputFormat,
} from "./sniff";
import { PLAIN, stripStyle, styleFor, type Style } from "./style";

declare const __REVISE_SDK_VERSION__: string | undefined;

// ── argument parsing ────────────────────────────────────────────────────────

const BOOLEAN_FLAGS = new Set([
  "details",
  "dry-run",
  "regex",
  "ignore-case",
  "exact",
  "json",
  "all",
  "final",
  "direct",
  "layout",
  "detailed",
  "help",
  "version",
  "force",
  "no-color",
]);

/** Every flag some command reads: anything else is a typo, and a typo that
 * silently does nothing is worse than an error. */
const VALUE_FLAGS = new Set([
  "author",
  "blocks",
  "content",
  "page-size",
  "results-page",
  "dir",
  "except-section",
  "file",
  "find",
  "from",
  "ids",
  "key",
  "limit",
  "margins",
  "match",
  "mode",
  "name",
  "offset",
  "orientation",
  "out",
  "reference",
  "page",
  "pages",
  "replace",
  "section",
  "since",
  "title",
  "to",
  "view",
  ...["header", "footer"].flatMap((side) => [
    side,
    `${side}-left`,
    `${side}-center`,
    `${side}-right`,
  ]),
]);

interface ParsedArgs {
  positional: string[];
  flags: Record<string, string | true>;
  /** The document a command operates on (its first argument). */
  file?: string;
}

function parseArgs(argv: string[]): ParsedArgs {
  const positional: string[] = [];
  const flags: Record<string, string | true> = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--") {
      positional.push(...argv.slice(i + 1));
      break;
    }
    if (arg.startsWith("--") && arg.length > 2) {
      const eq = arg.indexOf("=");
      const key = eq === -1 ? arg.slice(2) : arg.slice(2, eq);
      if (eq !== -1) flags[key] = arg.slice(eq + 1);
      else if (BOOLEAN_FLAGS.has(key) || i + 1 >= argv.length)
        flags[key] = true;
      else flags[key] = argv[++i];
    } else if (arg === "-h") {
      flags.help = true;
    } else {
      positional.push(arg);
    }
  }
  for (const key of Object.keys(flags)) {
    if (!BOOLEAN_FLAGS.has(key) && !VALUE_FLAGS.has(key)) {
      throw fail(`Unknown option --${key}.`, "usage", "revise help");
    }
  }
  return { positional, flags };
}

function flagString(args: ParsedArgs, name: string): string | undefined {
  const value = args.flags[name];
  return typeof value === "string" ? value : undefined;
}

/** --blocks (or its older name, --ids): block ids, comma-separated. */
function blockIdsFlag(args: ParsedArgs, command: string): string[] | undefined {
  const value = flagString(args, "blocks") ?? flagString(args, "ids");
  if (value === undefined) return undefined;
  const blocks = parseBlockQuery(value);
  if (blocks.some((b) => typeof b !== "string")) {
    throw fail(
      `revise ${command} takes block ids, not ranges.`,
      "usage",
      `revise ${command} --help`,
    );
  }
  return blocks as string[];
}

const VIEWS: readonly ReadView[] = ["tracked", "accepted", "original"];

/** --view: how pending tracked changes read. --final is the older name of
 * --view accepted. */
function viewFlag(args: ParsedArgs, fallback: ReadView): ReadView {
  const view = flagString(args, "view");
  if (view !== undefined && !VIEWS.includes(view as ReadView)) {
    throw fail("--view must be tracked, accepted, or original.", "usage");
  }
  if (args.flags.final === true) {
    if (view !== undefined && view !== "accepted") {
      throw fail(`--final means --view accepted, not --view ${view}.`, "usage");
    }
    return "accepted";
  }
  return (view as ReadView | undefined) ?? fallback;
}

function flagInt(args: ParsedArgs, name: string): number | undefined {
  const value = flagString(args, name);
  if (value === undefined) return undefined;
  return strictInt(value, `--${name}`);
}

/** A whole number, or a usage error: "3x" is not 3. */
function strictInt(value: string, name: string): number {
  if (!/^-?\d+$/.test(value.trim())) {
    throw new CliError(
      `${name} expects a whole number, got "${value}".`,
      "usage",
    );
  }
  return Number.parseInt(value, 10);
}

/** Commands that write a new file refuse to replace one unless --force. */
function assertCanOverwrite(args: ParsedArgs, target: string): void {
  if (fs.existsSync(target) && args.flags.force !== true) {
    throw wouldOverwrite(target);
  }
}

function wouldOverwrite(target: string): CliError {
  return fail(
    `${path.basename(target)} already exists. Add --force to replace it.`,
    "would_overwrite",
  );
}

// ── output ──────────────────────────────────────────────────────────────────
//
// Every command returns an Outcome; main() prints it once, as the JSON
// envelope (revise.io/developer/docs/revise/json) or as text. Errors are thrown as
// CliError and printed the same way. Warnings collect as the command runs.

/** Version of the --json contract. Bump only for a breaking change. */
export const JSON_SCHEMA = 1;

let jsonOutput = false;
/** The command being run (for usage errors). */
let currentCommand: string | null = null;

export interface CliIo {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
  /** Whether each stream is a terminal (color, JSON indentation). */
  stdoutIsTTY?: boolean;
  stderrIsTTY?: boolean;
  env?: Record<string, string | undefined>;
  /** A document read from or written to the standard streams (`-`). */
  readStdin?: () => Promise<Uint8Array>;
  stdinIsTTY?: boolean;
  writeStdout?: (bytes: Uint8Array) => void | Promise<void>;
}

let io: CliIo = defaultIo();

function defaultIo(): CliIo {
  return {
    stdout: (text) => process.stdout.write(text),
    stderr: (text) => process.stderr.write(text),
    stdoutIsTTY: Boolean(process.stdout.isTTY),
    stderrIsTTY: Boolean(process.stderr.isTTY),
    env: process.env,
    // A stream, not readFileSync(0): a pipe whose writer hasn't written yet
    // makes a synchronous read fail with EAGAIN.
    readStdin: async () => {
      const chunks: Uint8Array[] = [];
      for await (const chunk of process.stdin) chunks.push(chunk as Uint8Array);
      return new Uint8Array(Buffer.concat(chunks));
    },
    stdinIsTTY: Boolean(process.stdin.isTTY),
    writeStdout: (bytes) =>
      new Promise<void>((resolve, reject) => {
        process.stdout.write(bytes, (error) =>
          // A reader that stopped early (| head) wanted no more: not an error.
          error && (error as NodeJS.ErrnoException).code !== "EPIPE"
            ? reject(error)
            : resolve(),
        );
      }),
  };
}

/** A document piped in for `revise convert -`, with its format: named by
 * --from, else recognized from its bytes. */
async function readStdinDocument(args: ParsedArgs): Promise<{
  bytes: Uint8Array;
  format: InputFormat;
}> {
  const from = flagString(args, "from");
  const named = from === undefined ? undefined : inputFormatFor(from);
  const output = shellArg(args.positional[1] ?? "out.docx");
  if (from !== undefined && !named) {
    throw fail(
      `--from takes ${INPUT_FORMATS.join(", ")}.`,
      "unsupported_format",
      // Without --from, revise recognizes the format itself.
      `revise convert - ${output}`,
    );
  }
  if (io.stdinIsTTY || !io.readStdin) {
    throw fail(
      "Nothing is piped in. - reads the document from stdin.",
      "usage",
      `cat notes.md | revise convert - ${output}`,
    );
  }
  const bytes = await io.readStdin();
  if (named) {
    // Named text gets the same encoding handling recognition gives it.
    const text = named !== "docx" && named !== "odt";
    return { bytes: text ? utf8Of(bytes) : bytes, format: named };
  }
  try {
    return sniffFormat(bytes);
  } catch (error) {
    if (!(error instanceof UnreadableInput)) throw error;
    throw fail(
      error.message,
      "unsupported_format",
      bytes.length === 0
        ? `cat notes.md | revise convert - ${output}`
        : undefined,
    );
  }
}

function inputFormatOfPath(file: string): InputFormat | null {
  const match = /\.([^./\\]+)$/.exec(file);
  return match ? (inputFormatFor(match[1]) ?? null) : null;
}

/** Color for stdout and stderr, decided per stream in main(). */
let ui: Style = PLAIN;
let uiErr: Style = PLAIN;

export interface Problem {
  code: string;
  message: string;
  details: Record<string, unknown> | null;
}

interface CliErrorInfo extends Problem {
  hint: string | null;
}

/** Who produced the output: the package version and the commit it was
 * built from. */
declare const __REVISE_SDK_BUILD__: string | undefined;
function producer(): { version: string; build: string | null } {
  return {
    version:
      typeof __REVISE_SDK_VERSION__ === "string"
        ? __REVISE_SDK_VERSION__
        : "dev",
    build:
      typeof __REVISE_SDK_BUILD__ === "string" ? __REVISE_SDK_BUILD__ : null,
  };
}

/** What a command did. With `error` it failed after all: exit 2 when it
 * saved or wrote part of what was asked, 1 when it changed nothing. */
interface Outcome {
  data: Record<string, unknown>;
  text: () => string;
  error?: CliErrorInfo;
  exit?: 1 | 2;
}

let warnings: Problem[] = [];

/** Something worth knowing that did not stop the command. */
function warn(
  code: string,
  message: string,
  details: Record<string, unknown> | null = null,
): void {
  if (warnings.some((w) => w.code === code && w.message === message)) return;
  warnings.push({ code, message, details });
  if (!jsonOutput) {
    io.stderr(`${uiErr.yellow(uiErr.bold(`warning[${code}]`))}: ${message}\n`);
  }
}

/** An expected failure, with a command to run next and structured details. */
/** The per-target errors a tool's failure names ("Resolve every listed
 * target error"), listed so the message is complete without --json. */
function listedErrors(details: unknown): string {
  const errors = (details as { errors?: unknown } | undefined)?.errors;
  if (!Array.isArray(errors) || !errors.length) return "";
  const lines = errors.slice(0, 20).map((error) => {
    const e = error as Record<string, unknown>;
    const where = [
      typeof e.row_id === "string" && e.row_id ? `row ${e.row_id}` : null,
      typeof e.column === "number" && e.column >= 0
        ? `column ${e.column}`
        : null,
      typeof e.id === "string" && e.id ? `block ${e.id}` : null,
    ]
      .filter(Boolean)
      .join(", ");
    return `\n  - ${where ? `${where}: ` : ""}${String(e.reason ?? e.message ?? JSON.stringify(e))}`;
  });
  const more =
    errors.length > 20 ? `\n  - ...and ${errors.length - 20} more` : "";
  return `${lines.join("")}${more}\n`;
}

function fail(
  message: string,
  code: string,
  hint?: string | null,
  details?: Record<string, unknown>,
  receipt?: { data: Record<string, unknown>; exit: 2 | 70 },
): CliError {
  const error = new CliError(message, code);
  extras.set(error, {
    hint: hint ?? null,
    details: details ?? null,
    receipt: receipt ?? null,
  });
  return error;
}

const extras = new WeakMap<
  object,
  {
    hint: string | null;
    details: Record<string, unknown> | null;
    /** What was done before the failure (a partial result). */
    receipt: { data: Record<string, unknown>; exit: 2 | 70 } | null;
  }
>();

/** Library codes, in the CLI's vocabulary. */
const LIBRARY_CODES: Record<string, string> = { not_found: "file_not_found" };

function problemOf(error: CliError): CliErrorInfo {
  const extra = extras.get(error);
  return {
    code: LIBRARY_CODES[error.code] ?? error.code,
    message: error.message,
    hint: extra?.hint ?? null,
    details: extra?.details ?? null,
  };
}

function out(text: string): void {
  io.stdout(text.endsWith("\n") ? text : `${text}\n`);
}

function printEnvelope(
  command: string | null,
  data: Record<string, unknown> | null,
  error: CliErrorInfo | null,
): void {
  const envelope = {
    schema: JSON_SCHEMA,
    ok: error === null,
    command,
    meta: producer(),
    data,
    warnings,
    error,
  };
  out(JSON.stringify(envelope, null, io.stdoutIsTTY ? 2 : undefined));
}

function printError(error: CliErrorInfo): void {
  io.stderr(
    `${uiErr.red(uiErr.bold(`error[${error.code}]`))}: ${error.message}\n` +
      (error.hint ? `  ${uiErr.dim("hint:")} ${uiErr.cyan(error.hint)}\n` : ""),
  );
}

/** A path as the person typed it, quoted for a shell when it needs it. */
function shellArg(value: string): string {
  return /^[\w@%+=:,./-]+$/.test(value)
    ? value
    : `'${value.replace(/'/g, `'\\''`)}'`;
}

/** "1 block", "3 blocks". */
function plural(n: number, one: string, many = `${one}s`): string {
  return `${n} ${n === 1 ? one : many}`;
}

/** A format as the --json contract names it: its file extension. */
function formatName(format: string): string {
  return (
    ({ markdown: "md", htm: "html" } as Record<string, string | undefined>)[
      format
    ] ?? format
  );
}

/** The tracked-change mode in CLI words. */
function modeName(mode: CliDocumentMeta["mode"]): "suggest" | "edit" {
  return mode === "suggesting" ? "suggest" : "edit";
}

// ── document loading ────────────────────────────────────────────────────────
//
// Every document command names its file. The cache in .revise/ (keyed by the
// file's path) holds the model, the original package, and a checkpoint per
// edit; the file on disk is the truth. Edits write the file in place.

interface LoadedDocument {
  meta: CliDocumentMeta;
  doc: ReviseDocument;
  /** The file as the person named it, ready to paste into a command. */
  shown: string;
}

const DEFAULT_AUTHOR = "Agent";
/** Formats revise can write back in place. */
const WRITABLE = /\.(docx|md|markdown|html?|txt)$/i;
/** Formats with no tracked changes: edits there apply directly. */
const PLAIN_FORMATS = /\.(md|markdown|html?|txt)$/i;

function authorFor(args: ParsedArgs, meta?: CliDocumentMeta): string {
  return (
    flagString(args, "author") ??
    process.env.REVISE_AUTHOR ??
    meta?.author ??
    DEFAULT_AUTHOR
  );
}

function requireFile(args: ParsedArgs): string {
  const file = args.file;
  if (!file) {
    const help = HELP_BY_NAME.get(currentCommand ?? "");
    throw help
      ? usageError(help.name)
      : fail(
          "Name the file: revise <command> <file> ...",
          "usage",
          "revise help",
        );
  }
  if (!fs.existsSync(file))
    throw new CliError(`No such file: ${file}`, "file_not_found");
  if (!READABLE.test(file)) {
    throw new CliError(
      `Cannot read ${path.basename(file)}. Supported: .docx, .md, .html, .odt, .rtf, .txt.`,
      "unsupported_format",
    );
  }
  return path.resolve(file);
}

async function loadDocument(args: ParsedArgs): Promise<LoadedDocument> {
  const file = canonicalPath(requireFile(args));
  const name = sessionNameFor(file);
  const bytes = new Uint8Array(fs.readFileSync(file));
  const hash = hashFileBytes(bytes);
  const isDocx = /\.docx$/i.test(file);
  let meta: CliDocumentMeta;
  let ydoc: Y.Doc;
  if (!sessionExists(name)) {
    // First touch: import the file into a new cache.
    ydoc = await importFile(bytes, file);
    meta = createDocumentRecord(
      {
        name,
        file,
        fileHash: hash,
        documentId: `cli-${name}`,
        source: file,
        sourceFormat: path.extname(file).slice(1).toLowerCase(),
        createdAt: new Date().toISOString(),
        mode: PLAIN_FORMATS.test(file) ? "editing" : "suggesting",
        ...((flagString(args, "author") ?? process.env.REVISE_AUTHOR)
          ? { author: flagString(args, "author") ?? process.env.REVISE_AUTHOR }
          : {}),
      },
      encodeYDoc(ydoc),
      isDocx ? bytes : null,
    );
  } else {
    meta = readMeta(name);
    // Finish or discard a save an interrupted command left half done.
    const recovered = recoverPending(meta, hash);
    if (recovered === "finished") {
      warn(
        "save_recovered",
        `Finished saving ${path.basename(file)}: an earlier command was interrupted after writing it.`,
      );
    }
    ydoc = decodeYDoc(readState(meta));
    if (meta.fileHash !== hash) {
      // The file changed outside revise: re-learn it. Block ids carry over
      // only where identity is provable (Word's paragraph ids, identical
      // content); reads of changed or new blocks are no longer valid.
      const fresh = await relearn(ydoc, bytes, file);
      ydoc = fresh.ydoc;
      const stale = new Set([...fresh.changedIds, ...fresh.removedIds]);
      meta.readBlockIds = (meta.readBlockIds ?? []).filter(
        (id) => !stale.has(id),
      );
      // Saved searches ran against the old content: their ids are void.
      meta.searches = {};
      commitState(
        meta,
        encodeYDoc(ydoc),
        "file changed",
        "Re-read the file after it changed outside revise",
        { fileHash: hash, package: isDocx ? bytes : null },
      );
      warn(
        "file_changed_outside",
        `${path.basename(file)} changed outside revise; re-read it. ` +
          [
            `${fresh.idsKept} block ids kept`,
            fresh.changedIds.length
              ? `${fresh.changedIds.length} kept blocks have new content`
              : "",
            fresh.idsNew ? `${fresh.idsNew} blocks are new` : "",
            fresh.removedIds.length
              ? `${fresh.removedIds.length} ids no longer exist`
              : "",
          ]
            .filter(Boolean)
            .join(", ") +
          (fresh.changedIds.length || fresh.idsNew || fresh.removedIds.length
            ? ". Read again before editing changed or new blocks."
            : "."),
        {
          keptBlockCount: fresh.idsKept,
          changedBlockIds: fresh.changedIds,
          newBlockCount: fresh.idsNew,
          removedBlockIds: fresh.removedIds,
        },
      );
    } else if (meta.packageDirty) {
      // An xml patch rewrote the file: re-learn the model from it before
      // anything reads the model. Folded into the patch's checkpoint.
      const synced = await relearn(ydoc, bytes, file);
      ydoc = synced.ydoc;
      amendLatestCheckpoint(meta, encodeYDoc(ydoc));
    }
  }
  const doc = await ReviseDocument.fromYDoc(ydoc, {
    documentId: meta.documentId,
    mode: meta.mode,
    author: authorFor(args, meta),
    package: readPackage(meta),
    // Exports without a saved title use the file's name, not the cache's.
    title: path.parse(file).name,
    // The CLI fetches images for the user, on the user's machine.
    images: { allowPrivateNetwork: true },
  });
  return { meta, doc, shown: shellArg(args.file ?? file) };
}

async function importFile(bytes: Uint8Array, file: string): Promise<Y.Doc> {
  try {
    return await fileToYDoc(bytes, file);
  } catch (error) {
    throw new CliError(
      `${path.basename(file)} does not open: ${(error as Error).message}`,
      "import_failed",
    );
  }
}

/** The document as its file's format would hold it. */
async function exportBytes(
  meta: CliDocumentMeta,
  ydoc: Y.Doc,
  pkg: Uint8Array | null,
): Promise<Uint8Array> {
  return (
    await exportYDoc(ydoc, pkg, {
      format: formatForPath(meta.file) ?? "txt",
      title: path.basename(meta.file).replace(/\.[^.]+$/, ""),
    })
  ).bytes;
}

/**
 * The file as revise last knew it, or a refusal: when another program saved
 * (or deleted) the file since revise read it, nothing may overwrite it.
 * Returns the current bytes (null when the file is new).
 */
function assertUnchanged(
  meta: CliDocumentMeta,
  expectHash: string | null,
): Uint8Array | null {
  if (!fs.existsSync(meta.file)) {
    if (expectHash === null) return null;
    throw new CliError(
      `${path.basename(meta.file)} was deleted while revise was working on it. Nothing was written.`,
      "file_changed",
    );
  }
  const current = new Uint8Array(fs.readFileSync(meta.file));
  if (expectHash !== null && hashFileBytes(current) !== expectHash) {
    throw new CliError(
      `${path.basename(meta.file)} changed while revise was working on it (another program saved it). Nothing was written; run the command again.`,
      "file_changed",
    );
  }
  return current;
}

function writeOrFail(file: string, bytes: Uint8Array): void {
  try {
    writeFileAtomic(file, bytes);
  } catch (error) {
    throw new CliError(
      `Could not write ${path.basename(file)}: ${(error as Error).message}. Nothing was changed.`,
      "write_failed",
    );
  }
}

/**
 * Write the document to its file atomically and return the bytes' hash,
 * refusing if the file changed since revise read it. For files with no
 * history yet; saves that add to history go through saveJournaled.
 */
export async function writeDocumentFile(
  meta: CliDocumentMeta,
  ydoc: Y.Doc,
  pkg: Uint8Array | null,
  options: { expectHash?: string | null; bytes?: Uint8Array } = {},
): Promise<string> {
  const bytes = options.bytes ?? (await exportBytes(meta, ydoc, pkg));
  assertUnchanged(
    meta,
    options.expectHash === undefined ? meta.fileHash : options.expectHash,
  );
  writeOrFail(meta.file, bytes);
  return hashFileBytes(bytes);
}

/**
 * Save the file and its history as one change: stage the history entry,
 * write the file, then commit the entry. If the commit fails, the file is
 * put back as it was; if the process dies midway, the next command finishes
 * or discards the staged entry by what the file holds (recoverPending).
 * Returns the checkpoint the file now corresponds to.
 */
async function saveJournaled(
  meta: CliDocumentMeta,
  bytes: Uint8Array,
  stage: (fileHash: string) => string,
): Promise<number> {
  const previous = assertUnchanged(meta, meta.fileHash);
  const tx = stage(hashFileBytes(bytes));
  try {
    writeOrFail(meta.file, bytes);
  } catch (error) {
    abandonPending(meta);
    throw error;
  }
  try {
    return finishPending(meta);
  } catch (error) {
    // Committed after all (the log names this save): it succeeded.
    if (isCommitted(meta.name, tx)) {
      Object.assign(meta, readMeta(meta.name));
      return meta.checkpoints.at(-1)?.n ?? 0;
    }
    try {
      if (previous) writeFileAtomic(meta.file, previous);
      abandonPending(meta);
    } catch (rollback) {
      // The next command recovers from pending.json, but this one can't say
      // what the file holds now.
      throw fail(
        `Could not record the change in revise's history (${(error as Error).message}), and could not put ${path.basename(meta.file)} back (${(rollback as Error).message}). The next revise command on it will finish or undo the save.`,
        "save_uncertain",
        `revise status ${shellArg(meta.file)}`,
        { file: meta.file, saved: null, checkpoint: null },
      );
    }
    throw new CliError(
      `Could not record the change in revise's history (${(error as Error).message}); ${path.basename(meta.file)} was put back as it was.`,
      "commit_failed",
    );
  }
}

function assertWritable(meta: CliDocumentMeta): void {
  if (!WRITABLE.test(meta.file)) {
    throw new CliError(
      `revise can't write ${path.extname(meta.file)} files, so it won't edit ${path.basename(meta.file)} in place. Convert it first: revise convert ${path.basename(meta.file)} ${path.basename(meta.file).replace(/\.[^.]+$/, "")}.docx`,
      "read_only_format",
    );
  }
}

/**
 * Run `work`; if it changed the document, save the file and its history
 * together (saveJournaled): a failed command leaves both as they were.
 */
async function mutate<T>(
  loaded: LoadedDocument,
  command: string,
  work: () => Promise<T>,
  summarize: (result: T) => string,
  options: { saveIf?: (result: T) => boolean } = {},
): Promise<{ result: T; checkpoint: number | null; discarded?: boolean }> {
  const before = Buffer.from(loaded.doc.version());
  const result = await work();
  const changed = !before.equals(Buffer.from(loaded.doc.version()));
  if (!changed) return { result, checkpoint: null };
  // A call that failed saves nothing, even if some engine path changed the
  // document before failing: the file and history stay as they were.
  if (options.saveIf && !options.saveIf(result)) {
    return { result, checkpoint: null, discarded: true };
  }
  // Only a real change needs a writable file; reads through tools pass.
  assertWritable(loaded.meta);
  const bytes = await exportBytes(
    loaded.meta,
    loaded.doc.ydoc,
    loaded.doc.package,
  );
  const state = encodeYDoc(loaded.doc.ydoc);
  const checkpoint = await saveJournaled(loaded.meta, bytes, (fileHash) =>
    stageCommit(
      loaded.meta,
      state,
      { command, summary: summarize(result).slice(0, 300) },
      fileHash,
    ),
  );
  return { result, checkpoint };
}

/** The checkpoint the file is at now. */
function currentCheckpoint(meta: CliDocumentMeta): number {
  return meta.checkpoints.at(-1)?.n ?? 0;
}

/** "Saved contract.docx (checkpoint 3). Undo: revise undo contract.docx" */
function savedLine(loaded: LoadedDocument, checkpoint: number): string {
  return `${ui.green("Saved")} ${loaded.shown} (checkpoint ${checkpoint}). ${ui.dim("Undo:")} ${ui.cyan(`revise undo ${loaded.shown}`)}`;
}

/** "Next: revise changes contract.docx" */
function nextLine(label: string, ...commands: string[]): string {
  const pad = " ".repeat(label.length + 2);
  return commands
    .map((c, i) => `${i ? pad : ui.dim(`${label}:`) + " "}${ui.cyan(c)}`)
    .join("\n");
}

function readInputJson(raw: string | undefined): ReviseJsonObject {
  if (raw === undefined) return {};
  let text = raw;
  if (raw === "-") text = fs.readFileSync(0, "utf8");
  else if (raw.startsWith("@")) text = fs.readFileSync(raw.slice(1), "utf8");
  try {
    const parsed = JSON.parse(text) as unknown;
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      throw new Error("not an object");
    }
    return parsed as ReviseJsonObject;
  } catch (error) {
    throw new CliError(
      `Tool input must be a JSON object (${(error as Error).message}).`,
      "invalid_json",
    );
  }
}

/** Give a search's result set a workspace-unique id and remember how to
 * rebuild it. In-process ids restart at search_result_1 in every process. */
function persistSearch<V extends { message: string | null; data: unknown }>(
  meta: CliDocumentMeta,
  input: ReviseJsonObject,
  value: V,
): V {
  const data = value.data as { search_result_id?: string } | null;
  const processId = data?.search_result_id;
  if (!processId) return value;
  meta.searchCounter = (meta.searchCounter ?? 0) + 1;
  const id = `search_result_${meta.searchCounter}`;
  const { page: _page, ...query } = input;
  // Bind the id to exactly what the agent was shown: a replay that finds
  // anything else is refused instead of silently widening the target.
  const found = data as {
    totalMatches?: number;
    matches?: Array<{ blockId?: string }>;
  };
  const searches = {
    ...(meta.searches ?? {}),
    [id]: {
      query: { ...query, page: 1 },
      total: found.totalMatches ?? null,
      targets: (found.matches ?? []).map((m) => m.blockId ?? ""),
    },
  };
  // Keep the most recent 20.
  meta.searches = Object.fromEntries(Object.entries(searches).slice(-20));
  writeMeta(meta);
  const rename = (text: string) => text.split(processId).join(id);
  return {
    ...value,
    message: value.message ? rename(value.message) : value.message,
    data: JSON.parse(rename(JSON.stringify(value.data))),
  };
}

const READ_TOOLS = new Set([
  "read_blocks_from_index",
  "read_specific_blocks",
  "search_document",
  "find_highlights",
  "read_table",
]);

/** Every element id in read/search output, nested cells included. */
function deliveredIds(html: string): string[] {
  return Array.from(
    html.matchAll(/<[a-z][a-z0-9]*\b[^>]*?\sid="([^"]+)"/g),
    (m) => m[1],
  );
}

function rememberReadBlocks(meta: CliDocumentMeta, html: string): void {
  const ids = new Set(meta.readBlockIds ?? []);
  for (const id of deliveredIds(html)) ids.add(id);
  meta.readBlockIds = Array.from(ids).slice(-20_000);
  writeMeta(meta);
}

function stringsIn(value: unknown, into: Set<string> = new Set()): Set<string> {
  if (typeof value === "string") into.add(value);
  else if (Array.isArray(value)) value.forEach((v) => stringsIn(v, into));
  else if (value && typeof value === "object") {
    Object.values(value).forEach((v) => stringsIn(v, into));
  }
  return into;
}

/**
 * The tools refuse some operations (comments, most notably) on blocks the
 * caller has not read in the current session. A CLI session spans processes,
 * so re-deliver — silently, in this process — the targeted blocks the agent
 * was already shown by an earlier read or search. Blocks it never read stay
 * unread, so the guard keeps its meaning.
 */
async function restoreReadContext(
  loaded: LoadedDocument,
  input: ReviseJsonObject,
): Promise<void> {
  const read = new Set(loaded.meta.readBlockIds ?? []);
  if (!read.size) return;
  const targets = Array.from(stringsIn(input)).filter((s) => read.has(s));
  if (!targets.length) return;
  await loaded.doc.session.tools.executeDynamic("read_specific_blocks", {
    block_ids: targets,
    context_notes: "N/A",
  });
  // Cell edits check a table read, not a block read: re-deliver the rows
  // of that table the agent was shown, and only those.
  const tableId = input.table_id;
  const cells = Array.isArray(input.cells) ? input.cells : [];
  const rowIds = [
    ...new Set(
      cells
        .map((cell) => (cell as { row_id?: unknown })?.row_id)
        .filter((id): id is string => typeof id === "string" && read.has(id)),
    ),
  ];
  if (typeof tableId === "string" && read.has(tableId) && rowIds.length) {
    await loaded.doc.session.tools.executeDynamic("read_table", {
      table_id: tableId,
      mode: "rows",
      row_ids: rowIds.slice(0, 50),
      context_notes: "N/A",
    });
  }
}

/** Tool messages speak in tool calls ("read_blocks_from_index(index=65)");
 * say the same thing as the CLI command an agent would actually run. */
function cliVocabulary(message: string, file: string): string {
  return message
    .replace(/read_blocks_from_index\(index=(\d+)\)/g, `revise read ${file} $1`)
    .replace(/read_specific_blocks\b/g, `revise read ${file} --blocks`)
    .replace(/search_document\b/g, `revise search ${file}`);
}

/** A tool call's result, in the shape every tool command reports. */
function toolData(
  loaded: LoadedDocument,
  tool: string,
  fields: {
    saved: boolean;
    checkpoint: number;
    message: string | null;
    changeIds: string[];
    html: string | null;
    result: unknown;
  },
): Record<string, unknown> {
  return { file: loaded.meta.file, tool, ...fields };
}

/** Run one tool call. */
async function runTool(
  loaded: LoadedDocument,
  name: string,
  input: ReviseJsonObject,
  options: { direct?: boolean; view?: ReadView } = {},
): Promise<Outcome> {
  const definition = loaded.doc.session.tools
    .getDefinitions()
    .find((tool) => tool.name === name);
  if (!definition) {
    throw fail(`Unknown tool "${name}".`, "unknown_tool", "revise tools");
  }
  const required = (definition.inputSchema.required as string[]) ?? [];
  if (required.includes("context_notes") && input.context_notes === undefined) {
    input = { ...input, context_notes: "N/A" };
  }
  if (!READ_TOOLS.has(name)) await restoreReadContext(loaded, input);
  // search_result_id names a result set held in process memory. Each CLI call
  // is a new process, so replay the saved search first (against the current
  // document) and point the call at the fresh in-process id.
  if (typeof input.search_result_id === "string") {
    const searchAgain = `revise search ${loaded.shown} "<text>"`;
    const saved = loaded.meta.searches?.[input.search_result_id];
    if (!saved) {
      throw fail(
        `Unknown search_result_id "${input.search_result_id}". Run the search again and use the id it prints.`,
        "unknown_search_result",
        searchAgain,
      );
    }
    const bound = saved as {
      query?: ReviseJsonObject;
      total?: number | null;
      targets?: string[];
    };
    if (!bound.query || !bound.targets) {
      throw fail(
        `search_result_id "${input.search_result_id}" is from an older revise. Run the search again.`,
        "stale_search_result",
        searchAgain,
      );
    }
    const replay = await loaded.doc.session.tools.executeDynamic(
      "search_document",
      bound.query,
    );
    const found =
      replay.ok && replay.value.data && typeof replay.value.data === "object"
        ? (replay.value.data as {
            search_result_id?: string;
            totalMatches?: number;
            matches?: Array<{ blockId?: string }>;
          })
        : undefined;
    const freshId = found?.search_result_id;
    const sameTargets =
      found &&
      (bound.total === null || found.totalMatches === bound.total) &&
      JSON.stringify((found.matches ?? []).map((m) => m.blockId ?? "")) ===
        JSON.stringify(bound.targets);
    if (!freshId || !sameTargets) {
      throw fail(
        `The document changed since search ${input.search_result_id}, so it no longer names the blocks you were shown. Run the search again.`,
        "stale_search_result",
        searchAgain,
      );
    }
    input = { ...input, search_result_id: freshId };
  }
  // revise.io editor display settings: a file has nowhere to keep them, so
  // the call would report success and change nothing.
  const displayOnly =
    name === "set_page_layout"
      ? ["pageless", "blockSpacing", "indentStep"].filter((key) => key in input)
      : [];
  if (displayOnly.length) {
    throw fail(
      `${displayOnly.join(", ")} ${displayOnly.length > 1 ? "are" : "is"} a revise.io editor display setting that a document file can't store. Set page size, margins, orientation, columns, or page numbers instead.`,
      "usage",
      `revise ${name} --help`,
    );
  }
  const { result, checkpoint } = await mutate(
    loaded,
    name,
    () => loaded.doc.tools.run(name, input, { direct: options.direct }),
    (r: ReviseToolResult<string>) =>
      "error" in r ? r.error.message : (r.value.message ?? name),
    { saveIf: (r) => !("error" in r) },
  );
  const now = currentCheckpoint(loaded.meta);
  if ("error" in result) {
    const failure = result.error;
    // Asking for what is already there is not a failure.
    if (/^No changes detected/i.test(failure.message)) {
      return {
        data: toolData(loaded, name, {
          saved: false,
          checkpoint: now,
          message: "Already matches: nothing to change.",
          changeIds: [],
          html: null,
          result: null,
        }),
        text: () => "Already matches: nothing to change. File not rewritten.",
      };
    }
    return {
      data: toolData(loaded, name, {
        saved: false,
        checkpoint: now,
        message: null,
        changeIds: [],
        html: null,
        result: null,
      }),
      text: () => "",
      error: {
        code: "tool_error",
        message: `${cliVocabulary(failure.message, loaded.shown)}${listedErrors(failure.details)} Nothing was saved.`,
        // An input the tool refused: its fields are one command away.
        hint: /invalid|require[sd]?|must be|unknown (field|property|key)|expected|missing/i.test(
          failure.message,
        )
          ? `revise ${name} --help`
          : null,
        details: {
          tool: name,
          toolCode: failure.code,
          ...(failure.details !== undefined
            ? { toolDetails: failure.details }
            : {}),
        },
      },
      exit: 1,
    };
  }
  let value = result.value;
  if (name === "search_document")
    value = persistSearch(loaded.meta, input, value);
  if (value.context?.html) rememberReadBlocks(loaded.meta, value.context.html);
  const changeIds = value.suggestionIds ?? [];
  // The tools read the accepted view; tracked and original are rendered
  // from the same blocks.
  const html =
    value.context?.html && options.view && options.view !== "accepted"
      ? await loaded.doc.viewHtml(value.context.html, options.view)
      : (value.context?.html ?? null);
  return {
    data: {
      ...toolData(loaded, name, {
        saved: checkpoint !== null,
        checkpoint: now,
        message: value.message,
        changeIds,
        html,
        result: value.data ?? null,
      }),
      ...(options.view ? { view: options.view } : {}),
    },
    text: () => {
      const lines: string[] = [];
      if (value.message)
        lines.push(cliVocabulary(value.message, loaded.shown).trimEnd());
      if (changeIds.length) {
        lines.push(
          `${plural(changeIds.length, "tracked change")}: ${changeIds.map((id) => ui.cyan(id)).join(", ")}`,
        );
      }
      if (checkpoint !== null) lines.push(savedLine(loaded, checkpoint));
      if (html) lines.push("", html);
      else if (
        value.data &&
        typeof value.data === "object" &&
        !("effect" in (value.data as object)) &&
        !READ_TOOLS.has(name)
      ) {
        lines.push("", JSON.stringify(value.data, null, 2));
      }
      return lines.join("\n");
    },
  };
}

// ── commands ────────────────────────────────────────────────────────────────

/** The document summary in the --json contract's shape. */
function summaryData(summary: DocumentSummary["data"]) {
  const [top, right, bottom, left] = summary.marginsInches;
  return {
    pageCount: summary.pages ?? null,
    wordCount: summary.words,
    pageSize: summary.pageSize,
    orientation: summary.orientation,
    pageInches: { width: summary.pageInches[0], height: summary.pageInches[1] },
    marginsInches: { top, right, bottom, left },
    title: summary.title ?? null,
    headingCounts: summary.headings,
    lists: summary.lists.map((itemCount) => ({ itemCount })),
    tables: summary.tables.map((size) => {
      const [rowCount, columnCount] = size.split("×").map(Number);
      return { rowCount, columnCount };
    }),
    imageCount: summary.images,
    headersFooters: summary.headersFooters,
    trackedChangeCount: summary.pendingTrackedChanges,
  };
}

/** The summary as two or three short lines for people. */
function summaryText(summary: ReturnType<typeof summaryData>): string {
  const m = summary.marginsInches;
  const same = m.top === m.right && m.top === m.bottom && m.top === m.left;
  const size =
    summary.pageSize.length <= 2
      ? summary.pageSize.toUpperCase()
      : summary.pageSize[0].toUpperCase() + summary.pageSize.slice(1);
  const lines = [
    [
      summary.pageCount !== null ? plural(summary.pageCount, "page") : null,
      plural(summary.wordCount, "word"),
      `${size} ${summary.orientation} (${summary.pageInches.width} × ${summary.pageInches.height} in)`,
      same
        ? `margins ${m.top} in`
        : `margins ${m.top}/${m.right}/${m.bottom}/${m.left} in (top/right/bottom/left)`,
    ]
      .filter(Boolean)
      .join(", "),
  ];
  const contents = [
    summary.title !== null ? `title "${summary.title}"` : "",
    Object.keys(summary.headingCounts).length
      ? `headings ${Object.entries(summary.headingCounts)
          .map(([level, n]) => `${n}×${level}`)
          .join(", ")}`
      : "",
    summary.lists.length
      ? `${plural(summary.lists.length, "list")} (${summary.lists.map((l) => l.itemCount).join(", ")} items)`
      : "",
    summary.tables.length
      ? `${plural(summary.tables.length, "table")} (${summary.tables.map((t) => `${t.rowCount}×${t.columnCount}`).join(", ")})`
      : "",
    summary.imageCount ? plural(summary.imageCount, "image") : "",
  ].filter(Boolean);
  if (contents.length) lines.push(`Contents: ${contents.join("; ")}`);
  for (const c of summary.headersFooters) {
    lines.push(
      `${c.side[0].toUpperCase()}${c.side.slice(1)}${c.slot === "default" ? "" : ` (${c.slot})`}: ${c.text}`,
    );
  }
  if (summary.trackedChangeCount) {
    lines.push(plural(summary.trackedChangeCount, "tracked change"));
  }
  return lines.join("\n");
}

function sizeText(bytes: number): string {
  return bytes < 1024
    ? plural(bytes, "byte")
    : `${(bytes / 1024).toFixed(1)} KiB`;
}

/**
 * Create a document in one step: from the same HTML the tools speak
 * (data-style-id="Title", tables, lists — exactly what `read` shows) or
 * Markdown, with page setup and header/footer applied, written to a file.
 */
async function cmdNew(args: ParsedArgs): Promise<Outcome> {
  const target = args.positional[0];
  if (!target) throw usageError("new");
  if (!WRITABLE.test(target)) {
    throw fail(
      `revise can't write ${path.extname(target) || "files without an extension"}. Create a .docx, .md, .html, or .txt file.`,
      "unsupported_format",
    );
  }
  const requestedMode = flagString(args, "mode");
  if (requestedMode && !["suggest", "edit"].includes(requestedMode)) {
    throw fail("--mode must be suggest or edit.", "usage", "revise new --help");
  }
  // --from is the 3.0 name: other commands use it for a format.
  const from = flagString(args, "content") ?? flagString(args, "from");
  if (from !== undefined && !fs.existsSync(from)) {
    throw fail(`No such file: ${from}`, "file_not_found");
  }
  const title = flagString(args, "title");
  const created =
    from && !/\.html?$/i.test(from)
      ? await ReviseDocument.open(from, { mode: "editing" })
      : await ReviseDocument.create(
          from ? { html: fs.readFileSync(from, "utf8"), title } : { title },
        );
  // Everything the page setup needs is checked before anything is written.
  const steps = newDocumentSteps(args, created);
  const format = formatName(formatForPath(target) ?? "txt");
  const resolved = canonicalPath(target);
  assertCanOverwrite(args, resolved);
  writeOutput(
    args,
    resolved,
    (
      await created.export({
        format: formatForPath(target) ?? "txt",
        title: path.basename(target).replace(/\.[^.]+$/, ""),
      })
    ).bytes,
  );
  created.dispose();
  // From here on the file exists: a failure is a partial result that says so.
  const mode = requestedMode === "suggest" ? "suggesting" : "editing";
  try {
    return await setUpNewDocument(args, target, resolved, format, mode, steps);
  } catch (error) {
    const problem =
      error instanceof CliError
        ? problemOf(error)
        : (systemProblem(error) ?? {
            code: "internal",
            message: (error as Error)?.message ?? String(error),
            hint: null,
            details: null,
          });
    // Still uncertain, or still a bug, with the file written: not merely
    // partly done.
    const exit =
      UNCERTAIN.has(problem.code) || problem.code === "internal" ? 70 : 2;
    throw fail(
      `${shellArg(target)} was created, but setting it up failed: ${problem.message}`,
      problem.code,
      `revise status ${shellArg(target)}`,
      problem.details ?? undefined,
      {
        exit,
        data: {
          output: resolved,
          outputWritten: true,
          format,
          mode: modeName(mode),
          checkpoint: null,
          summary: null,
        },
      },
    );
  }
}

/** The tool calls `new` runs for its page-setup flags, checked up front
 * against the tools' own schemas. */
function newDocumentSteps(
  args: ParsedArgs,
  probe: ReviseDocument,
): Array<[string, ReviseJsonObject]> {
  const layout = probe.tools
    .definitions()
    .find((tool) => tool.name === "set_page_layout");
  const allowed = (property: string) =>
    ((layout?.inputSchema.properties as Record<string, { enum?: string[] }>)?.[
      property
    ]?.enum ?? []) as string[];
  const layoutInput: ReviseJsonObject = {};
  // --page is the 3.0 name: render and search use it for a page number.
  const page = (
    flagString(args, "page-size") ?? flagString(args, "page")
  )?.toLowerCase();
  if (page !== undefined) {
    if (!allowed("pageSize").includes(page)) {
      throw fail(
        `--page-size must be one of: ${allowed("pageSize").join(", ")}.`,
        "usage",
        "revise new --help",
      );
    }
    layoutInput.pageSize = page;
  }
  const orientation = flagString(args, "orientation");
  if (orientation !== undefined) {
    if (!allowed("orientation").includes(orientation)) {
      throw fail(
        `--orientation must be ${allowed("orientation").join(" or ")}.`,
        "usage",
        "revise new --help",
      );
    }
    layoutInput.orientation = orientation;
  }
  const margins = flagString(args, "margins");
  if (margins !== undefined) {
    if (!/^\d+(\.\d+)?(in|cm|px)$/.test(margins)) {
      throw fail(
        "--margins needs a length with a unit: 1in, 2.5cm, or 96px.",
        "usage",
        "revise new --help",
      );
    }
    layoutInput.margin = margins;
  }
  const steps: Array<[string, ReviseJsonObject]> = [];
  if (Object.keys(layoutInput).length)
    steps.push(["set_page_layout", layoutInput]);
  for (const side of ["header", "footer"] as const) {
    const zones: ReviseJsonObject = {};
    for (const zone of ["left", "center", "right"] as const) {
      const value =
        flagString(args, `${side}-${zone}`) ??
        (zone === "center" ? flagString(args, side) : undefined);
      if (value !== undefined) zones[zone] = value;
    }
    if (Object.keys(zones).length)
      steps.push(["set_header_footer", { side, ...zones }]);
  }
  return steps;
}

async function setUpNewDocument(
  args: ParsedArgs,
  target: string,
  resolved: string,
  format: string,
  mode: CliDocumentMeta["mode"],
  steps: Array<[string, ReviseJsonObject]>,
): Promise<Outcome> {
  const stale = sessionNameFor(resolved);
  if (sessionExists(stale)) {
    fs.rmSync(documentDir(stale), { recursive: true, force: true });
  }
  const fileArgs = { ...args, file: target };
  let loaded = await loadDocument(fileArgs);
  // A new document starts in edit mode: there is nothing to track yet.
  if (loaded.meta.mode !== mode) {
    loaded.meta.mode = mode;
    writeMeta(loaded.meta);
    loaded = await loadDocument(fileArgs);
  }
  if (steps.length) {
    const setup = await mutate(
      loaded,
      "new (page setup)",
      async () => {
        for (const [tool, input] of steps) {
          const result = await loaded.doc.session.tools.executeDynamic(
            tool,
            input,
            { directMode: true },
          );
          if ("error" in result) return `${tool}: ${result.error.message}`;
        }
        return null;
      },
      () => "Page setup",
    );
    if (setup.result) throw fail(setup.result, "page_setup_failed");
  }
  const summary = summaryData((await loaded.doc.summary()).data);
  return {
    data: {
      output: resolved,
      outputWritten: true,
      format,
      mode: modeName(loaded.meta.mode),
      checkpoint: currentCheckpoint(loaded.meta),
      summary,
    },
    text: () =>
      [
        `${ui.green("Created")} ${loaded.shown} in ${modeName(loaded.meta.mode)} mode${loaded.meta.mode === "editing" ? " (edits apply directly)" : ""}.`,
        summaryText(summary),
        nextLine(
          "Next",
          `revise read ${loaded.shown}`,
          `revise render ${loaded.shown}`,
        ),
      ].join("\n"),
  };
}

async function cmdStatus(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const { meta, doc } = loaded;
  const status = await doc.status();
  const checkpoint = currentCheckpoint(meta);
  const fidelity = status.fidelity;
  const sourceFidelity = fidelity
    ? {
        fidelity: fidelity.fidelity,
        regeneratedBlockIds: fidelity.regeneratedBlockIds,
        changedDomains: fidelity.changedDomains,
        convertedReason: fidelity.convertedReason ?? null,
      }
    : null;
  return {
    data: {
      file: meta.file,
      format: formatName(meta.sourceFormat),
      mode: modeName(meta.mode),
      blockCount: status.blocks,
      wordCount: status.words,
      trackedChangeCount: status.pendingSuggestions,
      checkpoint,
      sourceFidelity,
    },
    text: () => {
      const row = (label: string, value: string) =>
        `${ui.dim(label.padEnd(16))}${value}`;
      const fidelityText = !sourceFidelity
        ? null
        : sourceFidelity.fidelity === "exact"
          ? "exact (export reproduces the original file)"
          : sourceFidelity.fidelity === "grafted"
            ? `grafted (${plural(sourceFidelity.regeneratedBlockIds.length, "edited block")} rewritten; everything else kept from the original file)`
            : "converted (export writes the file fresh)";
      return [
        ui.bold(loaded.shown),
        row(
          "Mode",
          meta.mode === "suggesting"
            ? "suggest (edits become tracked changes)"
            : "edit (edits apply directly)",
        ),
        row(
          "Size",
          `${plural(status.blocks, "block")}, ${plural(status.words, "word")}`,
        ),
        row(
          "Tracked changes",
          status.pendingSuggestions
            ? ui.yellow(String(status.pendingSuggestions))
            : "none",
        ),
        row("Checkpoint", String(checkpoint)),
        ...(fidelityText ? [row("Word fidelity", fidelityText)] : []),
      ].join("\n");
    },
  };
}

async function cmdRead(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const ids = blockIdsFlag(args, "read");
  const view = viewFlag(args, "tracked");
  if (ids) {
    return runTool(
      loaded,
      "read_specific_blocks",
      { block_ids: ids },
      { view },
    );
  }
  const index = strictInt(args.positional[0] ?? "0", "The block index");
  return runTool(loaded, "read_blocks_from_index", { index }, { view });
}

async function cmdSearch(args: ParsedArgs): Promise<Outcome> {
  if (!args.positional.length) throw usageError("search");
  const view = viewFlag(args, "tracked");
  const loaded = await loadDocument(args);
  // The search tool reads the accepted text. What it can't see, matches in
  // deleted text (tracked) or the whole original, comes from grep's view of
  // the same blocks, matched the way the tool matches: ignoring case, with
  // "regex:" for a pattern.
  const textMatches = (only: "deleted" | null) =>
    args.positional.flatMap((query) => {
      const matcher = query.startsWith("regex:")
        ? safeRegExp(query.slice(6))
        : query;
      return loaded.doc
        .grep(matcher, { ignoreCase: true, view: "original" })
        .filter((hit) =>
          only ? hit.matches.some((m) => m.change === only) : true,
        )
        .map((hit) => ({
          query,
          blockId: hit.id,
          blockIndex: hit.blockIndex,
          text: hit.text,
          changeIds: [
            ...new Set(
              hit.matches
                .filter((m) => m.change === "deleted")
                .map((m) => m.changeId ?? ""),
            ),
          ].filter(Boolean),
        }));
    });
  const viewLines = (
    matches: ReturnType<typeof textMatches>,
    heading: string,
  ) =>
    matches.length
      ? [
          "",
          heading,
          ...matches.map(
            (m) =>
              `${ui.cyan(m.blockId)}  ${highlight(m.text, m.query.startsWith("regex:") ? safeRegExp(m.query.slice(6)) : m.query, true)}`,
          ),
        ]
      : [];
  if (view === "original") {
    const matches = textMatches(null);
    rememberReadBlocks(
      loaded.meta,
      matches.map((m) => `<p id="${m.blockId}">`).join(""),
    );
    return {
      data: {
        ...toolData(loaded, "search_document", {
          saved: false,
          checkpoint: currentCheckpoint(loaded.meta),
          message: null,
          changeIds: [],
          html: null,
          result: null,
        }),
        view,
        viewMatches: matches,
      },
      text: () =>
        matches.length
          ? [
              ...viewLines(
                matches,
                `${plural(matches.length, "block")} match in the original text (before the pending changes):`,
              ).slice(1),
              "",
              ui.dim(
                "No search_result_id: it names blocks of the current text. Search without --view original for one.",
              ),
            ].join("\n")
          : `No matches in the original text for ${args.positional.map((q) => `"${q}"`).join(", ")}.`,
    };
  }
  const outcome = await runTool(
    loaded,
    "search_document",
    {
      queries: args.positional,
      // --page is the 3.0 name.
      page: flagInt(args, "results-page") ?? flagInt(args, "page") ?? 1,
    },
    { view },
  );
  if (view !== "tracked" || outcome.error) return outcome;
  const deleted = textMatches("deleted");
  rememberReadBlocks(
    loaded.meta,
    deleted.map((m) => `<p id="${m.blockId}">`).join(""),
  );
  return {
    ...outcome,
    data: { ...outcome.data, viewMatches: deleted },
    text: () =>
      [
        outcome.text(),
        ...viewLines(
          deleted,
          `Also in deleted text (pending changes; ${ui.dim("[-deleted-]{+inserted+}")}):`,
        ),
      ].join("\n"),
  };
}

/** A search pattern as a case-insensitive RegExp; an invalid one is
 * matched literally, as the search tool does. */
function safeRegExp(pattern: string): RegExp | string {
  try {
    return new RegExp(pattern, "gi");
  } catch {
    return pattern;
  }
}

async function cmdOutline(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const limit = flagInt(args, "limit");
  const data = await loaded.doc.outline({
    limit,
    offset: flagInt(args, "offset"),
  });
  const sections = (data.sections ?? []).map((section) => ({
    title: section.title,
    level: section.headingLevel,
    startIndex: section.startIndex,
    endIndex: section.endIndex,
    startBlockId: section.startBlockId,
    endBlockId: section.endBlockId,
    blockCount: section.blockCount,
    wordCount: section.wordCount,
  }));
  const nextOffset = data.hasMoreSections
    ? (data.nextSectionOffset ?? null)
    : null;
  return {
    data: {
      file: loaded.meta.file,
      blockCount: data.totalBlocks,
      wordCount: data.wordCount,
      sectionCount: data.sectionCount ?? sections.length,
      sections,
      nextOffset,
    },
    text: () => {
      const lines = [
        `${plural(data.totalBlocks, "block")}, ${plural(data.wordCount, "word")}.`,
      ];
      for (const section of sections) {
        lines.push(
          `${"  ".repeat(section.level - 1)}${ui.bold(section.title)}  ${ui.dim(`blocks ${section.startIndex}–${section.endIndex}, ${plural(section.wordCount, "word")}`)}  ${ui.cyan(section.startBlockId)}`,
        );
      }
      if (!sections.length) lines.push("No headings.");
      if (nextOffset !== null) {
        lines.push(
          nextLine(
            "More",
            `revise outline ${loaded.shown} --offset ${nextOffset}${limit !== undefined ? ` --limit ${limit}` : ""}`,
          ),
        );
      }
      return lines.join("\n");
    },
  };
}

async function cmdTools(args: ParsedArgs): Promise<Outcome> {
  // Definitions do not depend on the document; any session will do.
  const ydoc = await fileToYDoc(new TextEncoder().encode("\n"), "tools.md");
  const session = await createServerDocumentSession(ydoc, {
    documentId: "tools",
  });
  const definitions = session.tools.getDefinitions();
  const name = args.positional[0];
  if (name) {
    const definition = definitions.find((tool) => tool.name === name);
    if (!definition)
      throw fail(`Unknown tool "${name}".`, "unknown_tool", "revise tools");
    return {
      data: {
        name: definition.name,
        description: definition.description,
        inputSchema: definition.inputSchema,
      },
      text: () =>
        toolHelpText(
          definition,
          `revise ${definition.name} <file> ['<json>' | - | @<file.json>]`,
        ),
    };
  }
  const tools = definitions.map(({ name: n, description }) => ({
    name: n,
    summary: description.split(/(?<=\.)\s/)[0],
  }));
  return {
    data: { tools },
    text: () =>
      [
        ...tools.map(
          (tool) => `${ui.cyan(tool.name.padEnd(24))} ${tool.summary}`,
        ),
        "",
        nextLine("Details", "revise tools <name>"),
        nextLine("Run", "revise call <file> <name> '<json>'"),
      ].join("\n"),
  };
}

async function cmdCall(args: ParsedArgs): Promise<Outcome> {
  const name = args.positional[0];
  if (!name) throw usageError("call");
  if (args.positional[1] === undefined) {
    const definition = await toolDefinition(name);
    const required = (definition?.inputSchema as JsonSchemaNode | undefined)
      ?.required;
    if (definition && required?.length) {
      throw fail(
        `${name} needs its input as JSON:\n${stripStyle(inputFields(definition.inputSchema).join("\n"))}`,
        "usage",
        `revise ${name} --help`,
        { inputSchema: definition.inputSchema },
      );
    }
  }
  const loaded = await loadDocument(args);
  return runTool(loaded, name, readInputJson(args.positional[1]), {
    direct: args.flags.direct === true,
  });
}

async function cmdMode(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const { meta } = loaded;
  const requested = args.positional[0];
  const before = meta.mode;
  if (requested) {
    if (!["suggest", "edit"].includes(requested)) throw usageError("mode");
    meta.mode = requested === "suggest" ? "suggesting" : "editing";
    if (meta.mode !== before) writeMeta(meta);
  }
  const changed = meta.mode !== before;
  return {
    data: { file: meta.file, mode: modeName(meta.mode), changed },
    text: () =>
      `${changed ? "Mode is now" : "Mode:"} ${ui.bold(modeName(meta.mode))} — ${
        meta.mode === "suggesting"
          ? "edits become tracked changes."
          : "edits apply directly."
      }`,
  };
}

function clip(text: string, max = 80): string {
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length > max ? `${flat.slice(0, max - 1)}…` : flat;
}

const BEFORE_FIRST_HEADING = "(before the first heading)";

async function cmdChanges(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const section = flagString(args, "section");
  const blocksFlag = flagString(args, "blocks");
  const blocks = blocksFlag ? parseBlockQuery(blocksFlag) : undefined;
  const changes = await loaded.doc.changes({ section, blocks });
  const filtered = Boolean(section || blocks);
  const details = args.flags.details === true || filtered;
  const entries = changes.map((c) => ({
    changeId: c.id,
    type: c.type,
    author: c.author,
    blockId: c.blockId ?? null,
    blockIndex: c.blockIndex ?? null,
    section: c.section ?? null,
    deleted: c.deleted,
    inserted: c.inserted,
  }));
  const next = () =>
    nextLine(
      "Next",
      `revise accept ${loaded.shown} <change-id...>`,
      `revise reject ${loaded.shown} <change-id...>`,
    ) +
    `\n${ui.dim(`      (or --all, --section <number or title>, --blocks <id,id or first..last>; add --dry-run to preview)`)}`;
  return {
    data: {
      file: loaded.meta.file,
      section: section ?? null,
      blocks: blocksFlag ?? null,
      trackedChangeCount: entries.length,
      changes: entries,
    },
    text: () => {
      if (!entries.length) {
        return section
          ? `No tracked changes in section "${section}"${blocks ? " and those blocks" : ""}.`
          : blocks
            ? "No tracked changes in those blocks."
            : "No tracked changes.";
      }
      if (!details && entries.length > 30) {
        return [
          `${plural(entries.length, "tracked change")}, by section:`,
          ...bySectionLines(entries),
          "",
          nextLine(
            "Details",
            `revise changes ${loaded.shown} --section <number or title>`,
            `revise changes ${loaded.shown} --details`,
          ),
          next(),
        ].join("\n");
      }
      const lines: string[] = [`${plural(entries.length, "tracked change")}:`];
      let lastSection: string | null | undefined;
      for (const c of entries) {
        if (c.section !== lastSection || lines.length === 1) {
          lines.push(ui.bold(c.section ?? BEFORE_FIRST_HEADING));
          lastSection = c.section;
        }
        const what = [
          c.deleted ? ui.red(`-"${clip(c.deleted)}"`) : "",
          c.inserted ? ui.green(`+"${clip(c.inserted)}"`) : "",
        ]
          .filter(Boolean)
          .join(" ");
        lines.push(
          `  ${ui.cyan(c.changeId)}  ${what || ui.dim(`(${c.type})`)}  ${ui.dim(`by ${c.author}`)}`,
        );
      }
      lines.push("", next());
      return lines.join("\n");
    },
  };
}

/** "  12  8. Allocation of risk" per section, in document order. */
function bySectionLines(changes: Array<{ section: string | null }>): string[] {
  const bySection = new Map<string, number>();
  for (const c of changes) {
    const key = c.section ?? BEFORE_FIRST_HEADING;
    bySection.set(key, (bySection.get(key) ?? 0) + 1);
  }
  return [...bySection].map(
    ([title, n]) => `  ${String(n).padStart(4)}  ${title}`,
  );
}

async function cmdSettle(args: ParsedArgs, accept: boolean): Promise<Outcome> {
  const action = accept ? "accept" : "reject";
  const loaded = await loadDocument(args);
  const all = args.flags.all === true;
  const section = flagString(args, "section");
  const exceptSection = flagString(args, "except-section");
  const blocksFlag = flagString(args, "blocks");
  const blocks = blocksFlag ? parseBlockQuery(blocksFlag) : undefined;
  if (
    !all &&
    !section &&
    !exceptSection &&
    !blocks &&
    !args.positional.length
  ) {
    throw usageError(action);
  }
  const selection = {
    ids: args.positional,
    all,
    section,
    exceptSection,
    blocks,
  };
  const base = {
    file: loaded.meta.file,
    action,
  };
  if (args.flags["dry-run"] === true) {
    // What would be accepted or rejected, without changing anything.
    const [chosen, total] = await Promise.all([
      loaded.doc.preview(selection),
      loaded.doc.changes(),
    ]);
    const pending = new Set(total.map((c) => c.id));
    const notFound = args.positional.filter((id) => !pending.has(id));
    const data = {
      ...base,
      dryRun: true,
      resolved: [],
      wouldResolve: chosen.map((c) => c.id),
      notFound,
      failed: [],
      saved: false,
      checkpoint: currentCheckpoint(loaded.meta),
    };
    const text = () =>
      [
        `Would ${action} ${chosen.length} of ${plural(total.length, "tracked change")}:`,
        ...bySectionLines(chosen.map((c) => ({ section: c.section ?? null }))),
        ...(notFound.length
          ? [`Not pending: ${notFound.map((id) => ui.cyan(id)).join(", ")}`]
          : []),
        "Nothing changed (dry run).",
      ].join("\n");
    if (notFound.length) {
      return {
        data,
        text,
        exit: 1,
        error: {
          code: "change_not_found",
          message: `${plural(notFound.length, "requested change is", "requested changes are")} not pending: ${notFound.join(", ")}.`,
          hint: `revise changes ${loaded.shown}`,
          details: { notFound },
        },
      };
    }
    return { data, text };
  }
  const { result, checkpoint } = await mutate(
    loaded,
    action,
    () =>
      accept ? loaded.doc.accept(selection) : loaded.doc.reject(selection),
    (r) =>
      `${accept ? "Accepted" : "Rejected"} ${plural(r.resolved.length, "tracked change")}`,
  );
  const failed = result.unresolved.map((changeId) => ({
    changeId,
    code: "not_resolved",
    message: `Could not ${action} this change.`,
  }));
  const data = {
    ...base,
    dryRun: false,
    resolved: result.resolved,
    wouldResolve: [],
    notFound: result.missing,
    failed,
    saved: checkpoint !== null,
    checkpoint: currentCheckpoint(loaded.meta),
  };
  const done = accept ? "Accepted" : "Rejected";
  const failedOutright =
    !result.resolved.length && (result.missing.length || failed.length);
  const text = () =>
    failedOutright
      ? ""
      : [
          result.resolved.length
            ? `${ui.green(done)} ${plural(result.resolved.length, "tracked change")}.`
            : section || exceptSection || blocks
              ? "No tracked changes matched."
              : "No tracked changes to " + action + ".",
          ...(checkpoint !== null ? [savedLine(loaded, checkpoint)] : []),
        ].join("\n");
  if (result.missing.length || failed.length) {
    const problems = [
      result.missing.length ? `not pending: ${result.missing.join(", ")}` : "",
      failed.length
        ? `could not ${action}: ${failed.map((f) => f.changeId).join(", ")}`
        : "",
    ].filter(Boolean);
    return {
      data,
      text,
      exit: result.resolved.length ? 2 : 1,
      error: {
        code: result.missing.length
          ? "change_not_found"
          : "change_not_resolved",
        message:
          `${problems.join("; ")}.`.replace(/^./, (c) => c.toUpperCase()) +
          (checkpoint !== null
            ? ` ${plural(result.resolved.length, "other change was", "other changes were")} ${accept ? "accepted" : "rejected"} and saved.`
            : " Nothing was saved."),
        hint: `revise changes ${loaded.shown}`,
        details: { notFound: result.missing, failed },
      },
    };
  }
  return { data, text };
}

async function cmdDiff(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const since = flagInt(args, "since");
  const label =
    since !== undefined
      ? `Checkpoint ${since} → now, tracked changes accepted`
      : "Tracked changes, before → after";
  const { text, changed } = await loaded.doc.diff(
    since !== undefined ? readCheckpoint(loaded.meta.name, since) : undefined,
  );
  return {
    data: {
      file: loaded.meta.file,
      since: since ?? null,
      changedParagraphCount: changed,
      text,
    },
    text: () =>
      changed
        ? `${label}: ${plural(changed, "paragraph")} differ${changed === 1 ? "s" : ""}.\n\n${colorDiff(text)}`
        : `${label}: no text differences.`,
  };
}

/** Color [-removed-] red and {+added+} green; the markers stay. */
function colorDiff(text: string): string {
  if (!ui.enabled) return text;
  return text
    .replace(/\[-([\s\S]*?)-\]/g, (m) => ui.red(m))
    .replace(/\{\+([\s\S]*?)\+\}/g, (m) => ui.green(m));
}

async function cmdHistory(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const { meta } = loaded;
  const checkpoints = meta.checkpoints.map((c) => ({
    checkpoint: c.n,
    at: c.at,
    command: c.command,
    summary: c.summary,
  }));
  const current = currentCheckpoint(meta);
  return {
    data: { file: meta.file, checkpoint: current, checkpoints },
    text: () =>
      [
        ui.dim(
          `${"#".padStart(3)}  ${"Time (UTC)".padEnd(16)}  ${"Command".padEnd(22)} Summary`,
        ),
        ...checkpoints.map(
          (c) =>
            `${(c.checkpoint === current ? ui.bold : (t: string) => t)(String(c.checkpoint).padStart(3))}  ${c.at.slice(0, 16).replace("T", " ")}  ${c.command.padEnd(22)} ${c.summary.split("\n")[0]}`,
        ),
        "",
        nextLine(
          "Undo",
          `revise undo ${loaded.shown}`,
          `revise restore ${loaded.shown} <checkpoint>`,
        ),
      ].join("\n"),
  };
}

async function cmdUndo(args: ParsedArgs, target?: number): Promise<Outcome> {
  const loaded = await loadDocument(args);
  const { meta } = loaded;
  assertWritable(meta);
  const last = currentCheckpoint(meta);
  const steps =
    target === undefined ? strictInt(args.positional[0] ?? "1", "steps") : 0;
  if (steps < 0) throw fail("Undo steps must be positive.", "usage");
  const to = target ?? Math.max(0, last - steps);
  if (to === last) {
    return {
      data: {
        file: meta.file,
        checkpoint: last,
        from: last,
        discarded: [],
        saved: false,
      },
      text: () => `Already at checkpoint ${last}. Nothing changed.`,
    };
  }
  const restored = checkpointAt(meta, to);
  const discarded = meta.checkpoints.filter((c) => c.n > to).map((c) => c.n);
  // An xml-patch checkpoint's truth is its package; any other is its model
  // exported onto its package.
  const bytes =
    restored.modelStale && restored.package
      ? restored.package
      : await exportBytes(meta, decodeYDoc(restored.state), restored.package);
  await saveJournaled(meta, bytes, (fileHash) =>
    stageTruncate(meta, to, fileHash),
  );
  return {
    data: {
      file: meta.file,
      checkpoint: to,
      from: last,
      discarded,
      saved: true,
    },
    text: () =>
      `${ui.green("Restored")} ${loaded.shown} to checkpoint ${to}; discarded ${plural(discarded.length, "later checkpoint")}.`,
  };
}

async function cmdRender(args: ParsedArgs): Promise<Outcome> {
  const loaded = await loadDocument(args);
  let startPage = flagInt(args, "page") ?? 1;
  let pageCount = 1;
  const range = flagString(args, "pages");
  if (range) {
    const match = /^(\d+)(?:-(\d+))?$/.exec(range);
    if (!match) throw fail("--pages expects N or N-M.", "usage");
    startPage = Number(match[1]);
    pageCount = Math.max(1, Number(match[2] ?? match[1]) - startPage + 1);
  }
  if (pageCount > 6) throw fail("Render at most 6 pages per call.", "usage");
  const view = (flagString(args, "view") ?? "tracked") as
    | "tracked"
    | "accepted"
    | "original";
  if (!["tracked", "accepted", "original"].includes(view)) {
    throw fail("--view must be tracked, accepted, or original.", "usage");
  }
  const result = await loaded.doc.renderPages({
    startPage,
    pageCount,
    view,
    layout: args.flags.layout === true,
    quality:
      args.flags.detailed === true || pageCount === 1 ? "detailed" : "compact",
  });
  const extension = result.mimeType === "image/png" ? "png" : "jpg";
  const outPath = path.resolve(
    flagString(args, "out") ??
      path.join(
        documentDir(loaded.meta.name),
        "renders",
        `page-${result.pages[0]}${result.pages.length > 1 ? `-${result.pages.at(-1)}` : ""}-${view}${args.flags.layout ? "-layout" : ""}.${extension}`,
      ),
  );
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  try {
    fs.writeFileSync(outPath, result.bytes);
  } catch (error) {
    throw fail(
      `Could not write ${outPath}: ${(error as Error).message}`,
      "write_failed",
    );
  }
  return {
    data: {
      file: loaded.meta.file,
      output: outPath,
      outputWritten: true,
      mimeType: result.mimeType,
      width: result.width,
      height: result.height,
      pages: result.pages,
      totalPages: result.totalPages,
      view,
      description: result.description,
      placeholderImageCount: result.placeholderImageCount,
      failedImageCount: result.failedImageCount,
    },
    text: () =>
      [
        `${ui.green("Rendered")} ${result.pages.length > 1 ? "pages" : "page"} ${result.pages.join(", ")} of ${result.totalPages} → ${outPath}`,
        result.description,
        ...(result.placeholderImageCount || result.failedImageCount
          ? [
              ui.yellow(
                `Images not drawn: ${result.placeholderImageCount} placeholder, ${result.failedImageCount} failed.`,
              ),
            ]
          : []),
      ].join("\n"),
  };
}

async function cmdExport(args: ParsedArgs): Promise<Outcome> {
  const target = args.positional[0];
  if (!target) throw usageError("export");
  const format = formatForPath(target);
  if (!format) {
    throw fail(
      "Export to .docx, .pdf, .md, .html, or .txt. (For a picture of the pages, use revise render.)",
      "unsupported_format",
    );
  }
  const loaded = await loadDocument(args);
  const out = canonicalPath(target);
  if (out === loaded.meta.file) {
    throw fail(
      `${target} is the file itself, and edits already save it. Export to a different file.`,
      "same_file",
    );
  }
  assertCanOverwrite(args, out);
  const result = await loaded.doc.export({
    format,
    view: viewFlag(args, "tracked"),
  });
  writeOutput(args, out, result.bytes);
  // Describe the file that was written, not the source: re-read a .docx.
  let summary: ReturnType<typeof summaryData> | null = null;
  if (format === "docx") {
    try {
      const written = await ReviseDocument.open({
        bytes: result.bytes,
        filename: out,
      });
      summary = summaryData((await written.summary()).data);
      written.dispose();
    } catch (error) {
      warn(
        "summary_unavailable",
        `Wrote ${shellArg(target)}, but could not read it back to describe it (${(error as Error).message}).`,
      );
    }
  }
  const docxExport = result.fidelity
    ? {
        method: result.fidelity.method,
        exact: result.fidelity.exact,
        regeneratedBlockCount: result.fidelity.regeneratedBlocks,
        fallbackCode: result.fidelity.fallbackCode ?? null,
      }
    : null;
  for (const message of result.warnings) {
    if (/can't hold tracked changes/.test(message)) {
      const pending = (await loaded.doc.status()).pendingSuggestions;
      warn("changes_accepted", message, { trackedChangeCount: pending });
    } else warn("export_note", message);
  }
  // What the written file holds: a .docx as revise reads it back (Word
  // stores a replacement as a deletion and an insertion); other formats
  // can't hold tracked changes.
  const trackedChangeCount =
    format === "docx" ? (summary?.trackedChangeCount ?? null) : 0;
  return {
    data: {
      file: loaded.meta.file,
      output: out,
      outputWritten: true,
      format: formatName(format),
      byteCount: result.bytes.length,
      final: result.mode === "final",
      view: result.view,
      trackedChangeCount,
      docxExport,
      summary,
    },
    text: () =>
      [
        `${ui.green("Wrote")} ${shellArg(target)} (${sizeText(result.bytes.length)}${result.view === "original" ? ", as it was before the tracked changes" : result.mode === "final" ? ", tracked changes accepted" : ""}).`,
        ...(summary ? [summaryText(summary)] : []),
        ...(format === "docx" && trackedChangeCount
          ? [
              `Keeps the tracked changes and comments for review. ${ui.dim("For a clean copy, add --view accepted.")}`,
            ]
          : []),
      ].join("\n"),
  };
}

/** Write an output file (not the session's own), atomically. Without
 * --force the write itself refuses to replace a file, even one that
 * appeared a moment ago: a hard link to the finished temp file either
 * creates the name or fails. */
function writeOutput(args: ParsedArgs, file: string, bytes: Uint8Array): void {
  try {
    if (args.flags.force === true) {
      writeFileAtomic(file, bytes);
      return;
    }
    const temp = path.join(
      path.dirname(file),
      `.${path.basename(file)}.${process.pid}.${Date.now()}.tmp`,
    );
    fs.writeFileSync(temp, bytes);
    try {
      fs.linkSync(temp, file);
    } finally {
      fs.rmSync(temp, { force: true });
    }
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") {
      throw wouldOverwrite(file);
    }
    throw fail(
      `Could not write ${path.basename(file)}: ${(error as Error).message}`,
      "write_failed",
    );
  }
}

async function cmdConvert(args: ParsedArgs): Promise<Outcome> {
  const [input, target] = args.positional;
  if (!input || !target) throw usageError("convert");
  const fromStdin = input === "-";
  const toStdout = target === "-";
  const to = flagString(args, "to");
  if (to !== undefined && !toStdout) {
    throw fail(
      "--to names the format of stdout; a file's extension already names its format.",
      "usage",
      `revise convert ${shellArg(input)} - --to ${to}`,
    );
  }
  const format = toStdout
    ? formatForPath(`stdout.${to ?? "md"}`)
    : formatForPath(target);
  if (!format) {
    throw fail(
      toStdout
        ? "--to takes md, html, txt, docx, or pdf."
        : "Convert to .docx, .md, .html, .txt, or .pdf.",
      "unsupported_format",
      toStdout ? `revise convert ${shellArg(input)} - --to md` : undefined,
    );
  }
  if (toStdout && jsonOutput) {
    throw fail(
      "--json prints its report on stdout, so the document can't go there too. Write it to a file.",
      "usage",
      `revise convert ${shellArg(input)} out.${formatName(format)} --json`,
    );
  }
  if (toStdout && (format === "docx" || format === "pdf") && io.stdoutIsTTY) {
    throw fail(
      `A .${format} is binary; it won't print to a terminal. Redirect it to a file.`,
      "usage",
      `revise convert ${shellArg(input)} - --to ${format} > out.${format}`,
    );
  }
  const source = fromStdin ? await readStdinDocument(args) : null;
  const out = toStdout ? null : canonicalPath(target);
  if (out !== null) {
    if (!fromStdin && out === canonicalPath(input)) {
      throw fail("Convert to a different file.", "same_file");
    }
    assertCanOverwrite(args, out);
  }
  if (!fromStdin && flagString(args, "from") !== undefined) {
    throw fail(
      "--from names the format of stdin; a file's extension already names its format.",
      "usage",
      `revise convert ${shellArg(input)} ${shellArg(target)}`,
    );
  }
  // Only a .docx can hold tracked changes.
  const asked = viewFlag(args, "tracked");
  const view = asked === "tracked" && format !== "docx" ? "accepted" : asked;
  const final = view !== "tracked";
  const bytes = await convert(
    source
      ? { bytes: source.bytes, filename: `document.${source.format}` }
      : input,
    format,
    { view },
  );
  if (out === null) await (io.writeStdout ?? defaultIo().writeStdout!)(bytes);
  else writeOutput(args, out, bytes);
  return {
    data: {
      input: fromStdin ? null : canonicalPath(input),
      inputFormat: source ? source.format : inputFormatOfPath(input),
      output: out,
      outputWritten: out !== null,
      format: formatName(format),
      byteCount: bytes.length,
      final,
      view,
    },
    // The document is the output; a report on stdout would corrupt it.
    text: () =>
      out === null
        ? ""
        : `${ui.green("Wrote")} ${shellArg(target)} (${sizeText(bytes.length)})${
            source ? ` from stdin (${source.format})` : ""
          }.`,
  };
}

async function cmdCompare(args: ParsedArgs): Promise<Outcome> {
  const [originalFile, revisedFile] = args.positional;
  if (!originalFile || !revisedFile) throw usageError("compare");
  // The redline is a file like any other: written now, then edited by name.
  const outArg =
    flagString(args, "out") ??
    `${revisedFile.replace(/\.[^.]+$/, "")}-redline.docx`;
  const out = canonicalPath(outArg);
  if (!/\.docx$/i.test(out)) {
    throw fail("compare writes a .docx redline: --out <name>.docx", "usage");
  }
  if ([originalFile, revisedFile].some((f) => canonicalPath(f) === out)) {
    throw fail(
      "--out must be a new file, not one of the documents being compared.",
      "same_file",
    );
  }
  assertCanOverwrite(args, out);
  const { document, stats, untrackedRegions, verified } = await compare(
    originalFile,
    revisedFile,
    {
      ...((flagString(args, "author") ?? process.env.REVISE_AUTHOR)
        ? { author: flagString(args, "author") ?? process.env.REVISE_AUTHOR }
        : {}),
    },
  );
  const author = document.author ?? path.basename(revisedFile);
  const name = sessionNameFor(out);
  const record = {
    name,
    file: out,
    fileHash: "",
    documentId: `cli-${name}`,
    source: canonicalPath(revisedFile),
    sourceFormat: "docx",
    createdAt: new Date().toISOString(),
    mode: "suggesting" as const,
    author,
  };
  // The file first (never replacing one without --force), then its history.
  const bytes = await exportBytes(
    { ...record, checkpoints: [] },
    document.ydoc,
    document.package,
  );
  writeOutput(args, out, bytes);
  const fileHash = hashFileBytes(bytes);
  createDocumentRecord(
    { ...record, fileHash },
    encodeYDoc(document.ydoc),
    document.package,
  );
  document.dispose();
  if (untrackedRegions.length) {
    warn(
      "untracked_regions",
      `Also different, but can't carry tracked changes (taken from the revised file): ${untrackedRegions.join(", ")}.`,
      { regions: untrackedRegions },
    );
  }
  const shownOut = shellArg(outArg);
  const verification = {
    scope: "normalized_text",
    acceptAllMatchesRevised: verified.acceptAllGivesRevised,
    rejectAllMatchesOriginal: verified.rejectAllGivesOriginal,
  };
  const passed =
    verified.acceptAllGivesRevised && verified.rejectAllGivesOriginal;
  const outcome: Outcome = {
    data: {
      original: canonicalPath(originalFile),
      revised: canonicalPath(revisedFile),
      output: out,
      outputWritten: true,
      author,
      inlineChangeCount: stats.inlineChanges,
      insertedBlockCount: stats.insertedBlocks,
      deletedBlockCount: stats.deletedBlocks,
      unmarkedBlockCount: stats.unmarkedBlocks,
      untrackedRegions,
      verification,
    },
    text: () =>
      [
        `${ui.green("Wrote")} ${shownOut}: ${shellArg(originalFile)} → ${shellArg(revisedFile)} as tracked changes by "${author}".`,
        `${plural(stats.inlineChanges, "inline change")}, ${plural(stats.insertedBlocks, "inserted block")}, ${plural(stats.deletedBlocks, "deleted block")}${stats.unmarkedBlocks ? `, ${plural(stats.unmarkedBlocks, "changed block")} that can't carry a tracked change` : ""}.`,
        ...(passed
          ? [
              ui.dim(
                `Checked: accepting every change gives the revised text; rejecting every change gives the original text.`,
              ),
            ]
          : []),
        nextLine(
          "Next",
          `revise changes ${shownOut}`,
          `revise diff ${shownOut}`,
        ),
      ].join("\n"),
  };
  if (!passed) {
    // A redline that fails its own check is written, but not done.
    outcome.exit = 2;
    outcome.error = {
      code: "redline_unverified",
      message: `${shownOut} was written, but ${[
        verified.acceptAllGivesRevised
          ? ""
          : "accepting every change does not give the revised text",
        verified.rejectAllGivesOriginal
          ? ""
          : "rejecting every change does not give the original text",
      ]
        .filter(Boolean)
        .join(", and ")}.`,
      hint: `revise diff ${shownOut}`,
      details: { verification },
    };
  }
  return outcome;
}

async function cmdGrep(args: ParsedArgs): Promise<Outcome> {
  const pattern = args.positional[0];
  if (!pattern) throw usageError("grep");
  const loaded = await loadDocument(args);
  const ignoreCase = args.flags["ignore-case"] === true;
  const regex = args.flags.regex === true;
  let matcher: string | RegExp = pattern;
  if (regex) {
    try {
      matcher = new RegExp(pattern, ignoreCase ? "gi" : "g");
    } catch (error) {
      throw fail(
        `Not a valid regular expression: ${(error as Error).message.replace(/^Invalid regular expression: /, "")}.`,
        "invalid_regex",
      );
    }
  }
  const view = viewFlag(args, "tracked");
  const hits = loaded.doc.grep(matcher, { ignoreCase, view });
  // Grep output delivers the blocks' text: count them as read.
  rememberReadBlocks(
    loaded.meta,
    hits.map((hit) => `<p id="${hit.id}">`).join(""),
  );
  const matchCount = hits.reduce((n, h) => n + h.count, 0);
  return {
    data: {
      file: loaded.meta.file,
      pattern,
      regex,
      ignoreCase,
      view,
      matchCount,
      blocks: hits.map((hit) => ({
        blockId: hit.id,
        blockIndex: hit.blockIndex,
        matchCount: hit.count,
        text: hit.text,
        matches: hit.matches,
      })),
    },
    text: () =>
      hits.length
        ? [
            ...hits.map(
              (hit) =>
                `${ui.cyan(hit.id)}  ${highlight(hit.text, matcher, ignoreCase)}${inChanges(hit.matches)}`,
            ),
            "",
            `${plural(matchCount, "match", "matches")} in ${plural(hits.length, "block")}.`,
            nextLine(
              "Edit",
              `revise replace-all ${loaded.shown} "<find>" "<replace>"`,
              `revise call ${loaded.shown} replace '{"id":"<block-id>","replacements":[{"find":"…","replace":"…"}]}'`,
            ),
          ].join("\n")
        : `No matches for "${pattern}".`,
  };
}

/** "  (1 in deleted text, 1 in inserted text)", or nothing. */
function inChanges(matches: Array<{ change: string | null }>): string {
  const count = (change: string) =>
    matches.filter((m) => m.change === change).length;
  const parts = [
    count("deleted") ? `${count("deleted")} in deleted text` : "",
    count("inserted") ? `${count("inserted")} in inserted text` : "",
  ].filter(Boolean);
  return parts.length ? ui.dim(`  (${parts.join(", ")})`) : "";
}

/** Matches in bold, when color is on. */
function highlight(
  text: string,
  matcher: string | RegExp,
  ignoreCase: boolean,
): string {
  if (!ui.enabled) return text;
  const re =
    typeof matcher === "string"
      ? new RegExp(
          matcher.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          ignoreCase ? "gi" : "g",
        )
      : matcher;
  return text.replace(re, (m) => (m ? ui.bold(ui.yellow(m)) : m));
}

async function cmdReplaceAll(args: ParsedArgs): Promise<Outcome> {
  const [find, replacement] = args.positional;
  if (!find || replacement === undefined) throw usageError("replace-all");
  const loaded = await loadDocument(args);
  const blocks = blockIdsFlag(args, "replace-all");
  const { result, checkpoint } = await mutate(
    loaded,
    "replace-all",
    () =>
      loaded.doc.replaceAll(find, replacement, {
        blocks,
        ignoreCase: args.flags["ignore-case"] === true,
        exact: args.flags.exact === true,
      }),
    (r) =>
      `Replaced "${find}" with "${replacement}" in ${plural(r.blocks.length - r.failures.length, "block")}`,
  );
  const data = {
    file: loaded.meta.file,
    find,
    replace: replacement,
    matchedBlockIds: result.blocks,
    replacedOccurrences: result.replaced,
    regions: result.regions,
    nearMisses: result.nearMisses,
    changeIds: result.changeIds,
    failures: result.failures,
    saved: checkpoint !== null,
    checkpoint: currentCheckpoint(loaded.meta),
  };
  const edited = result.blocks.length - result.failures.length;
  // Where it replaced, when not only in the body: "in 5 blocks (2 in the
  // header, 1 in a footnote)".
  const elsewhere = (
    [
      ["header", "the header"],
      ["footer", "the footer"],
      ["footnote", "footnotes"],
      ["endnote", "endnotes"],
    ] as const
  )
    .filter(([region]) => result.regions[region])
    .map(([region, label]) => `${result.regions[region]} in ${label}`);
  const missedLines = result.nearMisses.map(
    (m) =>
      `${ui.yellow("Not replaced:")} ${plural(m.count, "match", "matches")} of "${m.text}" (${m.reason === "case" ? "different capitalization; add --ignore-case" : "different characters; drop --exact"}).`,
  );
  const text = () => {
    if (!result.blocks.length) {
      return [`No block contains "${find}".`, ...missedLines].join("\n");
    }
    if (find === replacement || !result.replaced) {
      return result.failures.length
        ? ""
        : "Already matches: nothing to change. File not rewritten.";
    }
    return [
      `${ui.green("Replaced")} ${plural(result.replaced, "occurrence")} of "${find}" in ${plural(edited, "block")}${elsewhere.length ? ` (${elsewhere.join(", ")})` : ""}${loaded.meta.mode === "suggesting" ? " as tracked changes" : ""}.`,
      ...missedLines,
      ...(result.changeIds.length
        ? [
            result.changeIds.length <= 6
              ? `Tracked changes: ${result.changeIds.map((id) => ui.cyan(id)).join(", ")}`
              : `${plural(result.changeIds.length, "tracked change")}. ${ui.dim("List:")} ${ui.cyan(`revise changes ${loaded.shown}`)}`,
          ]
        : []),
      ...(checkpoint !== null ? [savedLine(loaded, checkpoint)] : []),
    ].join("\n");
  };
  if (result.failures.length) {
    return {
      data,
      text,
      exit: checkpoint !== null ? 2 : 1,
      error: {
        code: "replace_failed",
        message: `Could not edit ${plural(result.failures.length, "block")}: ${result.failures.map((f) => `${f.blockId} (${f.message})`).join("; ")}${checkpoint !== null ? ` ${plural(edited, "other block was", "other blocks were")} edited and saved.` : " Nothing was saved."}`,
        hint: `revise read ${loaded.shown} --blocks ${result.failures.map((f) => f.blockId).join(",")}`,
        details: { failures: result.failures },
      },
    };
  }
  return { data, text };
}

/** Package errors, with the command to run next for this file. */
async function withXmlHints<T>(
  shown: string,
  part: string,
  work: () => Promise<T>,
): Promise<T> {
  try {
    return await work();
  } catch (error) {
    if (!(error instanceof CliError) || extras.get(error)?.hint) throw error;
    const hint =
      error.code === "no_such_part"
        ? `revise xml ls ${shown}`
        : error.code === "find_not_found" || error.code === "find_ambiguous"
          ? `revise xml cat ${shown} ${shellArg(part)} --match '<text>'`
          : null;
    throw hint ? fail(error.message, error.code, hint) : error;
  }
}

async function cmdXml(args: ParsedArgs, sub: string): Promise<Outcome> {
  const part = args.positional[0];
  const xml = await import("../lib/package");
  const file = requireFile(args);
  const shown = shellArg(args.file ?? file);
  if (!/\.docx$/i.test(file)) {
    throw fail(
      `${shown} is not a .docx, so it has no package XML.`,
      "no_package",
      `revise convert ${shown} ${shellArg(path.basename(file).replace(/\.[^.]+$/, ""))}.docx`,
    );
  }
  if (sub === "patch") {
    // A patch edits the package directly: the model is not involved.
    if (!part) throw usageError("xml patch");
    // Make sure the cache exists and knows the file as it is now.
    const loaded = await loadDocument(args);
    const { meta } = loaded;
    const original = new Uint8Array(fs.readFileSync(meta.file));
    const result = await withXmlHints(shown, part, () =>
      xml.patchPackage(
        original,
        {
          part,
          find: flagString(args, "find"),
          replace: flagString(args, "replace"),
          all: args.flags.all === true,
          file: flagString(args, "file"),
        },
        (source) => fs.readFileSync(source, "utf8"),
      ),
    );
    // The file is the truth: write it and record a checkpoint whose package
    // is the patched file and whose model is re-learned before use.
    const checkpoint = await saveJournaled(meta, result.bytes, (fileHash) =>
      stageCommit(
        meta,
        readState(meta),
        {
          command: `xml patch ${part}`,
          summary: `Patched ${part} (${plural(result.replacements, "replacement")})`,
          modelStale: true,
        },
        fileHash,
        result.bytes,
      ),
    );
    return {
      data: {
        file: meta.file,
        part,
        replacementCount: result.replacements,
        saved: true,
        checkpoint,
      },
      text: () =>
        [
          `${ui.green("Patched")} ${part}: ${plural(result.replacements, "replacement")}, well-formed. Applied directly, not as a tracked change.`,
          savedLine(loaded, checkpoint),
        ].join("\n"),
    };
  }
  // ls and cat read the file itself: exactly what patch edits.
  const bytes = new Uint8Array(fs.readFileSync(file));
  if (sub === "ls") {
    const parts = await xml.listParts(bytes);
    return {
      data: {
        file,
        parts: parts.map((p) => ({ path: p.path, byteCount: p.bytes })),
      },
      text: () =>
        parts
          .map((p) => `${ui.dim(String(p.bytes).padStart(9))}  ${p.path}`)
          .join("\n"),
    };
  }
  if (!part) throw usageError("xml cat");
  const match = flagString(args, "match") ?? null;
  // One more than we show, to know whether there are more.
  const all = await withXmlHints(shown, part, () =>
    xml.catPart(bytes, part, match ?? undefined, xml.CAT_MATCH_LIMIT + 1),
  );
  const truncated = all.length > xml.CAT_MATCH_LIMIT;
  const found = all.slice(0, xml.CAT_MATCH_LIMIT);
  return {
    data: { file, part, match, fragments: found, truncated },
    text: () =>
      found.length
        ? found.join("\n\n") +
          (truncated
            ? `\n\n${ui.dim(`Showing the first ${xml.CAT_MATCH_LIMIT} matches.`)}`
            : "")
        : `No match in ${part}.`,
  };
}

/** What `revise examples` prints: one command per common job, ready to
 * adapt. Keep them runnable as written (given the files). */
const EXAMPLES: Array<{ description: string; command: string; docs?: string }> =
  [
    {
      description: "See a document's headings and structure",
      command: "revise outline contract.docx",
    },
    {
      description: "Rename a party everywhere, as tracked changes",
      command:
        'revise replace-all contract.docx "Acme" "Northstar" --author "Legal bot"',
    },
    {
      description:
        "Run any editing tool (revise tools lists them): here, a page-numbered footer",
      command:
        'revise call contract.docx set_header_footer \'{"side":"footer","center":"Page {PAGE} of {PAGES}"}\'',
      docs: "revise.io/developer/docs/revise/cli#editing",
    },
    {
      description: "Accept every tracked change",
      command: "revise accept contract.docx --all",
    },
    {
      description: "Redline two versions",
      command: "revise compare v1.docx v2.docx --out redline.docx",
    },
    {
      description: "Convert a Word file to PDF",
      command: "revise convert report.docx report.pdf",
    },
    {
      description: "Turn Markdown into a Word file",
      command: "cat notes.md | revise convert - notes.docx",
    },
  ];

function cmdExamples(): Outcome {
  return {
    data: { examples: EXAMPLES },
    text: () =>
      EXAMPLES.map(({ description, command, docs }) =>
        [
          ui.dim(`# ${description}`),
          ...(docs ? [ui.dim(`# Docs: ${docs}`)] : []),
          ui.cyan(command),
        ].join("\n"),
      ).join("\n\n"),
  };
}

async function cmdSkill(args: ParsedArgs, install: boolean): Promise<Outcome> {
  if (!install) {
    const name = flagString(args, "reference");
    // Own keys only: "constructor" or "__proto__" are not reference files.
    if (
      name !== undefined &&
      !Object.prototype.hasOwnProperty.call(SKILL_REFERENCES, name)
    ) {
      throw fail(
        `No reference "${name}". There are: ${Object.keys(SKILL_REFERENCES).join(", ")}.`,
        "usage",
        `revise skill --reference ${Object.keys(SKILL_REFERENCES)[0]}`,
      );
    }
    const markdown = name ? SKILL_REFERENCES[name] : SKILL_MARKDOWN;
    return {
      data: { reference: name ?? null, markdown },
      text: () => markdown.replace(/\n$/, ""),
    };
  }
  const base = flagString(args, "dir") ?? skillsDirFor(io.env ?? {});
  const dir = path.join(base, "revise-cli");
  const file = path.join(dir, "SKILL.md");
  // revise 3.0 and 3.1 installed this skill as <dir>/revise/: replace it
  // rather than leave the agent two copies. Only a folder that holds this
  // skill (its frontmatter name) is removed.
  const older = path.join(base, "revise");
  const olderSkill = path.join(older, "SKILL.md");
  const replaced =
    fs.existsSync(olderSkill) &&
    /^---\nname: revise\n/.test(fs.readFileSync(olderSkill, "utf8"));
  if (replaced) fs.rmSync(older, { recursive: true, force: true });
  fs.mkdirSync(path.join(dir, "reference"), { recursive: true });
  fs.writeFileSync(file, SKILL_MARKDOWN);
  for (const [name, markdown] of Object.entries(SKILL_REFERENCES)) {
    fs.writeFileSync(path.join(dir, "reference", `${name}.md`), markdown);
  }
  return {
    data: { output: path.resolve(file), outputWritten: true },
    text: () =>
      [
        `${ui.green("Installed")} ${dir}/ (SKILL.md and reference/)${replaced ? `, replacing the older ${older}/` : ""}.`,
        // An agent's running session doesn't load a skill installed during
        // it; new sessions do.
        `Agents: read ${file} now and follow it. New sessions load it automatically.`,
      ].join("\n"),
  };
}

/**
 * Where an agent will find the skill. Claude Code reads only .claude/skills;
 * Codex, Cursor, Copilot, and the rest read the shared .agents/skills. So:
 * Claude Code's own shell (it sets CLAUDECODE) gets .claude/skills; then a
 * folder the project already has; then .agents/skills.
 */
function skillsDirFor(env: Record<string, string | undefined>): string {
  if (env.CLAUDECODE) return ".claude/skills";
  if (fs.existsSync(".agents/skills")) return ".agents/skills";
  if (fs.existsSync(".claude")) return ".claude/skills";
  return ".agents/skills";
}

// ── --out: edit a copy ─────────────────────────────────────────────────────

/** Older command names, still accepted: scripts written for 3.0 keep
 * working, and output names the command by its current name. */
const RENAMED_COMMANDS: Record<string, string> = {
  "replace-text": "replace-all",
};

/** Commands that save their file and take --out to save a copy instead. */
const OUT_COMMANDS = new Set([
  "replace-all",
  "call",
  "accept",
  "reject",
  "xml patch",
]);

interface OutCopy {
  /** The original, by real path. */
  from: string;
  /** The original as the person typed it, for hints. */
  fromShown: string;
  /** Where --out pointed (as typed, and absolute), the copy by real path,
   * and its history's name. */
  outShown: string;
  out: string;
  file: string;
  name: string;
  /** What --force moved aside: put back if the edit fails. */
  displaced: { file?: string; history?: string };
}

/**
 * Copy the file to --out and point the command at the copy. The copy's
 * history starts from the original's latest checkpoint, so block ids, saved
 * searches, mode, and author carry over: ids read from the original work on
 * the copy. Nothing of the original is written, its history included:
 * anything to reconcile (a change made outside revise, an interrupted save)
 * is reconciled in the copy, by the edit's own load.
 */
async function copyForOut(args: ParsedArgs): Promise<OutCopy> {
  const fromShown = args.file!;
  const input = canonicalPath(requireFile(args));
  const outArg = flagString(args, "out")!;
  if (args.flags["dry-run"] === true) {
    throw fail(
      "--dry-run saves nothing, so there is no copy to write. Drop --out.",
      "usage",
    );
  }
  const ext = path.extname(input).toLowerCase();
  if (path.extname(outArg).toLowerCase() !== ext) {
    throw fail(
      `--out must be a ${ext} file like ${path.basename(input)}: the copy is saved in the same format. Write other formats with revise export.`,
      "usage",
    );
  }
  const out = path.resolve(outArg);
  if (fs.existsSync(out) && canonicalPath(out) === input) {
    throw fail(
      "--out must be a new file, not the one being edited.",
      "same_file",
    );
  }
  assertCanOverwrite(args, out);
  const sourceName = sessionNameFor(input);
  const source = sessionExists(sourceName) ? readMeta(sourceName) : null;
  const bytes = new Uint8Array(fs.readFileSync(input));
  // --force: move what is there (the file and its history) aside, to put
  // back if the edit fails outright; deleted once it succeeds.
  if (fs.lstatSync(out, { throwIfNoEntry: false })?.isDirectory()) {
    throw fail(`${outArg} is a folder; --out names a file.`, "usage");
  }
  const stamp = `${process.pid}-${Date.now()}`;
  const displaced: OutCopy["displaced"] = {};
  const name = sessionNameFor(out);
  try {
    if (fs.lstatSync(out, { throwIfNoEntry: false })) {
      displaced.file = path.join(
        path.dirname(out),
        `.${path.basename(out)}.revise-replaced-${stamp}`,
      );
      fs.renameSync(out, displaced.file);
    }
    if (sessionExists(name)) {
      displaced.history = `${documentDir(name)}.replaced-${stamp}`;
      fs.renameSync(documentDir(name), displaced.history);
    }
  } catch (error) {
    // Moving it aside failed halfway: put back what moved.
    if (displaced.file) fs.renameSync(displaced.file, out);
    throw error;
  }
  const copy: OutCopy = {
    from: input,
    fromShown,
    outShown: outArg,
    out,
    file: out,
    name,
    displaced,
  };
  try {
    writeOutput(args, out, bytes);
    copy.file = canonicalPath(out);
    if (source) {
      const { checkpoints: _history, lastTx: _tx, ...kept } = source;
      const meta = createDocumentRecord(
        {
          ...kept,
          name,
          file: copy.file,
          // The original's own hash: if the file changed since its last
          // checkpoint, the edit's load re-learns it, in the copy's history.
          fileHash: source.fileHash,
          documentId: `cli-${name}`,
          source: input,
          createdAt: new Date().toISOString(),
        },
        readState(source),
        readPackage(source),
      );
      meta.checkpoints[0] = {
        ...meta.checkpoints[0],
        command: "copy",
        summary: `Copied from ${path.basename(input)}`,
      };
      writeMeta(meta);
    }
  } catch (error) {
    discardCopy(copy);
    throw error;
  }
  args.file = outArg;
  return copy;
}

/** The edit failed outright: remove the copy, put back what it replaced. */
function discardCopy(copy: OutCopy): void {
  fs.rmSync(copy.file, { force: true });
  removeDocumentRecord(copy.name);
  if (copy.displaced.file) {
    fs.renameSync(copy.displaced.file, copy.out);
  }
  if (copy.displaced.history) {
    fs.renameSync(copy.displaced.history, documentDir(copy.name));
  }
}

/** A failed --out edit left no copy: its hints name the original. */
function hintForOriginal(hint: string | null, copy: OutCopy): string | null {
  return hint
    ? hint.split(shellArg(copy.outShown)).join(shellArg(copy.fromShown))
    : hint;
}

function pointHintAtOriginal(error: CliError, copy: OutCopy): void {
  const info = extras.get(error);
  if (info?.hint) info.hint = hintForOriginal(info.hint, copy);
}

/** The edit succeeded: what --force replaced can go. */
function keepCopy(copy: OutCopy): void {
  if (copy.displaced.file) fs.rmSync(copy.displaced.file, { force: true });
  if (copy.displaced.history) {
    fs.rmSync(copy.displaced.history, { recursive: true, force: true });
  }
}

// ── help ────────────────────────────────────────────────────────────────────

interface CommandHelp {
  name: string;
  group: string;
  usage: string;
  summary: string;
  options?: Array<[string, string]>;
}

const FORCE: [string, string] = [
  "--force",
  "Replace the output file if it exists",
];

const SECTION_OPTION: [string, string] = [
  "--section <s>",
  'A heading and its subsections: "8", "8.1", or its title',
];
const BLOCKS_OPTION: [string, string] = [
  "--blocks <b>",
  "Block ids and ranges: id,id or first..last",
];
const VIEW_OPTION: [string, string] = [
  "--view <view>",
  "tracked (default): pending changes marked; accepted; or original",
];
const OUT_OPTION: [string, string] = [
  "--out <file>",
  "Save the result as a new file (same format) and leave this one untouched; --force replaces an existing one",
];

const COMMAND_HELP: CommandHelp[] = [
  {
    name: "status",
    group: "Reading",
    usage: "revise status <file>",
    summary: "Size, mode, tracked changes, checkpoint, Word fidelity",
  },
  {
    name: "outline",
    group: "Reading",
    usage: "revise outline <file> [--limit <n>] [--offset <n>]",
    summary: "Headings with block ranges",
    options: [
      ["--limit <n>", "Sections to show"],
      ["--offset <n>", "Section to start from (0-based)"],
    ],
  },
  {
    name: "read",
    group: "Reading",
    usage: "revise read <file> [<index>] [--blocks <id,id>] [--view <view>]",
    summary: "Blocks as HTML with block ids; pending changes as <ins>/<del>",
    options: [
      ["<index>", "Block index to start from (0-based, default 0)"],
      ["--blocks <id,id>", "Read these blocks instead"],
      VIEW_OPTION,
    ],
  },
  {
    name: "grep",
    group: "Reading",
    usage:
      'revise grep <file> "<text>" [--regex] [--ignore-case] [--view <view>]',
    summary: "Matching blocks as plain text with ids",
    options: [
      ["--regex", "Treat the text as a JavaScript regular expression"],
      ["--ignore-case", "Match regardless of case"],
      [
        "--view <view>",
        "tracked (default): matches deleted and inserted text, [-marked-]{+inline+}; accepted; or original",
      ],
    ],
  },
  {
    name: "search",
    group: "Reading",
    usage:
      'revise search <file> "<text>" ["<text>" ...] [--results-page <n>] [--view <view>]',
    summary: "Matching blocks as HTML, with a search result id",
    options: [
      ["--results-page <n>", "Page of results (1-based)"],
      [
        "--view <view>",
        "tracked (default): also lists matches in deleted text; accepted; or original (no result id)",
      ],
    ],
  },
  {
    name: "replace-all",
    group: "Editing (saves the file)",
    usage:
      'revise replace-all <file> "<find>" "<replace>" [--ignore-case] [--exact] [--blocks <id,id>]',
    summary: "Replace every occurrence: body, headers, footers, notes",
    options: [
      ["--ignore-case", "Also PROVIDER and provider"],
      [
        "--exact",
        'Quotes and spaces must match too (by default “ and ", no-break and plain spaces match)',
      ],
      ["--blocks <id,id>", "Only in these blocks"],
      OUT_OPTION,
    ],
  },
  {
    name: "call",
    group: "Editing (saves the file)",
    usage: "revise call <file> <tool> ['<json>' | - | @<file.json>] [--direct]",
    summary: "Run an editing tool (or: revise <tool> <file> '<json>')",
    options: [
      ["'<json>'", "Tool input; - reads stdin, @file reads a file"],
      ["--direct", "Apply this call directly, even in suggest mode"],
      OUT_OPTION,
    ],
  },
  {
    name: "tools",
    group: "Editing (saves the file)",
    usage: "revise tools [<tool>]",
    summary: "List editing tools, or show one's input schema",
  },
  {
    name: "mode",
    group: "Editing (saves the file)",
    usage: "revise mode <file> [suggest | edit]",
    summary: "Show or set: suggest (tracked changes) or edit (direct)",
  },
  {
    name: "changes",
    group: "Review",
    usage: "revise changes <file> [--section <s>] [--blocks <b>] [--details]",
    summary: "Tracked changes: where, and what text",
    options: [
      SECTION_OPTION,
      BLOCKS_OPTION,
      ["--details", "List every change, even when there are many"],
    ],
  },
  {
    name: "diff",
    group: "Review",
    usage: "revise diff <file> [--since <checkpoint>]",
    summary: "Word-level diff of the tracked changes",
    options: [
      [
        "--since <checkpoint>",
        "Diff from a checkpoint instead (changes accepted)",
      ],
    ],
  },
  {
    name: "accept",
    group: "Review",
    usage:
      "revise accept <file> (<change-id> ... | --all | --section <s> | --except-section <s> | --blocks <b>) [--dry-run]",
    summary: "Accept tracked changes",
    options: [
      ["--all", "Every tracked change"],
      SECTION_OPTION,
      ["--except-section <s>", "Every change outside one section"],
      BLOCKS_OPTION,
      ["--dry-run", "Show what would be accepted; change nothing"],
      OUT_OPTION,
    ],
  },
  {
    name: "reject",
    group: "Review",
    usage:
      "revise reject <file> (<change-id> ... | --all | --section <s> | --except-section <s> | --blocks <b>) [--dry-run]",
    summary: "Reject tracked changes",
    options: [
      ["--all", "Every tracked change"],
      SECTION_OPTION,
      ["--except-section <s>", "Every change outside one section"],
      BLOCKS_OPTION,
      ["--dry-run", "Show what would be rejected; change nothing"],
      OUT_OPTION,
    ],
  },
  {
    name: "history",
    group: "Review",
    usage: "revise history <file>",
    summary: "Checkpoints: one per saved change",
  },
  {
    name: "undo",
    group: "Review",
    usage: "revise undo <file> [<steps>]",
    summary: "Go back some checkpoints (default 1); later ones are discarded",
  },
  {
    name: "restore",
    group: "Review",
    usage: "revise restore <file> <checkpoint>",
    summary: "Go back to a checkpoint; later ones are discarded",
  },
  {
    name: "render",
    group: "Output",
    usage:
      "revise render <file> [--page <n> | --pages <n-m>] [--view <view>] [--out <file>]",
    summary: "Pages as an image",
    options: [
      ["--page <n>", "Page to render (1-based, default 1)"],
      ["--pages <n-m>", "Up to 6 pages in one image"],
      VIEW_OPTION,
      ["--layout", "Outline layout boxes"],
      ["--detailed", "Full resolution for several pages"],
      ["--out <file>", "Where to save it (default: inside .revise/)"],
    ],
  },
  {
    name: "export",
    group: "Output",
    usage: "revise export <file> <output> [--view <view>] [--force]",
    summary: "Write a copy as .docx, .pdf, .md, .html, or .txt",
    options: [
      [
        "--view <view>",
        "tracked (default; only .docx keeps changes), accepted, or original",
      ],
      FORCE,
    ],
  },
  {
    name: "new",
    group: "Create and convert",
    usage: "revise new <output> [--content <draft.html|.md>] [options]",
    summary: "Create a document, in edit mode",
    options: [
      ["--content <file>", "Start from HTML or Markdown"],
      ["--title <text>", "Document title"],
      ["--mode suggest", "Start in suggest mode instead"],
      ["--page-size <size>", "letter, a4, legal, ..."],
      ["--orientation <o>", "portrait or landscape"],
      ["--margins <length>", "For example 1in or 2.5cm"],
      ["--header <text>", "Header text (also --header-left/-center/-right)"],
      ["--footer <text>", 'Footer text, e.g. "Page {PAGE} of {PAGES}"'],
      FORCE,
    ],
  },
  {
    name: "compare",
    group: "Create and convert",
    usage:
      "revise compare <original> <revised> [--out <redline.docx>] [--force]",
    summary: "A redline: the revised file with every difference tracked",
    options: [
      ["--out <file>", "Where to write it (default: <revised>-redline.docx)"],
      ["--author <name>", "Sign the tracked changes"],
      FORCE,
    ],
  },
  {
    name: "convert",
    group: "Create and convert",
    usage: "revise convert <input|-> <output|-> [--view <view>] [--force]",
    summary:
      "Convert .docx .md .html .odt .rtf .txt to .docx .md .html .txt .pdf",
    options: [
      [
        "--view <view>",
        "tracked (default; only .docx keeps changes), accepted, or original",
      ],
      FORCE,
      [
        "--from <format>",
        "With - as input: the format on stdin (docx, odt, rtf, html, md, txt); recognized by default, and plain text reads as Markdown",
      ],
      [
        "--to <format>",
        "With - as output: the format to write to stdout (md, html, txt, docx, pdf; default md)",
      ],
    ],
  },
  {
    name: "xml ls",
    group: "Raw package (what tools don't cover)",
    usage: "revise xml ls <file>",
    summary: "Parts in the .docx package",
  },
  {
    name: "xml cat",
    group: "Raw package (what tools don't cover)",
    usage: "revise xml cat <file> <part> [--match <text>]",
    summary: "A part's XML, or the elements containing some text",
  },
  {
    name: "xml patch",
    group: "Raw package (what tools don't cover)",
    usage:
      "revise xml patch <file> <part> (--find '<xml>' --replace '<xml>' [--all] | --file <new.xml>)",
    summary: "Edit a part directly (not tracked)",
    options: [OUT_OPTION],
  },
  {
    name: "skill",
    group: "Agents",
    usage: "revise skill [--reference <name>]",
    summary: "Print the agent skill (SKILL.md)",
    options: [
      [
        "--reference <name>",
        "Print one of its reference files instead: library, json",
      ],
    ],
  },
  {
    name: "skill install",
    group: "Agents",
    usage: "revise skill install [--dir <dir>]",
    summary: "Install the agent skill as <dir>/revise-cli/SKILL.md",
    options: [
      [
        "--dir <dir>",
        "Skills folder (default .agents/skills or .claude/skills)",
      ],
    ],
  },
  {
    name: "examples",
    group: "Other",
    usage: "revise examples",
    summary: "Example commands for the most common jobs",
  },
  {
    name: "help",
    group: "Other",
    usage: "revise help [<command>]",
    summary: "Help for revise, or for one command",
  },
  {
    name: "version",
    group: "Other",
    usage: "revise version",
    summary: "The version and the commit it was built from",
  },
];

const HELP_BY_NAME = new Map(COMMAND_HELP.map((c) => [c.name, c]));

const GLOBAL_OPTIONS: Array<[string, string]> = [
  [
    "--json",
    "Machine-readable output (see revise.io/developer/docs/revise/json)",
  ],
  ["--author <name>", "Sign tracked changes (or REVISE_AUTHOR)"],
  ["--no-color", "Plain text (or NO_COLOR=1)"],
  ["--help", "Help for a command"],
];

function usageError(name: string): CliError {
  const help = HELP_BY_NAME.get(name)!;
  return fail(`Usage: ${help.usage}`, "usage", `revise ${name} --help`);
}

function optionLines(options: Array<[string, string]>): string[] {
  const width = Math.max(...options.map(([flag]) => flag.length)) + 2;
  return options.map(
    ([flag, text]) => `  ${ui.cyan(flag.padEnd(width))}${text}`,
  );
}

function globalHelp(): string {
  const lines = [
    `${ui.bold("revise")} — the swiss army knife for documents`,
    "",
    `${ui.dim("Docs:")}        revise.io/developer/docs/cli`,
    `${ui.dim("Examples:")}    ${ui.cyan("revise examples")}`,
    `${ui.dim("Agent Skill:")} ${ui.cyan("revise skill")}`,
  ];
  let group = "";
  for (const c of COMMAND_HELP) {
    if (c.group !== group) {
      group = c.group;
      lines.push("", ui.bold(group));
    }
    lines.push(`  ${ui.cyan(c.name.padEnd(13))}${c.summary}`);
  }
  lines.push(
    "",
    ui.bold("Options"),
    ...optionLines(GLOBAL_OPTIONS),
    "",
    `${ui.dim("Details:")} ${ui.cyan("revise <command> --help")}`,
  );
  return lines.join("\n");
}

function commandHelp(help: CommandHelp): string {
  return [
    `${ui.dim("Usage:")} ${help.usage}`,
    "",
    `${help.summary}.`,
    ...(help.options?.length
      ? ["", ui.bold("Options"), ...optionLines(help.options)]
      : []),
    "",
    ui.bold("Global options"),
    ...optionLines(GLOBAL_OPTIONS),
  ].join("\n");
}

// ── dispatch ────────────────────────────────────────────────────────────────

/** Commands whose first argument is the document they work on. */
const FILE_COMMANDS = new Set([
  "status",
  "outline",
  "read",
  "grep",
  "search",
  "replace-all",
  "call",
  "mode",
  "changes",
  "diff",
  "accept",
  "reject",
  "history",
  "undo",
  "restore",
  "render",
  "export",
]);

const COMMANDS = new Set([
  ...COMMAND_HELP.map((c) => c.name.split(" ")[0]),
  "help",
  "version",
]);

/** Whether `name` is an editing tool (for the `revise <tool> …` shorthand). */
async function isToolName(name: string): Promise<boolean> {
  return (await toolDefinition(name)) !== undefined;
}

async function toolDefinition(
  name: string,
): Promise<
  { name: string; description: string; inputSchema: unknown } | undefined
> {
  const probe = await ReviseDocument.create();
  try {
    return probe.tools.definitions().find((tool) => tool.name === name);
  } finally {
    probe.dispose();
  }
}

interface JsonSchemaNode {
  type?: string | string[];
  const?: unknown;
  enum?: unknown[];
  items?: JsonSchemaNode;
  properties?: Record<string, JsonSchemaNode>;
  required?: string[];
  description?: string;
  oneOf?: JsonSchemaNode[];
  anyOf?: JsonSchemaNode[];
}

function schemaType(node: JsonSchemaNode): string {
  if (node.const !== undefined) return JSON.stringify(node.const);
  if (node.enum) return node.enum.map((v) => JSON.stringify(v)).join(" | ");
  // An object variant shows its shape: { start_text, end_text }.
  if (
    node.properties &&
    (node.oneOf ?? node.anyOf) === undefined &&
    node.type === "object"
  ) {
    return `{ ${Object.entries(node.properties)
      .map(([key, value]) =>
        value.const !== undefined
          ? `${key}: ${JSON.stringify(value.const)}`
          : key,
      )
      .join(", ")} }`;
  }
  const variants = node.oneOf ?? node.anyOf;
  if (variants) return variants.map(schemaType).join(" | ");
  if (node.type === "array" && node.items) {
    const inner = schemaType(node.items);
    return inner.includes(" ") ? `(${inner})[]` : `${inner}[]`;
  }
  if (Array.isArray(node.type)) return node.type.join(" | ");
  return node.type ?? "any";
}

/** A tool's input as one line per field (and one level of nested fields),
 * required first: what an agent needs to write the JSON, in fewer tokens
 * than the schema. */
function inputFields(schema: unknown, indent = "  "): string[] {
  const node = schema as JsonSchemaNode;
  const required = new Set(node.required ?? []);
  const names = Object.keys(node.properties ?? {}).sort(
    (a, b) => Number(required.has(b)) - Number(required.has(a)),
  );
  return names.flatMap((name) => {
    const field = node.properties![name];
    const about = field.description?.split(/(?<=\.)\s/)[0] ?? "";
    const line = `${indent}${ui.cyan(name)} (${schemaType(field)}${required.has(name) ? ", required" : ""})${about ? `: ${about}` : ""}`;
    const nested = field.type === "array" ? field.items : field;
    // A choice of object shapes: each shape's fields, under the field.
    const variants = (nested?.oneOf ?? nested?.anyOf)?.filter(
      (v) => v.properties,
    );
    if (variants?.length) {
      return [
        line,
        ...variants.flatMap((v) => inputFields(v, `${indent}    `)),
      ];
    }
    return nested?.properties
      ? [line, ...inputFields(nested, `${indent}    `)]
      : [line];
  });
}

/** Help for one tool: usage, what it does, and its input fields. */
function toolHelpText(
  tool: { name: string; description: string; inputSchema: unknown },
  usage: string,
): string {
  const fields = inputFields(tool.inputSchema);
  return [
    `${ui.dim("Usage:")} ${usage}`,
    "",
    tool.description,
    "",
    ui.bold("Input (a JSON object)"),
    ...(fields.length ? fields : ["  (none)"]),
    "",
    `${ui.dim("Full JSON schema:")} ${ui.cyan(`revise tools ${tool.name} --json`)}`,
  ].join("\n");
}

/**
 * `revise <tool> <file> '<json>'` gets help of its own, so --help, usage
 * errors, and their hints are about the tool rather than about `call`.
 */
function shorthandHelp(tool: {
  name: string;
  description: string;
}): CommandHelp {
  const call = HELP_BY_NAME.get("call")!;
  return {
    name: tool.name,
    group: call.group,
    usage: `revise ${tool.name} <file> ['<json>' | - | @<file.json>] [--direct] [--out <file>]`,
    summary: tool.description.split(/(?<=\.)\s/)[0].replace(/\.$/, ""),
    options: call.options,
  };
}

/** Every file a command writes or edits, for locking: taken together, in
 * one order, before the command decides anything about them. */
function lockTargetsFor(command: string, args: ParsedArgs): string[] {
  const targets: string[] = [];
  if (args.file) targets.push(args.file);
  if (command === "export" && args.positional[0]) {
    targets.push(args.positional[0]);
  }
  if (
    command === "convert" &&
    args.positional[1] &&
    args.positional[1] !== "-"
  ) {
    targets.push(args.positional[1]);
  }
  if (command === "new" && args.positional[0]) {
    targets.push(args.positional[0]);
  }
  if (command === "compare" && args.positional[1]) {
    targets.push(
      flagString(args, "out") ??
        `${args.positional[1].replace(/\.[^.]+$/, "")}-redline.docx`,
    );
  }
  return targets.map((file) => path.resolve(file));
}

/** Errors that mean the outcome is unknown, not merely refused. */
const UNCERTAIN = new Set(["save_uncertain"]);

/**
 * PDF export warns through console.warn, once per missing font. Collect
 * those into one warning instead of letting them reach the terminal.
 */
async function capturingEngineWarnings<T>(work: () => Promise<T>): Promise<T> {
  const original = console.warn;
  const fonts: string[] = [];
  console.warn = (...parts: unknown[]) => {
    const text = parts.map(String).join(" ");
    const font = /^Failed to load font (\S+);/.exec(text);
    if (font) fonts.push(font[1]);
    else warn("engine_warning", text);
  };
  try {
    return await work();
  } finally {
    console.warn = original;
    if (fonts.length) {
      warn(
        "font_fallback",
        `The PDF uses fallback fonts for ${plural(fonts.length, "font")} that could not be loaded (${fonts.join(", ")}); layout may differ.`,
        { fonts },
      );
    }
  }
}

export async function main(argv: string[], cliIo?: CliIo): Promise<number> {
  io = cliIo ? { ...cliIo } : defaultIo();
  warnings = [];
  currentCommand = null;
  jsonOutput = argv.includes("--json");
  const env = io.env ?? {};
  // NO_COLOR wins (see style.ts). Node itself warns on stderr when both are
  // set, the first time something it loads asks about color; settle it
  // before anything does.
  if (process.env.NO_COLOR && process.env.FORCE_COLOR !== undefined) {
    delete process.env.FORCE_COLOR;
  }
  const noColor = argv.includes("--no-color");
  ui = styleFor({
    isTTY: Boolean(io.stdoutIsTTY),
    env,
    noColor,
    json: jsonOutput,
  });
  uiErr = styleFor({
    isTTY: Boolean(io.stderrIsTTY),
    env,
    noColor,
    json: jsonOutput,
  });
  let command: string | null = null;
  const finish = (
    data: Record<string, unknown> | null,
    error: CliErrorInfo | null,
    code: number,
    text?: string,
  ): number => {
    if (jsonOutput) printEnvelope(command, data, error);
    else {
      // A command that failed outright says so on stderr only.
      if (text && code !== 1) out(text);
      if (error) printError(error);
    }
    return code;
  };
  try {
    const args = parseArgs(argv);
    const typed = args.positional.shift();
    const first =
      typed && RENAMED_COMMANDS[typed] ? RENAMED_COMMANDS[typed] : typed;
    if (first === "help" && args.positional[0]) {
      args.positional[0] =
        RENAMED_COMMANDS[args.positional[0]] ?? args.positional[0];
    }
    if ((args.flags.version || first === "version") && !args.flags.help) {
      command = "version";
      const { version, build } = producer();
      return finish({ version, build }, null, 0, version);
    }
    if (!first || first === "help") {
      command = "help";
      const topic = first === "help" ? args.positional.join(" ") : "";
      const help = topic ? HELP_BY_NAME.get(topic) : undefined;
      if (topic && !help) {
        throw fail(`No command "${topic}".`, "unknown_command", "revise help");
      }
      return finish(
        helpData(help),
        null,
        0,
        help ? commandHelp(help) : globalHelp(),
      );
    }
    // Tool shorthand, resolved before anything locks: `revise <tool> <file>
    // '<json>'` is `revise call <file> <tool> '<json>'`.
    if (first === "open") {
      throw fail(
        "There's no open step: name the file in each command.",
        "usage",
        "revise read <file>",
      );
    }
    let resolved = first;
    let shorthandTool: string | null = null;
    if (!COMMANDS.has(first)) {
      if (!(await isToolName(first))) {
        throw fail(
          `Unknown command "${first}".`,
          "unknown_command",
          "revise help",
        );
      }
      resolved = "call";
      shorthandTool = first;
      args.positional = [
        args.positional[0],
        first,
        ...args.positional.slice(1),
      ];
    }
    if (resolved === "xml") {
      const sub = args.positional.shift();
      if (!sub || !["ls", "cat", "patch"].includes(sub)) {
        throw fail(
          "Usage: revise xml (ls | cat | patch) <file> ...",
          "usage",
          "revise xml ls --help",
        );
      }
      resolved = `xml ${sub}`;
    }
    if (resolved === "skill" && args.positional[0] === "install") {
      args.positional.shift();
      resolved = "skill install";
    }
    if (args.flags.version) resolved = "version";
    // Run under an older name, the output keeps that name (3.0 scripts).
    command = typed && RENAMED_COMMANDS[typed] ? typed : resolved;
    currentCommand = resolved;
    if (shorthandTool) {
      const definition = (await toolDefinition(shorthandTool))!;
      HELP_BY_NAME.set(shorthandTool, shorthandHelp(definition));
      currentCommand = shorthandTool;
      // `revise <tool>` alone, or with --help: what the tool takes.
      if (args.flags.help || args.positional[0] === undefined) {
        command = "help";
        const help = HELP_BY_NAME.get(shorthandTool)!;
        return finish(
          { ...helpData(help), inputSchema: definition.inputSchema },
          null,
          0,
          `${toolHelpText(definition, help.usage)}\n\n${ui.bold("Global options")}\n${optionLines(GLOBAL_OPTIONS).join("\n")}`,
        );
      }
    }
    if (args.flags.help) {
      command = "help";
      const help = HELP_BY_NAME.get(resolved);
      return finish(
        helpData(help),
        null,
        0,
        help ? commandHelp(help) : globalHelp(),
      );
    }
    if (FILE_COMMANDS.has(resolved) || resolved.startsWith("xml ")) {
      args.file = args.positional.shift();
    }
    const dispatch = resolved;
    const copying =
      OUT_COMMANDS.has(dispatch) && flagString(args, "out") !== undefined;
    let copy: OutCopy | null = null;
    const run = async (): Promise<Outcome> => {
      if (copying) copy = await copyForOut(args);
      switch (dispatch) {
        case "new":
          return cmdNew(args);
        case "status":
          return cmdStatus(args);
        case "read":
          return cmdRead(args);
        case "search":
          return cmdSearch(args);
        case "outline":
          return cmdOutline(args);
        case "tools":
          return cmdTools(args);
        case "call":
          return cmdCall(args);
        case "mode":
          return cmdMode(args);
        case "changes":
          return cmdChanges(args);
        case "accept":
          return cmdSettle(args, true);
        case "reject":
          return cmdSettle(args, false);
        case "diff":
          return cmdDiff(args);
        case "history":
          return cmdHistory(args);
        case "undo":
          return cmdUndo(args);
        case "restore": {
          if (!args.positional[0]) throw usageError("restore");
          return cmdUndo(
            args,
            strictInt(args.positional[0], "The checkpoint number"),
          );
        }
        case "render":
          return cmdRender(args);
        case "export":
          return cmdExport(args);
        case "convert":
          return cmdConvert(args);
        case "compare":
          return cmdCompare(args);
        case "grep":
          return cmdGrep(args);
        case "replace-all":
          return cmdReplaceAll(args);
        case "xml ls":
          return cmdXml(args, "ls");
        case "xml cat":
          return cmdXml(args, "cat");
        case "xml patch":
          return cmdXml(args, "patch");
        case "examples":
          return cmdExamples();
        case "skill":
          return cmdSkill(args, false);
        case "skill install":
          return cmdSkill(args, true);
        default:
          throw fail(
            `Unknown command "${dispatch}".`,
            "unknown_command",
            "revise help",
          );
      }
    };
    const targets = lockTargetsFor(dispatch.split(" ")[0], args);
    if (copying) targets.push(path.resolve(flagString(args, "out")!));
    // Every --out command's data says where the copy came from (or null).
    const stamped = (data: Record<string, unknown> | null) =>
      data && OUT_COMMANDS.has(dispatch)
        ? { ...data, copiedFrom: copy?.from ?? null }
        : data;
    let outcome: Outcome;
    try {
      outcome = await capturingEngineWarnings(() =>
        targets.length ? withFileLocks(targets, run) : run(),
      );
    } catch (error) {
      // Failed outright: nothing was saved, so neither was a copy.
      const receipt =
        error instanceof CliError ? extras.get(error)?.receipt : undefined;
      const failedOutright =
        !receipt &&
        (error instanceof CliError
          ? !UNCERTAIN.has(error.code)
          : systemProblem(error) !== null);
      if (copy && failedOutright) {
        discardCopy(copy);
        if (error instanceof CliError) pointHintAtOriginal(error, copy);
      } else {
        if (copy) keepCopy(copy);
        if (receipt) receipt.data = stamped(receipt.data)!;
      }
      throw error;
    }
    const code = outcome.error ? (outcome.exit ?? 2) : 0;
    if (copy && code === 1) {
      discardCopy(copy);
      if (outcome.error) {
        outcome.error = {
          ...outcome.error,
          hint: hintForOriginal(outcome.error.hint, copy),
        };
      }
    } else if (copy) keepCopy(copy);
    const copied = copy as { from: string } | null;
    const text = outcome.text();
    return finish(
      stamped(outcome.data),
      outcome.error ?? null,
      code,
      copied && text
        ? `${ui.dim(`Copied ${path.basename(copied.from)} to ${shellArg(args.file!)}.`)}\n${text}`
        : text,
    );
  } catch (error) {
    if (error instanceof CliError) {
      const receipt = extras.get(error)?.receipt;
      if (receipt) return finish(receipt.data, problemOf(error), receipt.exit);
      if (UNCERTAIN.has(error.code)) {
        return finish(null, problemOf(error), 70);
      }
      return finish(null, problemOf(error), 1);
    }
    const system = systemProblem(error);
    if (system) return finish(null, system, 1);
    // Anything else is a bug or an environment failure: still one clear
    // line (and valid JSON), never a bare stack trace.
    const message = (error as Error)?.message ?? String(error);
    const code = finish(
      null,
      {
        code: "internal",
        message,
        hint: null,
        details: null,
      },
      70,
    );
    if (process.env.REVISE_DEBUG) io.stderr(`${(error as Error)?.stack}\n`);
    return code;
  }
}

/** A file the command needed is missing, or the system refused a read or
 * write: an expected failure, not a bug. Null for anything else. */
function systemProblem(error: unknown): CliErrorInfo | null {
  const system = error as NodeJS.ErrnoException;
  if (system?.code === "ENOENT" && system.path) {
    return {
      code: "file_not_found",
      message: `No such file: ${system.path}`,
      hint: null,
      details: null,
    };
  }
  if (typeof system?.code === "string" && system.syscall) {
    return {
      code: "io_error",
      message: `${system.syscall}${system.path ? ` ${system.path}` : ""} failed: ${system.code}.`,
      hint: null,
      details: { errno: system.code, syscall: system.syscall },
    };
  }
  return null;
}

function helpData(help: CommandHelp | undefined): Record<string, unknown> {
  const describe = (c: CommandHelp) => ({
    name: c.name,
    usage: c.usage,
    summary: c.summary,
    options: (c.options ?? []).map(([flag, description]) => ({
      flag,
      description,
    })),
  });
  return help
    ? { command: describe(help) }
    : {
        commands: COMMAND_HELP.map(describe),
        globalOptions: GLOBAL_OPTIONS.map(([flag, description]) => ({
          flag,
          description,
        })),
      };
}
