/**
 * On-disk sessions for the CLI.
 *
 * Every CLI invocation is a fresh process, so a document's history lives in
 * a `.revise/` folder next to the file (REVISE_WORKSPACE overrides it), keyed
 * by the file's real path, so it follows the file whatever directory the
 * command runs from:
 *
 *   .revise/
 *     docs/<name>/
 *       meta.json             file, mode, checkpoint log (the commit point)
 *       checkpoints/<n>.bin   document state after checkpoint n (0 = opened)
 *       checkpoints/<n>.docx  the Word package as of checkpoint n, when it changed
 *       renders/              page images written by `revise render`
 *       pending.json          a save in progress (see "Saving" below)
 *     locks/<name>.lock       held while a command works on the file
 *
 * Every file is written atomically (temp file, then rename), and meta.json is
 * written last: a checkpoint exists only once the log names it. The current
 * state is always the latest checkpoint.
 *
 * Saving changes two things, the document file and its history, so it is
 * journaled: stage the checkpoint and write pending.json (what the file will
 * hash to), write the file, then commit the log and remove pending.json. A
 * command interrupted in between leaves pending.json behind, and the next
 * command settles it by its transaction id: already in the log (just clean
 * up), file written (finish), or neither (discard). A checkpoint the log
 * names is never deleted, and cleanup after the commit can't fail a save.
 */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { createHash, randomBytes } from "node:crypto";

export type CliDocumentMode = "suggesting" | "editing";

export interface CheckpointEntry {
  n: number;
  at: string;
  command: string;
  summary: string;
  /** An xml patch: the package (the file) is the truth here, and the model
   * is re-learned from it before use. */
  modelStale?: boolean;
}

export interface CliDocumentMeta {
  name: string;
  /** Absolute path of the file this session caches. */
  file: string;
  /** Hash of the file as revise last read or wrote it: a different hash
   * means it changed outside revise and is re-learned. */
  fileHash: string;
  documentId: string;
  source: string | null;
  sourceFormat: string;
  createdAt: string;
  mode: CliDocumentMode;
  /** Author stamped on tracked changes and comments (DOCX w:author). */
  author?: string;
  /** Saved searches, replayed so search_result_id works across processes. */
  searches?: Record<string, Record<string, unknown>>;
  searchCounter?: number;
  /** Block ids the agent has been shown by read/search output. */
  readBlockIds?: string[];
  /** The latest checkpoint's model is stale (an xml patch): the next command
   * that loads the model re-learns it from the file first. */
  packageDirty?: boolean;
  /** The last committed save's transaction id (see "Saving"). */
  lastTx?: string;
  checkpoints: CheckpointEntry[];
}

/** The CLI's errors are the library's: one type, one set of codes. */
import { ReviseError as CliError } from "../lib/errors";
export { CliError };

const rootByName = new Map<string, string>();

/** The file's real path: symlinks and `..` resolved, so one file is one
 * session however it is named. */
export function canonicalPath(file: string): string {
  const absolute = path.resolve(file);
  try {
    return fs.realpathSync(absolute);
  } catch {
    try {
      return path.join(
        fs.realpathSync(path.dirname(absolute)),
        path.basename(absolute),
      );
    } catch {
      return absolute;
    }
  }
}

function rootFor(file: string): string {
  return process.env.REVISE_WORKSPACE
    ? path.resolve(process.env.REVISE_WORKSPACE)
    : path.join(path.dirname(file), ".revise");
}

export function documentDir(name: string): string {
  const root = rootByName.get(name);
  if (!root) throw new Error(`Unknown session ${name}`);
  return path.join(root, "docs", name);
}

/** The cache name for a file: readable, unique per real path. */
export function sessionNameFor(file: string): string {
  const real = canonicalPath(file);
  const base =
    path
      .basename(real)
      .toLowerCase()
      .replace(/[^a-z0-9._-]+/g, "-")
      .slice(0, 40) || "document";
  const digest = createHash("sha1").update(real).digest("hex");
  const name = `${base}-${digest.slice(0, 8)}`;
  rootByName.set(name, rootFor(real));
  return name;
}

export function sessionExists(name: string): boolean {
  return fs.existsSync(path.join(documentDir(name), "meta.json"));
}

export function hashFileBytes(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

/** Write a file so readers see the old bytes or the new ones, never a
 * torn mix: write a sibling temp file, then rename over the target. */
export function writeFileAtomic(file: string, data: Uint8Array | string): void {
  const temp = path.join(
    path.dirname(file),
    `.${path.basename(file)}.${process.pid}.${randomBytes(4).toString("hex")}.tmp`,
  );
  // A file someone made read-only stays untouched: the rename below would
  // otherwise replace it, since only its directory needs to be writable.
  let mode: number | undefined;
  if (fs.existsSync(file)) {
    fs.accessSync(file, fs.constants.W_OK);
    mode = fs.statSync(file).mode & 0o7777;
  }
  try {
    fs.writeFileSync(temp, data);
    if (mode !== undefined) fs.chmodSync(temp, mode);
    fs.renameSync(temp, file);
  } catch (error) {
    fs.rmSync(temp, { force: true });
    throw error;
  }
}

export function readMeta(name: string): CliDocumentMeta {
  return JSON.parse(
    fs.readFileSync(path.join(documentDir(name), "meta.json"), "utf8"),
  ) as CliDocumentMeta;
}

export function writeMeta(meta: CliDocumentMeta): void {
  writeFileAtomic(
    path.join(documentDir(meta.name), "meta.json"),
    `${JSON.stringify(meta, null, 2)}\n`,
  );
}

function checkpointPath(name: string, n: number, ext = "bin"): string {
  return path.join(
    documentDir(name),
    "checkpoints",
    `${String(n).padStart(4, "0")}.${ext}`,
  );
}

export function readCheckpoint(name: string, n: number): Uint8Array {
  const file = checkpointPath(name, n);
  if (!fs.existsSync(file)) {
    throw new CliError(
      `Checkpoint ${n} does not exist.`,
      "checkpoint_not_found",
    );
  }
  return new Uint8Array(fs.readFileSync(file));
}

/** The current document state: the latest checkpoint's. */
export function readState(meta: CliDocumentMeta): Uint8Array {
  return readCheckpoint(meta.name, meta.checkpoints.at(-1)?.n ?? 0);
}

/** The Word package as of checkpoint n: the latest one saved at or before
 * it, or null for a document that never was a .docx. */
export function packageAt(name: string, n: number): Uint8Array | null {
  for (let k = n; k >= 0; k -= 1) {
    const file = checkpointPath(name, k, "docx");
    if (fs.existsSync(file)) return new Uint8Array(fs.readFileSync(file));
  }
  return null;
}

export function readPackage(meta: CliDocumentMeta): Uint8Array | null {
  return packageAt(meta.name, meta.checkpoints.at(-1)?.n ?? 0);
}

export function createDocumentRecord(
  meta: Omit<CliDocumentMeta, "checkpoints">,
  state: Uint8Array,
  pkg: Uint8Array | null = null,
): CliDocumentMeta {
  const dir = documentDir(meta.name);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(path.join(dir, "checkpoints"), { recursive: true });
  writeFileAtomic(checkpointPath(meta.name, 0), state);
  if (pkg) writeFileAtomic(checkpointPath(meta.name, 0, "docx"), pkg);
  const full: CliDocumentMeta = {
    ...meta,
    checkpoints: [
      {
        n: 0,
        at: new Date().toISOString(),
        command: "open",
        summary: meta.source ? `Opened ${meta.source}` : "Created",
      },
    ],
  };
  writeMeta(full);
  return full;
}

/**
 * Record the next checkpoint: its state (and package, if it changed), then
 * the log. Call only after the file itself is safely written, with that
 * file's hash: history must never describe a file that was not saved.
 */
export function commitState(
  meta: CliDocumentMeta,
  state: Uint8Array,
  command: string,
  summary: string,
  options: {
    fileHash: string;
    package?: Uint8Array | null;
    modelStale?: boolean;
  },
): CheckpointEntry {
  const n = (meta.checkpoints.at(-1)?.n ?? 0) + 1;
  writeFileAtomic(checkpointPath(meta.name, n), state);
  if (options.package) {
    writeFileAtomic(checkpointPath(meta.name, n, "docx"), options.package);
  }
  const entry: CheckpointEntry = {
    n,
    at: new Date().toISOString(),
    command,
    summary,
    ...(options.modelStale ? { modelStale: true } : {}),
  };
  const next: CliDocumentMeta = {
    ...meta,
    fileHash: options.fileHash,
    packageDirty: Boolean(options.modelStale),
    checkpoints: [...meta.checkpoints, entry],
  };
  writeMeta(next);
  Object.assign(meta, next);
  return entry;
}

/** Replace the latest checkpoint's model with one re-learned from its own
 * package: the same checkpoint, now coherent, so one undo still steps back
 * over the patch that caused it. */
export function amendLatestCheckpoint(
  meta: CliDocumentMeta,
  state: Uint8Array,
): number {
  const latest = meta.checkpoints.at(-1);
  const n = latest?.n ?? 0;
  writeFileAtomic(checkpointPath(meta.name, n), state);
  if (latest) delete latest.modelStale;
  meta.packageDirty = false;
  writeMeta(meta);
  return n;
}

/** What checkpoint n holds, without changing anything. */
export function checkpointAt(
  meta: CliDocumentMeta,
  n: number,
): { state: Uint8Array; package: Uint8Array | null; modelStale: boolean } {
  const entry = meta.checkpoints.find((c) => c.n === n);
  if (!entry) {
    throw new CliError(
      `Checkpoint ${n} does not exist.`,
      "checkpoint_not_found",
    );
  }
  return {
    state: readCheckpoint(meta.name, n),
    package: packageAt(meta.name, n),
    modelStale: Boolean(entry.modelStale),
  };
}

// ── saving, journaled ───────────────────────────────────────────────────────

type PendingSave =
  | {
      kind: "commit";
      tx: string;
      n: number;
      command: string;
      summary: string;
      modelStale?: boolean;
      fileHash: string;
    }
  | { kind: "truncate"; tx: string; to: number; fileHash: string };

function pendingPath(name: string): string {
  return path.join(documentDir(name), "pending.json");
}

/** Stage the next checkpoint (not yet in the log) for a file that will hash
 * to `fileHash`. Returns the save's transaction id. */
export function stageCommit(
  meta: CliDocumentMeta,
  state: Uint8Array,
  entry: { command: string; summary: string; modelStale?: boolean },
  fileHash: string,
  pkg: Uint8Array | null = null,
): string {
  const n = (meta.checkpoints.at(-1)?.n ?? 0) + 1;
  const tx = randomBytes(8).toString("hex");
  writeFileAtomic(checkpointPath(meta.name, n), state);
  if (pkg) writeFileAtomic(checkpointPath(meta.name, n, "docx"), pkg);
  const pending: PendingSave = { kind: "commit", tx, n, ...entry, fileHash };
  writeFileAtomic(pendingPath(meta.name), JSON.stringify(pending));
  return tx;
}

/** Stage making checkpoint `to` the latest, for a file hashing to
 * `fileHash`. Returns the save's transaction id. */
export function stageTruncate(
  meta: CliDocumentMeta,
  to: number,
  fileHash: string,
): string {
  const tx = randomBytes(8).toString("hex");
  const pending: PendingSave = { kind: "truncate", tx, to, fileHash };
  writeFileAtomic(pendingPath(meta.name), JSON.stringify(pending));
  return tx;
}

function readPending(name: string): PendingSave | null {
  try {
    return JSON.parse(
      fs.readFileSync(pendingPath(name), "utf8"),
    ) as PendingSave;
  } catch {
    return null;
  }
}

/** Best-effort removal after the commit point: never fails a save. */
function cleanUp(files: string[]): void {
  for (const file of files) {
    try {
      fs.rmSync(file, { force: true });
    } catch {
      // A leftover file is harmless: nothing in the log refers to it.
    }
  }
}

/** Checkpoint files above the log's latest: staged by a save that never
 * committed, or dropped by an undo. */
function unreferencedCheckpointFiles(meta: CliDocumentMeta): string[] {
  const latest = meta.checkpoints.at(-1)?.n ?? 0;
  const dir = path.join(documentDir(meta.name), "checkpoints");
  let names: string[] = [];
  try {
    names = fs.readdirSync(dir);
  } catch {
    return [];
  }
  return names
    .filter((file) => {
      const n = Number.parseInt(file, 10);
      return Number.isFinite(n) && n > latest;
    })
    .map((file) => path.join(dir, file));
}

/**
 * The file is written: commit the staged save to the log (the commit
 * point is meta.json naming its transaction id). Throws only if the commit
 * itself could not be written. Returns the checkpoint the file now is.
 */
export function finishPending(meta: CliDocumentMeta): number {
  const pending = readPending(meta.name);
  if (!pending) throw new CliError("No save in progress.", "internal");
  if (meta.lastTx === pending.tx) {
    // Already committed (an earlier attempt got this far): clean up only.
    cleanUp([pendingPath(meta.name), ...unreferencedCheckpointFiles(meta)]);
    return meta.checkpoints.at(-1)?.n ?? 0;
  }
  let next: CliDocumentMeta;
  if (pending.kind === "commit") {
    const entry: CheckpointEntry = {
      n: pending.n,
      at: new Date().toISOString(),
      command: pending.command,
      summary: pending.summary,
      ...(pending.modelStale ? { modelStale: true } : {}),
    };
    next = {
      ...meta,
      fileHash: pending.fileHash,
      packageDirty: Boolean(pending.modelStale),
      lastTx: pending.tx,
      checkpoints: [...meta.checkpoints, entry],
    };
  } else {
    const kept = meta.checkpoints.filter((entry) => entry.n <= pending.to);
    next = {
      ...meta,
      fileHash: pending.fileHash,
      packageDirty: Boolean(kept.at(-1)?.modelStale),
      lastTx: pending.tx,
      checkpoints: kept,
    };
  }
  writeMeta(next);
  Object.assign(meta, next);
  cleanUp([pendingPath(meta.name), ...unreferencedCheckpointFiles(meta)]);
  return meta.checkpoints.at(-1)?.n ?? 0;
}

/** The file was not written (or was put back): forget the staged save,
 * removing only files the log does not name. */
export function abandonPending(meta: CliDocumentMeta): void {
  cleanUp([pendingPath(meta.name), ...unreferencedCheckpointFiles(meta)]);
}

/** Whether the log on disk already records transaction `tx`. */
export function isCommitted(name: string, tx: string): boolean {
  try {
    return readMeta(name).lastTx === tx;
  } catch {
    return false;
  }
}

/**
 * Settle a save an earlier command left half done, by its transaction id
 * and what the file on disk is now. Returns what happened, for a notice.
 */
export function recoverPending(
  meta: CliDocumentMeta,
  currentFileHash: string,
): "finished" | "discarded" | null {
  const pending = readPending(meta.name);
  if (!pending) return null;
  if (meta.lastTx === pending.tx) {
    finishPending(meta); // committed already: just the cleanup
    return null;
  }
  if (pending.fileHash === currentFileHash) {
    finishPending(meta);
    return "finished";
  }
  abandonPending(meta);
  return "discarded";
}

export function removeDocumentRecord(name: string): void {
  fs.rmSync(documentDir(name), { recursive: true, force: true });
}

function processIsAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}

/**
 * Serialize commands on one file (keyed by its real path). Agents run shell
 * commands in parallel; two first touches of the same file would each
 * import it with fresh block ids and one save would overwrite the other.
 *
 * Lamport's bakery algorithm over files. Each waiter owns one entry file
 * under a unique name, always written whole (temp file, then rename), with
 * its process id: it announces it is choosing, takes a number one above
 * every number it can see, then waits for every other entry that is still
 * choosing and every one with a lower (number, name). An entry whose
 * process is gone is ignored. Nobody ever deletes another's entry while
 * its owner lives, and a waiter that gives up deletes only its own.
 */
export async function withFileLock<T>(
  file: string,
  work: () => Promise<T>,
  options: { timeoutMs?: number } = {},
): Promise<T> {
  const name = sessionNameFor(file);
  const dir = path.join(rootByName.get(name)!, "locks", name);
  fs.mkdirSync(dir, { recursive: true });
  const me = `${Date.now().toString(36)}-${process.pid}-${randomBytes(6).toString("hex")}`;
  const mine = path.join(dir, `${me}.entry`);
  type Entry = { pid: number; host: string; choosing: boolean; number: number };
  const publish = (entry: Entry) =>
    writeFileAtomic(mine, JSON.stringify(entry));
  // Another waiter's entry. Only "not found" means it is gone (its owner
  // finished between listing and reading); any other failure to read it,
  // or contents that aren't a valid entry, block like a live owner would.
  type Seen = { id: string; entry: Entry | "gone" | "unreadable" };
  const isEntry = (value: unknown): value is Entry => {
    const e = value as Partial<Entry> | null;
    return (
      !!e &&
      typeof e.pid === "number" &&
      typeof e.host === "string" &&
      typeof e.choosing === "boolean" &&
      typeof e.number === "number" &&
      Number.isFinite(e.number)
    );
  };
  const others = (): Seen[] =>
    fs
      .readdirSync(dir)
      .filter((f) => f.endsWith(".entry") && f !== `${me}.entry`)
      .map((f): Seen => {
        const id = f.slice(0, -".entry".length);
        let text: string;
        try {
          text = fs.readFileSync(path.join(dir, f), "utf8");
        } catch (error) {
          return {
            id,
            entry:
              (error as NodeJS.ErrnoException).code === "ENOENT"
                ? "gone"
                : "unreadable",
          };
        }
        try {
          const parsed = JSON.parse(text) as unknown;
          return { id, entry: isEntry(parsed) ? parsed : "unreadable" };
        } catch {
          return { id, entry: "unreadable" };
        }
      });
  const alive = (entry: Entry) =>
    entry.host !== os.hostname() || processIsAlive(entry.pid);
  const base = { pid: process.pid, host: os.hostname() };
  publish({ ...base, choosing: true, number: 0 });
  // Choose a number only from a complete view: an entry that can't be
  // read might be a live owner with a higher number, so stay choosing and
  // look again rather than choose too low.
  const deadline = Date.now() + (options.timeoutMs ?? 120_000);
  let number = 0;
  for (;;) {
    const seen = others();
    if (!seen.some(({ entry }) => entry === "unreadable")) {
      number =
        1 +
        Math.max(
          0,
          ...seen.map(({ entry }) =>
            typeof entry === "object" && alive(entry) ? entry.number : 0,
          ),
        );
      break;
    }
    if (Date.now() > deadline) {
      fs.rmSync(mine, { force: true });
      throw new CliError(
        `${path.basename(file)} is busy (another revise command is working on it).`,
        "file_busy",
      );
    }
    await new Promise((resolve) => setTimeout(resolve, 40));
  }
  publish({ ...base, choosing: false, number });
  try {
    for (;;) {
      const blocking = others().some(({ id, entry }) => {
        if (entry === "gone") return false;
        if (entry === "unreadable") return true;
        if (!alive(entry)) {
          // An exited process's entry: its name is unique, so removing it
          // can't touch anyone else's.
          fs.rmSync(path.join(dir, `${id}.entry`), { force: true });
          return false;
        }
        if (entry.choosing) return true;
        return (
          entry.number > 0 &&
          (entry.number < number || (entry.number === number && id < me))
        );
      });
      if (!blocking) break;
      if (Date.now() > deadline) {
        throw new CliError(
          `${path.basename(file)} is busy (another revise command is working on it).`,
          "file_busy",
        );
      }
      await new Promise((resolve) => setTimeout(resolve, 40));
    }
    return await work();
  } finally {
    fs.rmSync(mine, { force: true });
  }
}

/** Several files' locks, always taken in one order (by real path), so two
 * commands that need the same files can never wait on each other. */
export async function withFileLocks<T>(
  files: string[],
  work: () => Promise<T>,
): Promise<T> {
  const ordered = [...new Set(files.map(canonicalPath))].sort();
  const take = (index: number): Promise<T> =>
    index >= ordered.length
      ? work()
      : withFileLock(ordered[index], () => take(index + 1));
  return take(0);
}
