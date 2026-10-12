/** Formats `revise convert` reads, by the name `--from` takes. */
export type InputFormat = "docx" | "odt" | "rtf" | "html" | "md" | "txt";

export const INPUT_FORMATS: readonly InputFormat[] = [
  "docx",
  "odt",
  "rtf",
  "html",
  "md",
  "txt",
];

/** Why bytes on stdin aren't a document revise can read. */
export class UnreadableInput extends Error {}

export interface Sniffed {
  format: InputFormat;
  /** The bytes to import: UTF-16 text is re-encoded as UTF-8. */
  bytes: Uint8Array;
}

const ODF_TEXT = "application/vnd.oasis.opendocument.text";
/** Tags that open a whole HTML document. */
const HTML_DOCUMENT = /^<(!doctype\s+html|html|head|body)[\s>]/i;
/** A tag of any kind, at the start of the text. */
const STARTS_WITH_TAG = /^<[a-z][a-z0-9-]*(\s[^>]*)?\/?>/i;
/** Markdown's own syntax, which HTML never needs: headings, fences, links,
 * emphasis. Looked for in the text with its tags removed. */
const MARKDOWN_SYNTAX = [
  /^\s{0,3}#{1,6}\s+\S/m,
  /^\s{0,3}(```|~~~)/m,
  /\[[^\]\n]+\]\([^)\s]+\)/,
  /(\*\*|__)\S([^\n]*?\S)?\1/,
  /(^|[\s(])[*_][^\s*_]([^*_\n]*[^\s*_])?[*_](?=$|[\s).,;:!?])/m,
];
/** Lists and quotes are Markdown too, but numbered paragraphs ("1. Scope")
 * are common in HTML, so these count only when no HTML block is closed. */
const MARKDOWN_BLOCKS = [
  /^\s{0,3}[-*+]\s+\S/m,
  /^\s{0,3}\d+[.)]\s+\S/m,
  /^\s{0,3}>\s/m,
];
const CLOSES_HTML_BLOCK =
  /<\/(p|div|li|ul|ol|h[1-6]|table|tr|td|th|section|article|blockquote|pre)>/i;

/**
 * Names the format of a document from its bytes. A .docx or .odt is a zip
 * package, told apart by the entries in its central directory; RTF and
 * HTML announce themselves in their first characters. Everything else that
 * is text reads as Markdown, which plain text already is. Markdown that
 * opens with an HTML tag (a README's banner) stays Markdown when it also
 * uses Markdown syntax.
 */
export function sniffFormat(input: Uint8Array): Sniffed {
  if (input.length === 0) {
    throw new UnreadableInput("Nothing came in on stdin.");
  }
  if (isZip(input)) {
    const names = zipEntryNames(input);
    if (names === null) {
      throw new UnreadableInput(
        "stdin is a damaged or incomplete zip file, so it can't be a .docx or .odt.",
      );
    }
    if (names.includes("word/document.xml"))
      return { format: "docx", bytes: input };
    if (names.includes("content.xml") && odfMimetype(input) === ODF_TEXT) {
      return { format: "odt", bytes: input };
    }
    throw new UnreadableInput(
      "stdin is a zip file, but not a .docx or .odt document.",
    );
  }
  if (ascii(input, 0, 5) === "%PDF-") {
    throw new UnreadableInput("revise can't read PDF.");
  }
  const bytes = utf8Of(input);
  const text = new TextDecoder("utf-8").decode(bytes).replace(/^﻿/, "");
  const head = text.trimStart().slice(0, 65536);
  if (head.startsWith("{\\rtf")) return { format: "rtf", bytes };
  if (head.includes("\u0000")) {
    throw new UnreadableInput(
      "stdin is binary data revise doesn't recognize. It reads .docx, .odt, .rtf, HTML, Markdown, and plain text.",
    );
  }
  return { format: looksLikeHtml(head) ? "html" : "md", bytes };
}

function looksLikeHtml(head: string): boolean {
  // Comments and an XML declaration can come before anything.
  const body = head.replace(/^(\s*(<!--[\s\S]*?-->|<\?xml[^>]*\?>))*\s*/, "");
  if (HTML_DOCUMENT.test(body)) return true;
  if (!STARTS_WITH_TAG.test(body)) return false;
  // Text that opens with a tag is HTML unless the rest of it writes
  // Markdown: a README's banner over "# Heading" or "**bold**".
  const text = body.replace(/<[^>]*>/g, "");
  if (MARKDOWN_SYNTAX.some((sign) => sign.test(text))) return false;
  return (
    CLOSES_HTML_BLOCK.test(body) ||
    !MARKDOWN_BLOCKS.some((sign) => sign.test(text))
  );
}

/** The format `--from` names, with the spellings file extensions use. */
export function inputFormatFor(name: string): InputFormat | undefined {
  const key = name.toLowerCase().replace(/^\./, "");
  const aliases: Record<string, InputFormat> = { markdown: "md", htm: "html" };
  const format = aliases[key] ?? key;
  return (INPUT_FORMATS as readonly string[]).includes(format)
    ? (format as InputFormat)
    : undefined;
}

// ── zip ────────────────────────────────────────────────────────────────────

function isZip(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}

/** Entry names from the central directory, or null when the end-of-central-
 * directory record is missing (a truncated or damaged file). */
function zipEntryNames(bytes: Uint8Array): string[] | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The record is 22 bytes plus a comment of up to 65,535.
  const earliest = Math.max(0, bytes.length - 22 - 65535);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= earliest; i -= 1) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;
  let count = view.getUint16(eocd + 10, true);
  let at = view.getUint32(eocd + 16, true);
  if (count === 0xffff || at === 0xffffffff) {
    // ZIP64: the real values are in the ZIP64 end record, which a locator
    // just before this record points to.
    const locator = eocd - 20;
    if (locator < 0 || view.getUint32(locator, true) !== 0x07064b50) {
      return null;
    }
    const record = Number(view.getBigUint64(locator + 8, true));
    if (
      record + 56 > bytes.length ||
      view.getUint32(record, true) !== 0x06064b50
    ) {
      return null;
    }
    count = Number(view.getBigUint64(record + 32, true));
    at = Number(view.getBigUint64(record + 48, true));
  }
  const names: string[] = [];
  for (let n = 0; n < count; n += 1) {
    if (at + 46 > bytes.length || view.getUint32(at, true) !== 0x02014b50) {
      return null;
    }
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    names.push(ascii(bytes, at + 46, nameLength));
    at += 46 + nameLength + extraLength + commentLength;
  }
  return names;
}

/** An ODF package stores its `mimetype` entry first, uncompressed. */
function odfMimetype(bytes: Uint8Array): string | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 30) return null;
  const nameLength = view.getUint16(26, true);
  const extraLength = view.getUint16(28, true);
  if (ascii(bytes, 30, nameLength) !== "mimetype") return null;
  const size = view.getUint32(18, true);
  return ascii(bytes, 30 + nameLength + extraLength, size).trim();
}

// ── text ───────────────────────────────────────────────────────────────────

/**
 * Text as UTF-8: UTF-16 (with a byte-order mark, or recognizable by its
 * zero bytes) is re-encoded, and so is text that isn't valid UTF-8, read as
 * Windows-1252 (what Windows editors write). Valid UTF-8 passes through.
 */
export function utf8Of(bytes: Uint8Array): Uint8Array {
  const encoding = utf16Of(bytes);
  if (encoding) {
    const marked =
      (bytes[0] === 0xff && bytes[1] === 0xfe) ||
      (bytes[0] === 0xfe && bytes[1] === 0xff);
    const bom = marked ? 2 : 0;
    return new TextEncoder().encode(
      new TextDecoder(encoding).decode(bytes.subarray(bom)),
    );
  }
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return bytes;
  } catch {
    if (bytes.includes(0)) return bytes; // binary: refused by the caller
    return new TextEncoder().encode(windows1252(bytes));
  }
}

/** Windows-1252 bytes 0x80-0x9F. The five it doesn't define stay C1. */
const WINDOWS_1252_C1 =
  "€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008dŽ\u008f\u0090‘’“”•–—˜™š›œ\u009džŸ";

/** Windows-1252 text. Node 20's decoder reads 0x80-0x9F as Latin-1 (C1
 * controls instead of € “ ” – —), so those are mapped here. */
function windows1252(bytes: Uint8Array): string {
  return new TextDecoder("windows-1252")
    .decode(bytes)
    .replace(
      /[\u0080-\u009f]/g,
      (char) => WINDOWS_1252_C1[char.charCodeAt(0) - 0x80],
    );
}

function utf16Of(bytes: Uint8Array): "utf-16le" | "utf-16be" | null {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return "utf-16le";
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return "utf-16be";
  // Without a mark: mostly-ASCII UTF-16 has a zero in every other byte.
  const pairs = Math.min(bytes.length, 2048) >> 1;
  if (pairs < 2) return null;
  let even = 0;
  let odd = 0;
  for (let i = 0; i < pairs * 2; i += 2) {
    if (bytes[i] === 0) even += 1;
    if (bytes[i + 1] === 0) odd += 1;
  }
  if (odd > pairs * 0.6 && even < pairs * 0.05) return "utf-16le";
  if (even > pairs * 0.6 && odd < pairs * 0.05) return "utf-16be";
  return null;
}

function ascii(bytes: Uint8Array, start: number, length: number): string {
  let text = "";
  const end = Math.min(bytes.length, start + length);
  for (let i = start; i < end; i += 4096) {
    text += String.fromCharCode(...bytes.subarray(i, Math.min(end, i + 4096)));
  }
  return text;
}
