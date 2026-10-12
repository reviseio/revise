/**
 * Raw OOXML access for the long tail the semantic tools do not cover:
 * style and numbering definitions, settings, document properties.
 *
 * `xml` works on the .docx file itself: cat and ls read it, patch rewrites
 * it — no model, no Y.js. The patched file becomes the graft package, and
 * the next command that reads the model re-learns it from that file.
 */
import { engineModules } from "@reviseio/engine";
import JSZip from "jszip";
import { installDomShims } from "@reviseio/engine";
import { ReviseError as CliError } from "./errors";

/** Parts of a .docx file, from its bytes. */
export async function listParts(
  bytes: Uint8Array,
): Promise<Array<{ path: string; bytes: number }>> {
  const zip = await JSZip.loadAsync(bytes);
  const parts: Array<{ path: string; bytes: number }> = [];
  for (const [path, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    parts.push({ path, bytes: (await entry.async("uint8array")).length });
  }
  return parts.sort((a, b) => a.path.localeCompare(b.path));
}

async function readPart(zip: JSZip, part: string): Promise<string> {
  const file = zip.file(part);
  if (!file) {
    throw new CliError(`No part "${part}" in this document.`, "no_such_part");
  }
  return file.async("string");
}

/** Elements worth showing whole: the smallest of these that contains a
 * match is printed, so the agent sees a complete, patchable unit. */
const UNITS = [
  "w:rPr",
  "w:pPr",
  "w:p",
  "w:tr",
  "w:style",
  "w:abstractNum",
  "w:num",
  "w:lvl",
  "w:sectPr",
  "w:tbl",
];

/** `catPart` returns at most this many matches. */
export const CAT_MATCH_LIMIT = 20;

export async function catPart(
  bytes: Uint8Array,
  part: string,
  match?: string,
  limit = CAT_MATCH_LIMIT,
): Promise<string[]> {
  const xml = await readPart(await JSZip.loadAsync(bytes), part);
  if (!match) return [xml];
  // Distinct elements: several matches in one paragraph show it once.
  const found = new Set<string>();
  let at = xml.indexOf(match);
  while (at >= 0 && found.size < limit) {
    let best: string | undefined;
    for (const tag of UNITS) {
      const start = Math.max(
        xml.lastIndexOf(`<${tag} `, at),
        xml.lastIndexOf(`<${tag}>`, at),
      );
      if (start < 0) continue;
      const close = xml.indexOf(`</${tag}>`, start);
      if (close < 0 || close < at) continue;
      const unit = xml.slice(start, close + tag.length + 3);
      if (!best || unit.length < best.length) best = unit;
    }
    found.add(
      best ?? xml.slice(Math.max(0, at - 200), at + match.length + 200),
    );
    at = xml.indexOf(match, at + match.length);
  }
  return [...found];
}

export interface PatchRequest {
  part: string;
  find?: string;
  replace?: string;
  all?: boolean;
  file?: string;
}

/** Parts Revise's own view is built from: after a patch to one of these,
 * the model re-learns the package before it is next read. */
const MODEL_READ_PARTS =
  /^word\/(document|styles|numbering|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$/;

export interface PatchResult {
  bytes: Uint8Array;
  replacements: number;
  /** The part feeds Revise's own view: the model re-learns the package on
   * the next command that reads it. */
  modelRead: boolean;
}

async function assertWellFormed(part: string, xml: string): Promise<void> {
  await installDomShims();
  const parsed = new DOMParser().parseFromString(xml, "application/xml");
  const error = parsed.getElementsByTagName("parsererror")[0];
  if (error) {
    throw new CliError(
      `The patched ${part} is not well-formed XML: ${error.textContent?.trim().split("\n")[0] ?? "parse error"}. Nothing was changed.`,
      "malformed_xml",
    );
  }
}

export async function patchPackage(
  pkg: Uint8Array,
  request: PatchRequest,
  readFile: (path: string) => string,
): Promise<PatchResult> {
  const zip = await JSZip.loadAsync(pkg);
  const before = await readPart(zip, request.part);
  let after: string;
  let replacements = 0;
  if (request.file !== undefined) {
    after = readFile(request.file);
    replacements = 1;
  } else {
    if (!request.find) {
      throw new CliError(
        "Give --find and --replace, or --file to replace the whole part.",
        "usage",
      );
    }
    const occurrences = before.split(request.find).length - 1;
    if (occurrences === 0) {
      throw new CliError(
        `"${request.find}" does not occur in ${request.part}.`,
        "find_not_found",
      );
    }
    if (occurrences > 1 && !request.all) {
      throw new CliError(
        `"${request.find}" occurs ${occurrences} times in ${request.part}; add context to match once, or pass --all.`,
        "find_ambiguous",
      );
    }
    after = before.split(request.find).join(request.replace ?? "");
    replacements = request.all ? occurrences : 1;
  }
  await assertWellFormed(request.part, after);
  if (request.part === "word/document.xml") {
    // Export maps model blocks onto body positions: the body must keep its
    // shape. Body structure goes through the editing tools.
    const { scanElementChildren } = await engineModules.xmlScan();
    const shape = (xml: string) =>
      scanElementChildren(xml, "w:body")
        .children.map((child) => child.tag)
        .join(",");
    if (shape(before) !== shape(after)) {
      throw new CliError(
        "This patch adds, removes, or reorders body elements, which a direct patch cannot do. Use the editing tools for body structure. Nothing was changed.",
        "body_shape_changed",
      );
    }
  }
  zip.file(request.part, after);
  const bytes = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });
  return {
    bytes,
    replacements,
    modelRead: MODEL_READ_PARTS.test(request.part),
  };
}
