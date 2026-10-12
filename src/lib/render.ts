/**
 * Page rendering for the CLI: the production agent page renderer (the one
 * behind the native agent's render_document_pages) drawn on a Node canvas.
 *
 * Loaded lazily — it pulls in the full editor and a native canvas module, and
 * most commands never need either.
 */
import { engineModules } from "@reviseio/engine";
import type * as Y from "yjs";
import { ReviseError } from "./errors";

export type RenderView = "tracked" | "accepted" | "original";

const VIEW_NAMES = {
  tracked: "show-tracked-changes",
  accepted: "show-with-tracked-changes-accepted",
  original: "show-before-tracked-changes",
} as const;

export interface RenderOptions {
  /** First page, 1-based. Default 1. */
  startPage?: number;
  /** Pages in one image, 1-6. Default 1. */
  pageCount?: number;
  /** Tracked changes shown as marks (default), accepted, or before them. */
  view?: RenderView;
  /** Overlay layout boxes (margins, block bounds) for debugging. */
  layout?: boolean;
  /** "detailed" (default for one page) or "compact". */
  quality?: "compact" | "detailed";
}

export interface RenderResult {
  bytes: Uint8Array;
  mimeType: string;
  width: number;
  height: number;
  pages: number[];
  totalPages: number;
  placeholderImageCount: number;
  failedImageCount: number;
  description: string;
}

type CreateCanvas = (width: number, height: number) => unknown;

async function loadCanvasFactory(): Promise<CreateCanvas> {
  const attempts: string[] = [];
  for (const moduleName of ["@napi-rs/canvas", "canvas"]) {
    try {
      // @ts-ignore optional native peer, resolved at runtime only
      const mod = (await import(/* @vite-ignore */ moduleName)) as {
        createCanvas?: CreateCanvas;
        default?: { createCanvas?: CreateCanvas };
      };
      const create = mod.createCanvas ?? mod.default?.createCanvas;
      if (create) return create;
    } catch (error) {
      attempts.push(
        `${moduleName}: ${(error as Error).message.split("\n")[0]}`,
      );
    }
  }
  throw new ReviseError(
    `Rendering pages needs a Node canvas. Install one next to revise: npm i @napi-rs/canvas\n(${attempts.join("; ")})`,
    "canvas_missing",
  );
}

/** Lay out the document and draw the requested pages as one image. */
export async function renderPages(
  ydoc: Y.Doc,
  documentId: string,
  options: RenderOptions,
): Promise<RenderResult> {
  const createCanvas = await loadCanvasFactory();
  const [
    { DocumentSession },
    { renderDocumentPagesForAgent, createHeadlessPageRenderHost },
    { createStandardLayout },
    { resolvePageLayoutSettings, getPageDimensionsPx },
  ] = await Promise.all([
    engineModules.documentSession(),
    engineModules.pageRenderer(),
    engineModules.documentLayout(),
    engineModules.pageLayout(),
  ]);
  const session = new DocumentSession(documentId, ydoc, undefined, {
    disableHistory: true,
    disableProposalPersistence: true,
  });
  // Pages render from the session alone: the interactive editor (input,
  // caret, selection) is not part of the headless library.
  const editor = createHeadlessPageRenderHost(session);
  const settings = resolvePageLayoutSettings(
    session.getWorkingDocument().pageLayout,
  );
  const dimensions = getPageDimensionsPx(settings);
  let result: Awaited<ReturnType<typeof renderDocumentPagesForAgent>>;
  try {
    result = await renderDocumentPagesForAgent({
      editor,
      pageLayoutSettings: settings,
      layout: createStandardLayout({
        ...dimensions,
        pageMargin: settings.marginLeftPx,
        pageMarginTop: settings.marginTopPx,
        pageMarginRight: settings.marginRightPx,
        pageMarginBottom: settings.marginBottomPx,
        pageMarginLeft: settings.marginLeftPx,
        indentStepPx: settings.indentStepPx,
        blockSpacingPx: settings.blockSpacingPx,
      }),
      request: {
        startPage: options.startPage ?? 1,
        pageCount: options.pageCount ?? 1,
        mode: options.layout ? "layout" : "appearance",
        quality:
          options.quality ??
          ((options.pageCount ?? 1) > 1 ? "compact" : "detailed"),
        view: VIEW_NAMES[options.view ?? "tracked"],
      },
      createCanvas: () => createCanvas(1, 1) as unknown as HTMLCanvasElement,
    });
  } catch (error) {
    const message = (error as Error).message;
    if (/outside this \d+-page document/.test(message)) {
      throw new ReviseError(message, "page_out_of_range");
    }
    if (error instanceof ReviseError) throw error;
    throw new ReviseError(`Rendering failed: ${message}`, "render_failed", {
      cause: error,
    });
  } finally {
    session.dispose();
  }
  const match = /^data:([^;]+);base64,([\s\S]*)$/.exec(result.dataUrl);
  if (!match) throw new Error("The renderer returned an unexpected image.");
  return {
    bytes: new Uint8Array(Buffer.from(match[2], "base64")),
    mimeType: match[1],
    width: result.width,
    height: result.height,
    pages: result.pages,
    totalPages: result.totalPages,
    placeholderImageCount: result.placeholderImageCount,
    failedImageCount: result.failedImageCount,
    description: result.snapshot.description,
  };
}
