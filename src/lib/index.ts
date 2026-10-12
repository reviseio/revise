/**
 * `revise`: read, edit, review, render, and convert documents from Node.
 *
 *   const doc = await ReviseDocument.open("contract.docx", { author: "Legal bot" });
 *   await doc.replaceAll("Provider", "Northstar");      // tracked changes
 *   await doc.reject({ section: "8" });
 *   const { bytes } = await doc.export({ format: "docx" });
 *   fs.writeFileSync("contract-out.docx", bytes);
 *
 * The `revise` CLI is built on exactly this API.
 */
export {
  ReviseDocument,
  compare,
  convert,
  formatForPath,
  seedYDoc,
  type ChangeEntry,
  type ChangeFilter,
  type ChangeSelection,
  type CompareResult,
  type DocumentMode,
  type DocumentSource,
  type DocumentSummary,
  type ExportFormat,
  type ExportOptions,
  type ExportResult,
  type GrepHit,
  type GrepMatch,
  type ReadView,
  type OpenOptions,
  type ReplaceAllOptions,
  type ReplaceTextResult,
  type TextRegion,
  type SettleResult,
  type TextDiff,
} from "./document";
export {
  agentTools,
  formatToolResult,
  type AgentProvider,
  type AgentTools,
  type AgentToolsOptions,
  type AnthropicToolUse,
  type OpenAIResponsesFunctionCall,
  type OpenAIToolCall,
} from "./agent";
export {
  catPart,
  listParts,
  patchPackage,
  type PatchRequest,
  type PatchResult,
} from "./package";
// One headless API: ReviseDocument and the helpers above. These types
// appear in its signatures (doc.session, tool results, image policy).
export type {
  ImageFetchPolicy,
  ReviseJsonObject,
  ReviseToolName,
  ReviseToolResult,
  ServerDocumentSession,
} from "@reviseio/engine";
export type { RenderOptions, RenderResult, RenderView } from "./render";
export { ReviseError } from "./errors";
