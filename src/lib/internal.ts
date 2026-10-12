/**
 * Engine functions the CLI needs for its own bookkeeping (the `.revise/`
 * checkpoint cache stores raw Y.js state). Not part of the public API: the
 * `revise` root exports `ReviseDocument` and friends, and nothing here.
 */
export {
  createServerDocumentSession,
  decodeYDoc,
  encodeYDoc,
  fileToYDoc,
} from "@reviseio/engine";
