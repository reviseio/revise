/**
 * The one error type the library throws for expected failures (a missing
 * file, an unsupported format, a malformed patch). `code` is stable and
 * machine-readable; the CLI prints it as `error <code>: <message>`.
 */
export class ReviseError extends Error {
  /** The underlying error, when this one wraps another (a filesystem or
   * parser failure). */
  readonly cause?: unknown;

  constructor(
    message: string,
    readonly code = "error",
    options: { cause?: unknown } = {},
  ) {
    super(message);
    this.name = "ReviseError";
    if (options.cause !== undefined) this.cause = options.cause;
  }
}
