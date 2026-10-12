/**
 * Exit once everything written to stdout and stderr has been flushed.
 *
 * Writes to a pipe are asynchronous in Node: `process.exit()` right after a
 * large write drops whatever is still buffered (output to a program that
 * pipes `revise --json` was cut at 8 KB). An empty write's callback runs
 * after every earlier write to that stream has been handed to the OS.
 */
export function exitWhenFlushed(
  code: number,
  streams: Array<Pick<NodeJS.WriteStream, "write">> = [
    process.stdout,
    process.stderr,
  ],
  exit: (code: number) => void = (value) => process.exit(value),
): Promise<void> {
  process.exitCode = code;
  return Promise.all(
    streams.map(
      (stream) =>
        new Promise<void>((resolve) => {
          try {
            stream.write("", () => resolve());
          } catch {
            resolve();
          }
        }),
    ),
  ).then(() => exit(code));
}
