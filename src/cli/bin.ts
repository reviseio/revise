/** `npx revise` entry point. */
import { exitWhenFlushed } from "./exit";
import { main } from "./main";

// `revise read f.docx | head`: the reader closing early isn't a failure, and
// Node would otherwise throw on the next write.
for (const stream of [process.stdout, process.stderr]) {
  stream.on("error", (error: NodeJS.ErrnoException) => {
    if (error.code === "EPIPE") process.exit(process.exitCode ?? 0);
    throw error;
  });
}

main(process.argv.slice(2)).then(
  (code) => exitWhenFlushed(code),
  (error) => {
    process.stderr.write(
      `revise: internal error: ${(error as Error).stack ?? error}\n`,
    );
    return exitWhenFlushed(70);
  },
);
