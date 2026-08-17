/**
 * `downword` — the binary.
 *
 * The command itself lives in the modules beside this one; what is left here is
 * the part that only makes sense as a process:
 *
 *  - it sets `process.exitCode` rather than calling `process.exit()`, because a
 *    `.docx` on its way down a pipe is still in a buffer when the last line of
 *    `run()` executes, and `process.exit()` would truncate it; and
 *  - it turns an unhandled rejection into one line and exit 70, rather than
 *    Node's crash banner.
 *
 * The module is also the package's entry point, so `run` and `CLI_VERSION` are
 * re-exported for anyone embedding it. Importing it *does* run the command —
 * which is what a `bin` is for.
 *
 * @module
 */

import { EXIT } from "./errors.js";
import { run } from "./run.js";

export { CliError, EXIT, type ExitCode } from "./errors.js";
export { run } from "./run.js";
export { CLI_VERSION } from "./version.js";

process.exitCode = await run(process.argv.slice(2)).catch((error: unknown) => {
  process.stderr.write(`downword: internal error: ${String(error)}\n`);
  return EXIT.software;
});
