/**
 * The command, from `argv` to an exit code.
 *
 * `run()` never throws and never calls `process.exit()`: it returns the code
 * and lets `src/index.ts` set `process.exitCode`, so that a `.docx` still being
 * flushed down a pipe is not truncated by an early exit. Everything it touches
 * arrives through an {@link Io} record, which is what lets the tests drive a
 * fake terminal.
 *
 * ## The four shapes of an invocation
 *
 * ```sh
 * downword notes.md                     # one file  -> notes.docx
 * downword notes.md -o /tmp/out.docx    # one file  -> where you said
 * downword "docs/*.md" --outdir build/  # many      -> one .docx each
 * cat notes.md | downword > out.docx    # stdin     -> stdout
 * ```
 *
 * In batch mode a file that fails does not stop the others: the remaining
 * documents are converted, each failure is reported under its own name, and the
 * process exits with the first failure's code. A build that converts a
 * directory should tell you about all four broken files, not the first one.
 */

import { dirname, resolve } from "node:path";

import { convertMarkdown } from "./convert.js";
import { CliError, EXIT, describeError, isCliError, type ExitCode } from "./errors.js";
import { createReporter, type Reporter } from "./diagnostics.js";
import { nativeGlob } from "./glob.js";
import { HELP } from "./help.js";
import { outputPathFor, resolveInputs } from "./inputs.js";
import {
  assertBinarySafeStdout,
  assertWritablePath,
  ensureDirectory,
  processIo,
  readMarkdownFile,
  readStdin,
  writeDocx,
  type Io,
} from "./io.js";
import { parseCliArgs, type CliOptions } from "./options.js";
import { CLI_VERSION } from "./version.js";

/** `12.3 kB`, in the units a person reads. */
function formatBytes(count: number): string {
  return count < 1024 ? `${count} B` : `${(count / 1024).toFixed(1)} kB`;
}

/** One document: where it came from, where it goes. */
interface Job {
  /** What to call it in messages: the path, or `<stdin>`. */
  readonly label: string;
  /** The path to read, or `null` for stdin. */
  readonly input: string | null;
  /** Where to write, or `null` for stdout. */
  readonly target: string | null;
}

/**
 * Checks the destination before any work happens.
 *
 * Failing here costs the user nothing; failing after the conversion costs them
 * the conversion.
 */
async function assertTargetUsable(
  target: string | null,
  options: CliOptions,
  io: Io,
): Promise<void> {
  if (target === null) {
    assertBinarySafeStdout(io);
    return;
  }
  // --outdir is created for the user (see io.ts), so only -o and the
  // alongside-the-input default need the directory to exist already.
  if (options.outdir === null) await assertWritablePath(target, io.cwd);
}

/** Converts one document and writes it. Throws {@link CliError} on failure. */
async function runJob(job: Job, options: CliOptions, reporter: Reporter, io: Io): Promise<void> {
  const started = Date.now();
  await assertTargetUsable(job.target, options, io);

  const markdown =
    job.input === null ? await readStdin(io.stdin) : await readMarkdownFile(job.input, io.cwd);

  const collector = reporter.collector(job.label);
  let bytes: Uint8Array;
  try {
    bytes = await convertMarkdown({
      markdown,
      // Relative image paths belong to the document, not to the shell's cwd:
      // `![](./diagram.png)` in docs/notes.md is docs/diagram.png wherever the
      // command was run from. Absolute, because the resolver has its own idea
      // of the working directory.
      baseDir: job.input === null ? io.cwd : resolve(io.cwd, dirname(job.input)),
      options,
      report: (diagnostic) => {
        collector.record(diagnostic);
      },
    });
  } finally {
    // Even a failed conversion has usually already found something worth
    // saying; losing it because of what happened afterwards helps nobody.
    collector.flush();
  }

  await writeDocx(job.target, bytes, io);
  reporter.info(
    `${job.label} -> ${job.target ?? "<stdout>"} ` +
      `(${formatBytes(bytes.length)}, ${Date.now() - started} ms)`,
  );
}

/** The stdin -> stdout form. */
async function runStdin(options: CliOptions, reporter: Reporter, io: Io): Promise<ExitCode> {
  if (options.outdir !== null) {
    throw new CliError(EXIT.usage, "--outdir needs input files", {
      hint: "Reading from stdin writes to stdout, or to the file named by -o.",
    });
  }
  await runJob(
    { label: "<stdin>", input: null, target: outputPathFor(null, options) },
    options,
    reporter,
    io,
  );
  return EXIT.ok;
}

/**
 * Refuses a batch in which two inputs would produce the same file.
 *
 * `--outdir` flattens the tree, so `docs/api.md` and `guides/api.md` both want
 * `build/api.docx`. Writing both means the first one is gone and the count in
 * the summary is a lie, so the whole batch stops before anything is written.
 */
function assertDistinctTargets(jobs: readonly Job[], io: Io): void {
  const byTarget = new Map<string, string>();
  for (const job of jobs) {
    if (job.target === null) continue;
    const key = resolve(io.cwd, job.target);
    const first = byTarget.get(key);
    if (first !== undefined) {
      throw new CliError(
        EXIT.usage,
        `${first} and ${job.label} would both be written to ${job.target}`,
        { hint: "Convert them separately, or into directories that mirror the inputs." },
      );
    }
    byTarget.set(key, job.label);
  }
}

/** The one-or-many-files form. */
async function runFiles(options: CliOptions, reporter: Reporter, io: Io): Promise<ExitCode> {
  const files = await resolveInputs(options.patterns, { cwd: io.cwd, glob: nativeGlob() });

  if (options.output !== null && files.length > 1) {
    throw new CliError(
      EXIT.usage,
      `--output takes one input file, but ${files.length} were matched`,
      { hint: "Use --outdir <dir> to convert several documents at once." },
    );
  }

  const jobs: Job[] = files.map((file) => ({
    label: file,
    input: file,
    target: outputPathFor(file, options),
  }));
  assertDistinctTargets(jobs, io);

  if (options.outdir !== null) await ensureDirectory(options.outdir, io.cwd);

  let failure: ExitCode = EXIT.ok;
  let converted = 0;

  for (const job of jobs) {
    try {
      await runJob(job, options, reporter, io);
      converted += 1;
    } catch (error: unknown) {
      if (!isCliError(error)) throw error;
      reporter.fatal(`${job.label}: ${error.message}`, error.hint);
      if (failure === EXIT.ok) failure = error.exitCode;
    }
  }

  if (files.length > 1) {
    reporter.info(`converted ${converted} of ${files.length} files`);
  }
  return failure;
}

/**
 * Runs the CLI.
 *
 * @param argv - Arguments after `node <script>`, i.e. `process.argv.slice(2)`.
 * @param io - The process's edges. Defaults to the real ones.
 * @returns The exit code. See `EXIT` in `./errors.ts`.
 */
export async function run(argv: readonly string[], io: Io = processIo()): Promise<ExitCode> {
  let options: CliOptions;
  try {
    options = parseCliArgs(argv);
  } catch (error: unknown) {
    // Verbosity is not known yet — a usage error is printed whatever it was.
    const failure = isCliError(error)
      ? error
      : new CliError(EXIT.usage, describeError(error), { cause: error });
    io.stderr.write(`downword: ${failure.message}\n`);
    if (failure.hint !== null) io.stderr.write(`  ${failure.hint}\n`);
    return failure.exitCode;
  }

  if (options.version) {
    io.stdout.write(`${CLI_VERSION}\n`);
    return EXIT.ok;
  }
  // No arguments and no pipe: the person is looking for the manual, not waiting
  // for a document that will never arrive on a terminal's stdin.
  if (options.help || (options.patterns.length === 0 && io.stdinIsTty)) {
    io.stdout.write(HELP);
    return EXIT.ok;
  }

  const reporter = createReporter(options.verbosity, (text) => {
    io.stderr.write(text);
  });

  try {
    return options.patterns.length === 0
      ? await runStdin(options, reporter, io)
      : await runFiles(options, reporter, io);
  } catch (error: unknown) {
    if (isCliError(error)) {
      reporter.fatal(error.message, error.hint);
      return error.exitCode;
    }
    // Nothing is supposed to reach this. If something does, say so plainly and
    // exit non-zero rather than printing a stack trace at a user.
    reporter.fatal(
      `internal error: ${describeError(error)}`,
      "Please report this at https://github.com/kspr-technologies/downword/issues",
    );
    return EXIT.software;
  }
}
