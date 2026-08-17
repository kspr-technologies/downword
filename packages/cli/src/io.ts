/**
 * Reading stdin, writing the `.docx`, and refusing to garble a terminal.
 *
 * Everything the CLI touches outside of `downword` itself goes through here,
 * behind an {@link Io} record rather than `process` globals, so `run()` can be
 * driven in a test with a fake terminal — which is the only way to check the
 * TTY guard without allocating a pty.
 *
 * ## The TTY guard
 *
 * A `.docx` is a zip. Writing one to a terminal prints a few kilobytes of
 * control characters, some of which the terminal will obey, and the user's next
 * prompt is in a random colour with a broken charset. Every tool that emits
 * binary refuses to do this (`gzip`, `tar cf -`, `curl -o -`), and so does
 * this one — *before* converting, so the answer arrives immediately rather than
 * after the work.
 */

import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";

import { CliError, EXIT, errnoOf } from "./errors.js";

/** The process's edges: three streams and a working directory. */
export interface Io {
  readonly stdin: NodeJS.ReadableStream;
  /** `isTTY` is what the guard reads; it is absent on a pipe or a file. */
  readonly stdout: NodeJS.WritableStream & { readonly isTTY?: boolean | undefined };
  readonly stderr: NodeJS.WritableStream;
  /** `isTTY` again: no stdin and no arguments means "print the help". */
  readonly stdinIsTty: boolean;
  readonly cwd: string;
}

/** The real process. */
export function processIo(): Io {
  return {
    stdin: process.stdin,
    stdout: process.stdout,
    stderr: process.stderr,
    stdinIsTty: process.stdin.isTTY === true,
    cwd: process.cwd(),
  };
}

/** Strips a UTF-8 byte-order mark, which markdown-it would otherwise treat as text. */
function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/**
 * Reads all of stdin as UTF-8.
 *
 * @param stream - Usually `process.stdin`.
 * @returns The text, without a byte-order mark.
 * @throws {CliError} `EXIT.noInput` if the stream errors.
 */
export async function readStdin(stream: NodeJS.ReadableStream): Promise<string> {
  const chunks: Buffer[] = [];
  try {
    for await (const chunk of stream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk, "utf8") : chunk);
    }
  } catch (error: unknown) {
    throw new CliError(EXIT.noInput, "cannot read stdin", { cause: error });
  }
  return stripBom(Buffer.concat(chunks).toString("utf8"));
}

/**
 * Reads a markdown file as UTF-8.
 *
 * @param path - As the user wrote it; resolved against `cwd`.
 * @param cwd - The working directory.
 * @returns The text, without a byte-order mark.
 * @throws {CliError} `EXIT.noInput`, naming the file and the reason.
 */
export async function readMarkdownFile(path: string, cwd: string): Promise<string> {
  try {
    return stripBom(await readFile(resolve(cwd, path), "utf8"));
  } catch (error: unknown) {
    const code = errnoOf(error);
    if (code === "EACCES" || code === "EPERM") {
      throw new CliError(EXIT.noInput, `cannot read ${path}: permission denied`, { cause: error });
    }
    if (code === "EISDIR") {
      throw new CliError(EXIT.noInput, `${path} is a directory`, { cause: error });
    }
    throw new CliError(EXIT.noInput, `cannot read ${path}`, { cause: error });
  }
}

/**
 * Refuses to write a `.docx` into a terminal.
 *
 * @param io - The process's edges.
 * @throws {CliError} `EXIT.usage` when stdout is a terminal.
 */
export function assertBinarySafeStdout(io: Io): void {
  if (io.stdout.isTTY !== true) return;
  throw new CliError(EXIT.usage, "refusing to write a .docx to the terminal", {
    hint: "Redirect it (downword in.md > out.docx) or name a file (-o out.docx).",
  });
}

/**
 * Checks that a file can plausibly be created at `path` before doing the work.
 *
 * Only the parent directory is checked, and deliberately not created: `-o` names
 * one file, and a typo in its directory should be a message rather than a new
 * tree of empty folders. `--outdir` is the flag that means "make me a place to
 * put these", and {@link ensureDirectory} is what serves it.
 *
 * @param path - The output path, as given.
 * @param cwd - The working directory.
 * @throws {CliError} `EXIT.cantCreate`.
 */
export async function assertWritablePath(path: string, cwd: string): Promise<void> {
  const parent = dirname(resolve(cwd, path));
  let info;
  try {
    info = await stat(parent);
  } catch (error: unknown) {
    if (errnoOf(error) === "ENOENT") {
      throw new CliError(EXIT.cantCreate, `no such directory: ${dirname(path)}`, {
        hint: "Create it first, or use --outdir <dir>, which downword creates for you.",
      });
    }
    throw new CliError(EXIT.cantCreate, `cannot write ${path}`, { cause: error });
  }
  if (!info.isDirectory()) {
    throw new CliError(EXIT.cantCreate, `not a directory: ${dirname(path)}`);
  }
}

/**
 * Creates `--outdir`, including any missing parents.
 *
 * @param dir - The directory, as given.
 * @param cwd - The working directory.
 * @throws {CliError} `EXIT.cantCreate`.
 */
export async function ensureDirectory(dir: string, cwd: string): Promise<void> {
  try {
    await mkdir(resolve(cwd, dir), { recursive: true });
  } catch (error: unknown) {
    throw new CliError(EXIT.cantCreate, `cannot create ${dir}`, { cause: error });
  }
}

/** Writes to a stream and waits for the platform to take the bytes. */
async function push(stream: NodeJS.WritableStream, bytes: Uint8Array): Promise<void> {
  await new Promise<void>((settle, fail) => {
    stream.write(bytes, (error?: Error | null) => {
      if (error === undefined || error === null) settle();
      else fail(error);
    });
  });
}

/**
 * Writes the document.
 *
 * @param path - Where to put it, or `null` for stdout.
 * @param bytes - The `.docx`.
 * @param io - The process's edges.
 * @throws {CliError} `EXIT.cantCreate` when the file cannot be written.
 */
export async function writeDocx(path: string | null, bytes: Uint8Array, io: Io): Promise<void> {
  if (path === null) {
    assertBinarySafeStdout(io);
    try {
      await push(io.stdout, bytes);
    } catch (error: unknown) {
      // `downword in.md | head -c 100` closes the pipe early. That is the
      // reader's decision, not a failure of ours.
      if (errnoOf(error) === "EPIPE") return;
      throw new CliError(EXIT.cantCreate, "cannot write to stdout", { cause: error });
    }
    return;
  }

  try {
    await writeFile(resolve(io.cwd, path), bytes);
  } catch (error: unknown) {
    const code = errnoOf(error);
    if (code === "EACCES" || code === "EPERM") {
      throw new CliError(EXIT.cantCreate, `cannot write ${path}: permission denied`, {
        cause: error,
      });
    }
    if (code === "ENOENT") {
      throw new CliError(EXIT.cantCreate, `no such directory: ${dirname(path)}`, {
        hint: "Create it first, or use --outdir <dir>, which downword creates for you.",
      });
    }
    throw new CliError(EXIT.cantCreate, `cannot write ${path}`, { cause: error });
  }
}
