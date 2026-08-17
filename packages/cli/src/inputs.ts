/**
 * Positional arguments in, a list of files to convert out.
 *
 * A positional is either a path or a pattern, and the difference is decided by
 * {@link hasMagic} rather than by trying one and falling back to the other —
 * so a typo in a plain filename is "no such file", not "matched nothing", and
 * the two get different, specific advice.
 *
 * Patterns are expanded by downword itself, which is why `--help` says to quote
 * them. An unquoted `docs/*.md` is expanded by the shell before the CLI sees
 * it, which works but silently differs: the shell's `*` does not cross
 * directories, `**` usually is not recursive, and a pattern matching nothing is
 * either an error or passed through literally depending on the shell.
 */

import { stat } from "node:fs/promises";
import { basename, dirname, extname, join, resolve } from "node:path";

import { CliError, EXIT, errnoOf } from "./errors.js";
import { expandPattern, hasMagic, type GlobFn } from "./glob.js";
import type { CliOptions } from "./options.js";

/** Where {@link resolveInputs} looks and what it looks with. */
export interface InputContext {
  /** Absolute path a relative positional is resolved against. */
  readonly cwd: string;
  /** `fs.promises.glob`, or `null` to use the built-in matcher. */
  readonly glob: GlobFn | null;
}

/** Turns a failed `stat` into the error the user needs to read. */
function inputError(path: string, error: unknown): CliError {
  const code = errnoOf(error);
  if (code === "ENOENT") {
    return new CliError(EXIT.noInput, `no such file: ${path}`, {
      hint: 'Quote glob patterns so downword expands them: downword "docs/*.md"',
    });
  }
  if (code === "EACCES" || code === "EPERM") {
    return new CliError(EXIT.noInput, `cannot read ${path}: permission denied`);
  }
  return new CliError(EXIT.noInput, `cannot read ${path}`, { cause: error });
}

/**
 * Expands every positional into a de-duplicated list of files.
 *
 * @param patterns - The positional arguments, in the order they were given.
 * @param context - See {@link InputContext}.
 * @returns One entry per file, patterns sorted within themselves and positionals
 *   kept in the order the user wrote them.
 * @throws {CliError} `EXIT.noInput` when a path is missing or a pattern matches
 *   nothing — a command that converted nothing at all is a failed command.
 */
export async function resolveInputs(
  patterns: readonly string[],
  context: InputContext,
): Promise<string[]> {
  const files: string[] = [];
  const seen = new Set<string>();

  const add = (path: string): void => {
    const key = resolve(context.cwd, path);
    if (seen.has(key)) return;
    seen.add(key);
    files.push(path);
  };

  for (const pattern of patterns) {
    if (hasMagic(pattern)) {
      const matches = await expandPattern(pattern, context.cwd, context.glob);
      if (matches.length === 0) {
        throw new CliError(EXIT.noInput, `no files match ${JSON.stringify(pattern)}`, {
          hint: "Patterns are matched by downword, not by the shell: * and ? stay inside one directory, ** crosses them.",
        });
      }
      for (const match of matches) add(match);
      continue;
    }

    let info;
    try {
      info = await stat(resolve(context.cwd, pattern));
    } catch (error: unknown) {
      throw inputError(pattern, error);
    }
    if (info.isDirectory()) {
      throw new CliError(EXIT.noInput, `${pattern} is a directory`, {
        hint: `Convert what is in it: downword "${pattern}/*.md" --outdir build/`,
      });
    }
    add(pattern);
  }

  return files;
}

/** `notes.md` -> `notes.docx`; `notes` -> `notes.docx`; `a.b.md` -> `a.b.docx`. */
export function docxName(input: string): string {
  const name = basename(input);
  const extension = extname(name);
  return `${extension === "" ? name : name.slice(0, -extension.length)}.docx`;
}

/**
 * Where one input's `.docx` goes.
 *
 * @param input - The input path as the user wrote it, or `null` for stdin.
 * @param options - The parsed command line.
 * @returns An absolute-or-relative path, or `null` meaning stdout.
 */
export function outputPathFor(input: string | null, options: CliOptions): string | null {
  if (options.output === "-") return null;
  if (options.output !== null) return options.output;
  if (options.outdir !== null) {
    return join(options.outdir, docxName(input ?? "stdin.md"));
  }
  // Reading stdin with nowhere named to write: stdout, which is what
  // `cat in.md | downword > out.docx` means.
  if (input === null) return null;
  return join(dirname(input), docxName(input));
}
