/**
 * Turning a quoted pattern such as `"docs/*.md"` into a list of files.
 *
 * `fs.promises.glob` landed in Node 22 and this package supports Node 20, so
 * there are two implementations and a feature test. Node's is used wherever it
 * exists; the fallback below covers the same syntax on Node 20 —
 * `*`, `?`, `**`, `[abc]`, `[!abc]` and `{a,b}` — with the same three
 * behaviours that matter:
 *
 *  - a wildcard never matches a leading dot (`*.md` skips `.eslintrc.md`), and
 *    a `**` never descends into a dot directory, unless the pattern says so
 *    literally;
 *  - `**` matches zero or more path segments; and
 *  - symbolic links to directories are not followed, which is also what makes
 *    the walk cycle-proof.
 *
 * `tests/glob.test.ts` runs both implementations over the same tree and asserts
 * they return the same files, so the Node 20 leg of CI is not a different
 * program from the Node 22 one. The alternative — depending on `fast-glob` or
 * `tinyglobby` — would put a transitive dependency tree behind
 * `npx @ksprtech/downword-cli`, for a feature the platform now ships.
 *
 * Neither implementation returns directories: this is a list of *files to
 * convert*, and `docs/*.md` happily matches a directory named `notes.md`.
 */

import * as fsp from "node:fs/promises";
import { join, resolve } from "node:path";

const { readdir, stat } = fsp;

/** The shape of `fs.promises.glob`, as much of it as is used here. */
export type GlobFn = (pattern: string) => AsyncIterable<string>;

/** Characters that make a positional a pattern rather than a path. */
const MAGIC = /[*?[\]{}]/;

/** Whether a positional argument needs expanding. */
export function hasMagic(pattern: string): boolean {
  return MAGIC.test(pattern);
}

/* -------------------------------------------------------------------------- */
/* Brace expansion                                                             */
/* -------------------------------------------------------------------------- */

/** Splits `a,b{c,d}` on its top-level commas only. */
function splitAlternatives(body: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let start = 0;
  for (let index = 0; index < body.length; index += 1) {
    const char = body[index];
    if (char === "{") depth += 1;
    else if (char === "}") depth -= 1;
    else if (char === "," && depth === 0) {
      parts.push(body.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(body.slice(start));
  return parts;
}

/** `a{b,c}d` -> `["abd", "acd"]`. Nested braces expand too; unbalanced ones are literal. */
export function expandBraces(pattern: string): string[] {
  const open = pattern.indexOf("{");
  if (open === -1) return [pattern];

  let depth = 0;
  let close = -1;
  for (let index = open; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === "{") depth += 1;
    else if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        close = index;
        break;
      }
    }
  }
  if (close === -1) return [pattern];

  const head = pattern.slice(0, open);
  const tail = pattern.slice(close + 1);
  return splitAlternatives(pattern.slice(open + 1, close)).flatMap((part) =>
    expandBraces(`${head}${part}${tail}`),
  );
}

/* -------------------------------------------------------------------------- */
/* One path segment                                                            */
/* -------------------------------------------------------------------------- */

/** Escapes a literal character for use inside a regular expression. */
function escapeRegex(char: string): string {
  return /[\\^$.*+?()[\]{}|]/.test(char) ? `\\${char}` : char;
}

/**
 * Compiles one path segment into an anchored regular expression.
 *
 * `[` only opens a character class when it is closed later in the segment;
 * otherwise it is a literal, which is what a shell does and what stops
 * `notes[draft.md` from being a syntax error.
 */
export function segmentToRegExp(segment: string): RegExp {
  let source = "^";
  let index = 0;
  while (index < segment.length) {
    const char = segment[index] ?? "";
    if (char === "*") {
      source += "[^/]*";
      index += 1;
    } else if (char === "?") {
      source += "[^/]";
      index += 1;
    } else if (char === "[") {
      // A `]` immediately after the opening (or after the negation) is literal.
      let scan = index + 1;
      if (segment[scan] === "!" || segment[scan] === "^") scan += 1;
      if (segment[scan] === "]") scan += 1;
      while (scan < segment.length && segment[scan] !== "]") scan += 1;
      if (scan >= segment.length) {
        source += "\\[";
        index += 1;
        continue;
      }
      const body = segment.slice(index + 1, scan);
      const negated = body.startsWith("!") || body.startsWith("^");
      const set = (negated ? body.slice(1) : body).replace(/\\/g, "\\\\");
      source += `[${negated ? "^" : ""}${set}]`;
      index = scan + 1;
    } else {
      source += escapeRegex(char);
      index += 1;
    }
  }
  return new RegExp(`${source}$`);
}

/** Whether a name is hidden and the pattern did not ask for hidden names. */
function hiddenMismatch(name: string, segment: string): boolean {
  return name.startsWith(".") && !segment.startsWith(".");
}

/* -------------------------------------------------------------------------- */
/* The fallback walk                                                           */
/* -------------------------------------------------------------------------- */

/** A directory listing, or nothing at all when the directory cannot be read. */
async function listing(dir: string): Promise<{ name: string; directory: boolean }[]> {
  try {
    const entries = await readdir(dir, { withFileTypes: true });
    // `isDirectory()` is false for a symlink, so links are neither descended
    // into nor mistaken for files - the walk cannot cycle.
    return entries.map((entry) => ({ name: entry.name, directory: entry.isDirectory() }));
  } catch {
    return [];
  }
}

/** Whether `path` is a regular file. */
async function isFile(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile();
  } catch {
    return false;
  }
}

/** Walks `dir`, matching the remaining pattern segments, collecting file paths. */
async function walk(
  dir: string,
  relative: string,
  segments: readonly string[],
  found: Set<string>,
): Promise<void> {
  const [segment, ...rest] = segments;
  if (segment === undefined) return;
  const last = rest.length === 0;

  if (segment === "**") {
    // Zero segments, then one-or-more: `**` stays in play as we descend.
    await walk(dir, relative, rest, found);
    for (const entry of await listing(dir)) {
      if (entry.directory && !entry.name.startsWith(".")) {
        await walk(join(dir, entry.name), join(relative, entry.name), segments, found);
      } else if (last && !entry.directory && !entry.name.startsWith(".")) {
        // A trailing `**` means "this directory and everything under it", which
        // is what `docs/**` matches in Node's implementation and in a shell.
        found.add(join(relative, entry.name));
      }
    }
    return;
  }

  if (!hasMagic(segment)) {
    const next = join(dir, segment);
    const nextRelative = join(relative, segment);
    if (last) {
      if (await isFile(next)) found.add(nextRelative);
    } else {
      await walk(next, nextRelative, rest, found);
    }
    return;
  }

  const pattern = segmentToRegExp(segment);
  for (const entry of await listing(dir)) {
    if (hiddenMismatch(entry.name, segment) || !pattern.test(entry.name)) continue;
    const nextRelative = join(relative, entry.name);
    if (last) {
      if (!entry.directory) found.add(nextRelative);
    } else if (entry.directory) {
      await walk(join(dir, entry.name), nextRelative, rest, found);
    }
  }
}

/**
 * Expands one pattern without `fs.glob`.
 *
 * @param pattern - Forward-slash separated, relative to `cwd` or absolute.
 * @param cwd - Where a relative pattern starts.
 * @returns Matching **files**, sorted, relative to `cwd` (absolute if the
 *   pattern was).
 */
export async function globFallback(pattern: string, cwd: string): Promise<string[]> {
  const found = new Set<string>();
  for (const expanded of expandBraces(pattern)) {
    const absolute = expanded.startsWith("/");
    const segments = expanded.split("/").filter((part) => part !== "" && part !== ".");
    await walk(absolute ? "/" : cwd, absolute ? "/" : "", segments, found);
  }
  return [...found].sort();
}

/* -------------------------------------------------------------------------- */
/* The front door                                                              */
/* -------------------------------------------------------------------------- */

/**
 * `fs.promises.glob` if this Node has it, otherwise `null`.
 *
 * The feature test is a runtime one on purpose: `@types/node` describes the
 * Node 22 surface, so TypeScript is certain this function exists and Node 20 is
 * certain it does not.
 */
export function nativeGlob(): GlobFn | null {
  const glob: unknown = fsp.glob;
  return typeof glob === "function" ? (glob as GlobFn) : null;
}

/**
 * Expands one pattern, preferring the platform's implementation.
 *
 * @param pattern - The glob.
 * @param cwd - Where a relative pattern starts.
 * @param glob - `fs.promises.glob`, or `null` to force the fallback.
 * @returns Matching files, sorted. Directories are filtered out.
 */
export async function expandPattern(
  pattern: string,
  cwd: string,
  glob: GlobFn | null,
): Promise<string[]> {
  if (glob === null) return globFallback(pattern, cwd);

  const matches: string[] = [];
  // Node's glob resolves relative patterns against process.cwd(), which is the
  // only cwd the CLI ever uses; the fallback takes it explicitly so it can be
  // pointed at a fixture tree in tests.
  for await (const match of glob(pattern)) {
    if (await isFile(resolve(cwd, match))) matches.push(match);
  }
  return matches.sort();
}
