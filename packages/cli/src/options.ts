/**
 * The command line: `parseArgs` in, one validated {@link CliOptions} out.
 *
 * Two rules this file keeps.
 *
 * **Nothing from `@ksprtech/downword` is imported here.** The enumerations below are
 * literal copies of the library's, so that `downword --help` and
 * `downword --version` answer without loading the converter, `docx` or
 * markdown-it at all. `tests/vocabulary.test.ts` compares each list against the
 * library's own and fails if they drift, which is the trade: a copy that is
 * checked, in exchange for a CLI that starts in milliseconds.
 *
 * **A bad value is a usage error, not a stack trace.** Every flag is validated
 * here, before a file is opened, and the message names the flag, the value it
 * was given and the values it accepts.
 */

import { parseArgs } from "node:util";

import { CliError, EXIT } from "./errors.js";
import { TRY_HELP } from "./help.js";

/* -------------------------------------------------------------------------- */
/* The vocabulary                                                              */
/* -------------------------------------------------------------------------- */

/** Built-in theme names. Mirrors `THEMES` in `@ksprtech/downword`. */
export const THEME_NAMES = ["default", "github", "academic", "academic-double", "print"] as const;

/** A built-in theme name. */
export type ThemeName = (typeof THEME_NAMES)[number];

/** Paper sizes `--page-size` accepts. Mirrors `PageSize`'s string form. */
export const PAGE_SIZES = ["A4", "Letter"] as const;

/** A paper size. */
export type PageSizeName = (typeof PAGE_SIZES)[number];

/** Equation handling. Mirrors `MathMode` in `@ksprtech/downword/plugins/math`. */
export const MATH_MODES = ["omml", "image", "off"] as const;

/** How equations are handled. */
export type MathMode = (typeof MATH_MODES)[number];

/** How much the CLI says while it works. */
export type Verbosity = "quiet" | "normal" | "verbose";

/* -------------------------------------------------------------------------- */
/* The result                                                                  */
/* -------------------------------------------------------------------------- */

/** Everything the command line asked for, validated. */
export interface CliOptions {
  /** Positional arguments: file paths, glob patterns, or neither (stdin). */
  readonly patterns: readonly string[];
  /** `--output`, or `null`. `"-"` means stdout. */
  readonly output: string | null;
  /** `--outdir`, or `null`. */
  readonly outdir: string | null;
  readonly theme: ThemeName;
  readonly pageSize: PageSizeName;
  readonly math: MathMode;
  readonly toc: boolean;
  readonly pageNumbers: boolean;
  /** `false` when `--no-footnotes` was given. */
  readonly footnotes: boolean;
  readonly title: string | null;
  readonly author: string | null;
  readonly highlight: boolean;
  /** `true` only when `--remote-images` was given. Default-deny. */
  readonly remoteImages: boolean;
  readonly allowOutsideImages: boolean;
  readonly verbosity: Verbosity;
  readonly help: boolean;
  readonly version: boolean;
}

/* -------------------------------------------------------------------------- */
/* Parsing                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * The flag table.
 *
 * `--no-remote-images` and `--no-footnotes` are declared as their own booleans
 * rather than relying on `parseArgs`'s `allowNegative`, which landed in Node
 * 22.4 — this package supports Node 20. Declaring them also means they appear
 * in `--help` as flags in their own right, which is the point of
 * `--no-remote-images`: the default-deny is a documented flag, not an inference.
 */
const FLAGS = {
  output: { type: "string", short: "o" },
  outdir: { type: "string" },
  theme: { type: "string" },
  "page-size": { type: "string" },
  math: { type: "string" },
  title: { type: "string" },
  author: { type: "string" },
  toc: { type: "boolean" },
  "page-numbers": { type: "boolean" },
  footnotes: { type: "boolean" },
  "no-footnotes": { type: "boolean" },
  highlight: { type: "boolean" },
  "remote-images": { type: "boolean" },
  "no-remote-images": { type: "boolean" },
  "allow-outside-images": { type: "boolean" },
  quiet: { type: "boolean", short: "q" },
  verbose: { type: "boolean" },
  version: { type: "boolean", short: "v" },
  help: { type: "boolean", short: "h" },
} as const;

/** Turns a `parseArgs` failure into a usage error carrying the help nudge. */
function usage(message: string, hint: string = TRY_HELP): CliError {
  return new CliError(EXIT.usage, message, { hint });
}

/** Validates one string flag against a closed set. */
function oneOf<const T extends readonly string[]>(
  value: string | undefined,
  allowed: T,
  flag: string,
  fallback: T[number],
): T[number] {
  if (value === undefined) return fallback;
  const match = allowed.find((candidate) => candidate === value);
  if (match === undefined) {
    throw usage(
      `--${flag} must be one of ${allowed.join(", ")}; got ${JSON.stringify(value)}`,
      `Example: downword --${flag} ${allowed[0] ?? ""} input.md`,
    );
  }
  return match;
}

/** Rejects `--flag ""`, which is a mistyped command rather than a real value. */
function nonEmpty(value: string | undefined, flag: string): string | null {
  if (value === undefined) return null;
  if (value.trim() === "") throw usage(`--${flag} needs a value`);
  return value;
}

/** Rejects a pair of flags that contradict each other. */
function exclusive(a: boolean, aName: string, b: boolean, bName: string, hint: string): void {
  if (a && b) throw usage(`${aName} and ${bName} contradict each other`, hint);
}

/**
 * Parses and validates `argv`.
 *
 * @param argv - Arguments after `node <script>`, i.e. `process.argv.slice(2)`.
 * @returns The validated options.
 * @throws {CliError} `EXIT.usage` for anything wrong with the command line.
 */
export function parseCliArgs(argv: readonly string[]): CliOptions {
  let parsed;
  try {
    parsed = parseArgs({
      args: [...argv],
      options: FLAGS,
      allowPositionals: true,
      strict: true,
    });
  } catch (error: unknown) {
    // parseArgs' own messages already name the offending token ("Unknown option
    // '--nope'"), so they are kept verbatim rather than paraphrased.
    throw usage(error instanceof Error ? error.message : String(error));
  }

  const values = parsed.values;

  exclusive(
    values.quiet === true,
    "--quiet",
    values.verbose === true,
    "--verbose",
    "Pick one: --quiet prints only fatal errors, --verbose prints everything.",
  );
  exclusive(
    values["remote-images"] === true,
    "--remote-images",
    values["no-remote-images"] === true,
    "--no-remote-images",
    "Remote images are off unless --remote-images is given.",
  );
  exclusive(
    values.footnotes === true,
    "--footnotes",
    values["no-footnotes"] === true,
    "--no-footnotes",
    "Footnotes are on unless --no-footnotes is given.",
  );
  exclusive(
    values.output !== undefined,
    "--output",
    values.outdir !== undefined,
    "--outdir",
    "Use --output for one file, --outdir to convert several at once.",
  );

  const output = values.output === undefined ? null : values.output;
  if (output !== null && output !== "-" && output.trim() === "") {
    throw usage('--output needs a path, or "-" for stdout');
  }

  return {
    patterns: parsed.positionals,
    output,
    outdir: nonEmpty(values.outdir, "outdir"),
    theme: oneOf(values.theme, THEME_NAMES, "theme", "default"),
    pageSize: oneOf(values["page-size"], PAGE_SIZES, "page-size", "A4"),
    // Off by default: `$` is a currency symbol far more often than it opens an
    // equation, and `--math omml` needs two optional peer packages the CLI does
    // not depend on. Opting in is one flag; un-mangling a price list is not.
    math: oneOf(values.math, MATH_MODES, "math", "off"),
    toc: values.toc === true,
    pageNumbers: values["page-numbers"] === true,
    footnotes: values["no-footnotes"] !== true,
    title: nonEmpty(values.title, "title"),
    author: nonEmpty(values.author, "author"),
    highlight: values.highlight === true,
    remoteImages: values["remote-images"] === true,
    allowOutsideImages: values["allow-outside-images"] === true,
    verbosity: values.quiet === true ? "quiet" : values.verbose === true ? "verbose" : "normal",
    help: values.help === true,
    version: values.version === true,
  };
}
