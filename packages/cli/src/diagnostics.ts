/**
 * What the CLI says on stderr while it works.
 *
 * `downword`'s pipeline never fails a document over a broken image link or a
 * construct OOXML cannot hold — it reports and carries on. That is the right
 * behaviour for a library and a trap for a CLI, because a conversion that
 * silently drops three pictures still exits 0. So the diagnostics channel is
 * wired to stderr by default rather than on request:
 *
 * | mode | what is printed |
 * | --- | --- |
 * | `--quiet` | nothing but fatal errors |
 * | default | every warning that **cost the document content**, plus a count of the notices |
 * | `--verbose` | everything, uncapped, plus a line per file |
 *
 * The split is the library's own `severity`: `"error"` means something in the
 * source is not in the `.docx`, `"notice"` means it is all there but expressed
 * differently. Only the first kind is worth interrupting someone for.
 *
 * Everything goes to **stderr**, without exception — stdout may be carrying the
 * `.docx` itself.
 */

import type { Verbosity } from "./options.js";

/** How many warnings one file may print before the rest are counted instead. */
const MAX_LISTED = 20;

/** One thing worth telling the user about a single document. */
export interface Diagnostic {
  /** `"error"`: content was lost. `"notice"`: content survived, in another form. */
  readonly severity: "error" | "notice";
  /** `stage/code`, e.g. `"image/not-found"`. Stable, greppable, machine-filterable. */
  readonly label: string;
  /** Human-readable, already naming the offending construct. */
  readonly message: string;
  /** 1-based source line, when the stage that found it knew one. */
  readonly line: number | null;
  /**
   * Print this one in the default mode even though it is only a `notice`.
   *
   * Exactly one thing uses it: a skipped ```` ```mermaid ```` fence. The library
   * grades that a notice because *nobody's document is wrong* — a CLI has no
   * DOM and never could have drawn it — but the reader still opens the file and
   * finds source code where a diagram was meant to be, so it has to be said out
   * loud rather than folded into "1 notice". Marking it here keeps the notice
   * count honest instead of relabelling the severity to force it through.
   */
  readonly always?: boolean | undefined;
}

/** Collects one document's diagnostics, then prints them under its name. */
export interface Collector {
  /** Records one problem. Never throws. */
  record(diagnostic: Diagnostic): void;
  /** Prints what was collected. Returns how many were errors. */
  flush(): number;
}

/** Writes one line of text, newline included. */
export type LineWriter = (text: string) => void;

/** The stderr channel, in one of the three {@link Verbosity} modes. */
export interface Reporter {
  /** A collector scoped to one input, labelled with `source` in the output. */
  collector(source: string): Collector;
  /** A progress line. `--verbose` only. */
  info(message: string): void;
  /** A failure. Printed in every mode, including `--quiet`. */
  fatal(message: string, hint?: string | null): void;
}

/** `path:12: ` — or `path: ` when the stage attached no line. */
function locate(source: string, line: number | null): string {
  return line === null ? `${source}: ` : `${source}:${line}: `;
}

/**
 * Builds the stderr reporter.
 *
 * @param verbosity - From `--quiet` / `--verbose`.
 * @param write - Receives one ready-to-print line, newline included.
 * @returns The reporter.
 */
export function createReporter(verbosity: Verbosity, write: LineWriter): Reporter {
  const quiet = verbosity === "quiet";
  const verbose = verbosity === "verbose";

  return {
    collector(source: string): Collector {
      const entries: Diagnostic[] = [];
      return {
        record(diagnostic: Diagnostic): void {
          entries.push(diagnostic);
        },
        flush(): number {
          const errorCount = entries.filter((entry) => entry.severity === "error").length;
          if (quiet) return errorCount;

          // Default mode prints what the reader lost; `--verbose` prints the lot.
          const worth = verbose
            ? entries
            : entries.filter((entry) => entry.severity === "error" || entry.always === true);
          const shown = verbose ? worth : worth.slice(0, MAX_LISTED);

          for (const entry of shown) {
            write(`downword: ${locate(source, entry.line)}${entry.label}: ${entry.message}\n`);
          }

          const truncated = worth.length - shown.length;
          if (truncated > 0) {
            write(
              `downword: ${source}: and ${truncated} more warning${truncated === 1 ? "" : "s"} ` +
                `(--verbose to see them all)\n`,
            );
          }

          const unsaid = entries.length - worth.length;
          if (unsaid > 0) {
            write(
              `downword: ${source}: ${unsaid} notice${unsaid === 1 ? "" : "s"} ` +
                `(--verbose to see them)\n`,
            );
          }
          return errorCount;
        },
      };
    },

    info(message: string): void {
      if (verbose) write(`downword: ${message}\n`);
    },

    fatal(message: string, hint?: string | null): void {
      write(`downword: ${message}\n`);
      if (hint !== undefined && hint !== null && hint !== "") write(`  ${hint}\n`);
    },
  };
}
