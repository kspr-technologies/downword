/**
 * The one error type the CLI throws, and the exit codes it maps onto.
 *
 * Every failure path ends here so that `run()` has exactly one place to decide
 * what the shell sees. The codes are the BSD `sysexits.h` values, which is what
 * `make`, CI runners and shell scripts already understand:
 *
 * | code | name | when |
 * | --- | --- | --- |
 * | 0 | — | the document was written |
 * | 64 | `EX_USAGE` | a bad flag, a bad flag value, or a combination that cannot be honoured (including refusing to write a `.docx` to a terminal) |
 * | 66 | `EX_NOINPUT` | an input file is missing, unreadable, or a pattern matched nothing |
 * | 70 | `EX_SOFTWARE` | the conversion itself failed |
 * | 73 | `EX_CANTCREAT` | the output could not be written |
 *
 * There is deliberately no code for "the document had warnings": a document
 * with a broken image link is still a document, and a build script that treats
 * it as a failure would be wrong.
 */

/** Exit codes, from `sysexits.h`. See the table on the module. */
export const EXIT = {
  ok: 0,
  /** `EX_USAGE` — the command line itself is wrong. */
  usage: 64,
  /** `EX_NOINPUT` — an input does not exist or cannot be read. */
  noInput: 66,
  /** `EX_SOFTWARE` — converting failed. */
  software: 70,
  /** `EX_CANTCREAT` — the output cannot be written. */
  cantCreate: 73,
} as const;

/** One of {@link EXIT}. */
export type ExitCode = (typeof EXIT)[keyof typeof EXIT];

/**
 * A failure with an exit code and, where there is one, a next step.
 *
 * `hint` is printed on its own line under the message. It exists because
 * "no such file" and "no such file — did you mean docs/notes.md?" cost the
 * same to print and not the same to receive.
 */
export class CliError extends Error {
  override readonly name = "CliError";

  readonly exitCode: ExitCode;

  /** One line of "here is what to do about it", or `null`. */
  readonly hint: string | null;

  /**
   * @param exitCode - What the shell should see. See {@link EXIT}.
   * @param message - What went wrong, in one line, without a trailing newline.
   * @param options - `hint` for the follow-up line; `cause` for the original error.
   */
  constructor(
    exitCode: ExitCode,
    message: string,
    options: { readonly hint?: string; readonly cause?: unknown } = {},
  ) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.exitCode = exitCode;
    this.hint = options.hint ?? null;
  }
}

/** Whether `value` is a {@link CliError}. Structural, so a bundled copy still matches. */
export function isCliError(value: unknown): value is CliError {
  return value instanceof Error && value.name === "CliError";
}

/** The message of anything that was thrown, without `[object Object]`. */
export function describeError(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error);
}

/** A `NodeJS.ErrnoException`'s `code`, when it has one. */
export function errnoOf(error: unknown): string | null {
  if (typeof error !== "object" || error === null) return null;
  const code: unknown = (error as { code?: unknown }).code;
  return typeof code === "string" ? code : null;
}
