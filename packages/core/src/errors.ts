/**
 * The one error type downword throws.
 *
 * Markdown has no syntax errors — the parser reports what it cannot represent
 * through `ConvertOptions.onWarning` and carries on — so nothing a user *types*
 * can fail a conversion. What can fail is everything around it: a plugin that
 * throws, an option out of range, a `Packer` that cannot allocate. Every one of
 * those arrives here, tagged with a {@link DownwordErrorCode} and carrying the
 * original as `cause`, rather than escaping as whatever markdown-it or docx
 * happened to throw.
 *
 * ```ts
 * import { convert, isDownwordError } from "@ksprtech/downword";
 *
 * try {
 *   await convert("# hi");
 * } catch (error) {
 *   if (isDownwordError(error) && error.code === "pack-failed") {
 *     console.error("could not build the zip", error.cause);
 *   }
 * }
 * ```
 */

/**
 * Machine-readable reason a {@link DownwordError} was raised.
 *
 * A small closed set, so a caller can branch on it instead of matching prose.
 */
export type DownwordErrorCode =
  /** `markdown` was not a string. Only reachable from JavaScript. */
  | "invalid-input"
  /** An option was outside the range its TSDoc documents. */
  | "invalid-options"
  /** markdown-it, or one of `ConvertOptions.plugins`, threw. */
  | "parse-failed"
  /** The renderer threw. Always a bug in downword — please report it. */
  | "render-failed"
  /** `docx`'s `Packer` could not build the `.docx` zip. */
  | "pack-failed"
  /** A documented entry point that is still a stub. */
  | "not-implemented";

/**
 * Every error downword throws deliberately.
 *
 * @see {@link DownwordErrorCode} for the tags, and {@link isDownwordError} for
 *   a check that works across a mixed ESM/CJS graph where `instanceof` does not.
 */
export class DownwordError extends Error {
  override readonly name = "DownwordError";

  /** Machine-readable reason; see {@link DownwordErrorCode}. */
  readonly code: DownwordErrorCode;

  /**
   * @param code - Machine-readable reason.
   * @param message - Human-readable description, safe to show to an end user.
   * @param options - Standard `ErrorOptions`; pass the original error as `cause`.
   */
  constructor(code: DownwordErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.code = code;
  }
}

/**
 * Whether `value` is a {@link DownwordError}.
 *
 * Prefer this to `instanceof`: a project that ends up with both the ESM and the
 * CJS build in its graph (a bundled app importing a CJS-only dependency that
 * also uses downword) has two distinct classes, and `instanceof` silently
 * returns `false` for one of them. This is a structural check, so it does not
 * care which copy minted the error.
 */
export function isDownwordError(value: unknown): value is DownwordError {
  if (!(value instanceof Error)) return false;
  const candidate = value as Partial<DownwordError>;
  return candidate.name === "DownwordError" && typeof candidate.code === "string";
}

/**
 * Wraps whatever a subsystem threw in a {@link DownwordError}.
 *
 * A `DownwordError` from further down (a plugin that uses this class, say)
 * passes through unchanged, so the innermost — and most specific — code
 * survives instead of being relabelled by its caller.
 *
 * @internal
 */
export function wrapError(code: DownwordErrorCode, message: string, cause: unknown): DownwordError {
  if (isDownwordError(cause)) return cause;
  const detail = cause instanceof Error ? cause.message : String(cause);
  return new DownwordError(code, `${message}: ${detail}`, { cause });
}
