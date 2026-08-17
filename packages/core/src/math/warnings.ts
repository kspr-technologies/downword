/**
 * The math pass's diagnostics vocabulary.
 *
 * Follows the package-wide contract in `src/warnings.ts`: a closed set of
 * codes, a total severity table so the answer cannot drift between call sites,
 * and a message that already names the offending equation.
 *
 * The line between `"error"` and `"notice"` here is *did the reader lose the
 * equation*. A broken `\frac{a` degrades to the literal characters `\frac{a`,
 * which is the source but is not an equation — that is an error. A construct
 * `mathml2omml` has no OMML for (`\phantom`, `\color`) still renders, just
 * without the cosmetic part — that is a notice.
 */

import type { BaseWarning, WarningSeverity } from "../warnings.js";

/** Machine-readable reason the math pass could not do the ideal thing. */
export type MathWarningCode =
  /** `import("temml")` or `import("mathml2omml")` failed. Raised once per pass. */
  | "engine-unavailable"
  /** TeX that the engine refused: unbalanced braces, an unknown macro, an expansion loop. */
  | "tex-invalid"
  /** The equation was longer than {@link import("./options.js").MathOptions.maxLength}. */
  | "tex-too-large"
  /** A construct with no OMML equivalent was dropped; the rest of the equation converted. */
  | "tex-unsupported"
  /** The converter produced markup this package refuses to embed. See `omml.ts`. */
  | "omml-rejected"
  /** `math: "image"` with no {@link import("../images/types.js").ImageRasterizer} to call. */
  | "image-no-rasterizer"
  /** The rasteriser threw, returned `null`, or returned an unusable raster. */
  | "image-failed";

/** Every {@link MathWarningCode}, for exhaustiveness tests and host-side tables. */
export const MATH_WARNING_CODES = [
  "engine-unavailable",
  "tex-invalid",
  "tex-too-large",
  "tex-unsupported",
  "omml-rejected",
  "image-no-rasterizer",
  "image-failed",
] as const satisfies readonly MathWarningCode[];

const MATH_WARNING_SEVERITY: Readonly<Record<MathWarningCode, WarningSeverity>> = {
  // No engine means *every* equation in the document is now literal TeX.
  "engine-unavailable": "error",
  "tex-invalid": "error",
  "tex-too-large": "error",
  // The equation is in the document and is editable; only a cosmetic layer went.
  "tex-unsupported": "notice",
  "omml-rejected": "error",
  // Both image failures fall back to OMML, which is the better format anyway.
  "image-no-rasterizer": "notice",
  "image-failed": "notice",
};

/** The severity of a math warning code. Total by construction. */
export function mathWarningSeverity(code: MathWarningCode): WarningSeverity {
  return MATH_WARNING_SEVERITY[code];
}

/** A non-fatal problem converting one equation. The document is still produced. */
export interface MathWarning extends BaseWarning {
  readonly code: MathWarningCode;
  /** The TeX source, delimiters stripped, or `""` for pass-wide problems. */
  readonly tex: string;
  /** `true` for `$$…$$`, `false` for `$…$`. */
  readonly display: boolean;
}

/** Called once per {@link MathWarning}. Must not throw; if it does, it is ignored. */
export type MathWarningHandler = (warning: MathWarning) => void;

/** Somewhere to send warnings that swallows handler exceptions. */
export interface MathWarningSink {
  report(code: MathWarningCode, tex: string, display: boolean, message: string): void;
}

/** A sink that forwards to `handler`, or drops everything when there is none. */
export function createMathWarningSink(
  handler?: MathWarningHandler | null | undefined,
): MathWarningSink {
  if (handler === null || handler === undefined) {
    return {
      report() {
        /* no handler: callers need no guard of their own */
      },
    };
  }
  return {
    report(code, tex, display, message) {
      try {
        handler({ code, severity: mathWarningSeverity(code), tex, display, message });
      } catch {
        // A host whose logger throws must not take the document down with it.
      }
    },
  };
}

/** Trims a TeX source down to something a one-line warning can carry. */
export function abbreviateTex(tex: string, limit = 60): string {
  const flat = tex.replace(/\s+/g, " ").trim();
  return flat.length <= limit ? flat : `${flat.slice(0, limit - 1)}…`;
}
