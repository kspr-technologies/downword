/**
 * Option validation, shared by the package's two public option types.
 *
 * The rule stated in `options.ts`'s docblock — *"Bad values fail loudly… instead
 * of producing a document Word quietly refuses to open"* — is a promise about
 * **downword**, not about `convert()`. `renderDocument(model, options)` is
 * exported and documented as one of "both halves of the pipeline", so a value
 * it cannot use has to fail the same way, with the same code, in the same
 * words. Keeping the primitives here rather than duplicating them is what makes
 * that literally true instead of approximately true.
 *
 * Everything in this file throws {@link DownwordError} with code
 * `"invalid-options"` and a message that names the option path (`options.page.
 * margin.top`), the constraint, and the value it got.
 */

import { DownwordError } from "./errors.js";

/** An `invalid-options` error. */
export function invalid(message: string): DownwordError {
  return new DownwordError("invalid-options", message);
}

/** Narrows an untrusted value to one of `allowed`, or throws. */
export function oneOf<T extends string>(
  value: unknown,
  allowed: readonly T[],
  option: string,
  fallback: T,
): T {
  if (value === undefined) return fallback;
  if (typeof value === "string" && (allowed as readonly string[]).includes(value))
    return value as T;
  throw invalid(
    `options.${option} must be one of ${allowed.map((a) => `"${a}"`).join(", ")}, got ${JSON.stringify(value)}`,
  );
}

/** A finite, non-negative measurement in twips. */
export function twips(value: number | undefined, option: string, fallback: number): number {
  if (value === undefined) return fallback;
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw invalid(`options.${option} must be a finite number of twips >= 0, got ${String(value)}`);
  }
  return value;
}

/** An integer in `[min, max]`, or throws. */
export function integer(
  value: number | undefined,
  option: string,
  fallback: number,
  min: number,
  max: number,
): number {
  if (value === undefined) return fallback;
  if (typeof value !== "number" || !Number.isInteger(value) || value < min || value > max) {
    throw invalid(
      `options.${option} must be an integer between ${min} and ${max}, got ${String(value)}`,
    );
  }
  return value;
}

/** A boolean, or throws. */
export function flag(value: boolean | undefined, option: string, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  if (typeof value !== "boolean") {
    throw invalid(`options.${option} must be a boolean, got ${JSON.stringify(value)}`);
  }
  return value;
}
