/**
 * WCAG 2.x contrast, for tests that make an accessibility claim.
 *
 * downword's docs promise specific ratios — `THEMES.print` says every ink
 * clears AAA against the `CodeBlock` shading — and a promise like that is only
 * worth making if something recomputes it. Shared rather than duplicated so the
 * palette suite and the regression suite can never disagree about the maths.
 */

function channel(value: number): number {
  const srgb = value / 255;
  return srgb <= 0.04045 ? srgb / 12.92 : Math.pow((srgb + 0.055) / 1.055, 2.4);
}

/** WCAG 2.x relative luminance of an `RRGGBB` colour. */
export function luminance(hex: string): number {
  const value = Number.parseInt(hex, 16);
  return (
    0.2126 * channel((value >> 16) & 0xff) +
    0.7152 * channel((value >> 8) & 0xff) +
    0.0722 * channel(value & 0xff)
  );
}

/** WCAG 2.x contrast ratio between two `RRGGBB` colours, 1:1 to 21:1. */
export function contrast(a: string, b: string): number {
  const [lighter, darker] = [luminance(a), luminance(b)].sort((x, y) => y - x) as [number, number];
  return (lighter + 0.05) / (darker + 0.05);
}

/** WCAG's two bars for normal-size text. */
export const WCAG_AA = 4.5;
/** WCAG's two bars for normal-size text. */
export const WCAG_AAA = 7;
