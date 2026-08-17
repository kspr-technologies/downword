/**
 * The copied enumerations, checked against the originals.
 *
 * `src/options.ts` lists the theme names, page sizes and maths modes itself
 * rather than importing them, so that `--help` does not load the converter.
 * That is a copy, and a copy is a bug waiting for a release — unless something
 * compares it. This is that something: it imports the library (which a test may
 * do freely) and asserts the two agree, so adding a theme to `downword` without
 * teaching the CLI about it fails here rather than in someone's terminal.
 */

import { THEMES } from "downword";
import { describe, expect, it } from "vitest";

import { MATH_MODES, PAGE_SIZES, THEME_NAMES } from "../src/options.js";

describe("the CLI's vocabulary matches the library's", () => {
  it("has exactly the library's themes", () => {
    expect([...THEME_NAMES].sort()).toEqual(Object.keys(THEMES).sort());
  });

  it("accepts every theme it advertises", async () => {
    const { convert } = await import("downword");

    for (const theme of THEME_NAMES) {
      await expect(convert("# hi", { theme })).resolves.toBeInstanceOf(Uint8Array);
    }
  });

  it("accepts every page size it advertises", async () => {
    const { convert } = await import("downword");

    for (const pageSize of PAGE_SIZES) {
      await expect(convert("# hi", { pageSize })).resolves.toBeInstanceOf(Uint8Array);
    }
  });

  it("accepts every maths mode it advertises", async () => {
    const { resolveMathOptions } = await import("downword/plugins/math");

    for (const math of MATH_MODES) {
      expect(resolveMathOptions({ math }).math).toBe(math);
    }
  });
});
