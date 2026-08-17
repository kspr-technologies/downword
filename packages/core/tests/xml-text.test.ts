import { describe, expect, it } from "vitest";

import {
  countXmlUnrepresentable,
  sanitizeXmlText,
  XML_REPLACEMENT_CHARACTER as FFFD,
} from "../src/xml-text.js";

/**
 * The XML 1.0 `Char` production, spelled out as data so the test states the
 * rule rather than restating the implementation:
 *
 * ```text
 * Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
 * ```
 *
 * Every character below is written as an escape sequence rather than typed
 * literally. A raw control byte in a source file is invisible in a diff and
 * silently rewritten by some editors - which is exactly how one ends up in
 * pasted markdown in the first place.
 */
const LEGAL = [
  ["tab", "\u0009"],
  ["line feed", "\u000A"],
  ["carriage return", "\u000D"],
  ["space", "\u0020"],
  ["DEL, which XML 1.0 does allow", "\u007F"],
  ["the last code point before the surrogates", "\uD7FF"],
  ["the first code point after them", "\uE000"],
  ["U+FFFD itself", "\uFFFD"],
  ["an astral character", "\u{1F600}"],
  ["CJK", "中文"],
] as const;

/** The complement: everything OOXML cannot carry, and cannot escape either. */
const ILLEGAL = [
  ["NUL", "\u0000"],
  ["SOH", "\u0001"],
  ["backspace", "\u0008"],
  ["vertical tab", "\u000B"],
  ["form feed", "\u000C"],
  ["shift out", "\u000E"],
  ["unit separator", "\u001F"],
  ["a lone high surrogate", "\uD800"],
  ["a lone low surrogate", "\uDC00"],
  ["noncharacter U+FFFE", "\uFFFE"],
  ["noncharacter U+FFFF", "\uFFFF"],
] as const;

describe("sanitizeXmlText", () => {
  it("returns the very same string when there is nothing to replace", () => {
    const clean = "# Heading\n\nBody with 中文, an emoji \u{1F600} and a\ttab.";
    expect(sanitizeXmlText(clean)).toBe(clean);
  });

  it.each(LEGAL)("leaves %s alone", (_label, value) => {
    expect(sanitizeXmlText(`a${value}b`)).toBe(`a${value}b`);
    expect(countXmlUnrepresentable(`a${value}b`)).toBe(0);
  });

  it.each(ILLEGAL)("replaces %s with U+FFFD", (_label, value) => {
    expect(sanitizeXmlText(`a${value}b`)).toBe(`a${FFFD}b`);
    expect(countXmlUnrepresentable(`a${value}b`)).toBe(1);
  });

  it("is length-preserving, so every source offset survives", () => {
    const dirty = "a\u0000b\u000Cc\uD800d\uFFFFe";
    expect(sanitizeXmlText(dirty)).toHaveLength(dirty.length);
    expect(sanitizeXmlText(dirty)).toBe(`a${FFFD}b${FFFD}c${FFFD}d${FFFD}e`);
  });

  it("does not break a surrogate pair adjacent to a lone surrogate", () => {
    expect(sanitizeXmlText("\uD800\u{1F600}")).toBe(`${FFFD}\u{1F600}`);
    expect(sanitizeXmlText("\u{1F600}\uDC00")).toBe(`\u{1F600}${FFFD}`);
  });

  it("handles a high surrogate at the very end of the string", () => {
    expect(sanitizeXmlText("ok\uD83D")).toBe(`ok${FFFD}`);
  });

  it("counts every replacement", () => {
    expect(countXmlUnrepresentable("\u0001\u0002\u0003")).toBe(3);
    expect(countXmlUnrepresentable("\u{1F600}\u{1F600}")).toBe(0);
    expect(countXmlUnrepresentable("")).toBe(0);
  });

  it("is idempotent", () => {
    const once = sanitizeXmlText("a\u0001b\uD800c");
    expect(sanitizeXmlText(once)).toBe(once);
  });
});
