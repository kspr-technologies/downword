import { describe, expect, it, vi } from "vitest";

import {
  convert,
  convertToBlob,
  convertToDocument,
  DOCX_MIME_TYPE,
  DownwordError,
  isDownwordError,
  THEMES,
  VERSION,
  type ConvertOptions,
  type ConvertWarning,
  type Highlighter,
  type ImageResolver,
  type MarkdownItPlugin,
} from "../src/index.js";
import { listDocxParts, readDocxPart } from "./helpers/docx.js";
import { describeScrubbers, normalizeOoxml } from "./helpers/normalize.js";
import { paragraphs, pStyle } from "./helpers/xml.js";

/** PK\x03\x04 - the local file header magic every zip (and therefore docx) starts with. */
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];

/** A real 8x4 PNG (solid #2F5496), as a data URI the default resolver accepts. */
const PNG_DATA_URI =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAEUlEQVR42mPQD5mGFTFQTwIA0B4jIb0iwL8AAAAASUVORK5CYII=";

/** Converts and hands back one unzipped part. */
async function partOf(
  markdown: string,
  part: string,
  options: ConvertOptions = {},
): Promise<string> {
  return readDocxPart(await convert(markdown, options), part);
}

/** Converts and collects every warning, in emission order. */
async function warningsOf(
  markdown: string,
  options: ConvertOptions = {},
): Promise<ConvertWarning[]> {
  const warnings: ConvertWarning[] = [];
  await convert(markdown, {
    ...options,
    onWarning: (warning) => {
      warnings.push(warning);
    },
  });
  return warnings;
}

/* -------------------------------------------------------------------------- */
/* Shape of the output                                                         */
/* -------------------------------------------------------------------------- */

describe("convert", () => {
  it("returns .docx bytes", async () => {
    const bytes = await convert("# Hello\n\nWorld.");

    expect(bytes).toBeInstanceOf(Uint8Array);
    expect([...bytes.slice(0, 4)]).toEqual(ZIP_MAGIC);
    expect(bytes.byteLength).toBeGreaterThan(1000);
  });

  it("emits the mandatory OOXML parts", async () => {
    const parts = await listDocxParts(await convert("hello"));

    expect(parts).toContain("[Content_Types].xml");
    expect(parts).toContain("word/document.xml");
    expect(parts).toContain("word/styles.xml");
    expect(parts).toContain("docProps/core.xml");
  });

  it("actually converts markdown rather than emitting plain paragraphs", async () => {
    const xml = await partOf("# Title\n\n- one\n- two\n\nsome `code`", "word/document.xml");

    expect(xml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(xml).toContain("<w:numPr>");
    expect(xml).toContain('<w:rStyle w:val="CodeChar"/>');
  });

  it("produces a valid, empty document for empty input", async () => {
    const bytes = await convert("");

    expect([...bytes.slice(0, 4)]).toEqual(ZIP_MAGIC);
    // A section with no children is a body Word treats as damaged, so the
    // renderer must still emit one empty paragraph - styled `Normal` like every
    // other body paragraph, so Design -> Style Set reaches it too.
    const document = await readDocxPart(bytes, "word/document.xml");
    expect(paragraphs(document)).toHaveLength(1);
    expect(pStyle(paragraphs(document)[0] ?? "")).toBe("Normal");
  });

  it("is deterministic: the same markdown converts to the same document", async () => {
    const first = normalizeOoxml(await partOf("# a\n\n1. b\n2. c", "word/document.xml"));
    const second = normalizeOoxml(await partOf("# a\n\n1. b\n2. c", "word/document.xml"));

    expect(second).toBe(first);
  });

  it("exposes the package version as a build-time constant", () => {
    expect(VERSION).toMatch(/^\d+\.\d+\.\d+/);
  });
});

/* -------------------------------------------------------------------------- */
/* convertToBlob / convertToDocument                                           */
/* -------------------------------------------------------------------------- */

describe("convertToBlob", () => {
  it("tags the blob with the OOXML mime type", async () => {
    const blob = await convertToBlob("# Hello");

    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe(DOCX_MIME_TYPE);
    expect(blob.size).toBeGreaterThan(1000);
  });

  it("carries the same document as convert()", async () => {
    const blob = await convertToBlob("# Hello");
    const bytes = new Uint8Array(await blob.arrayBuffer());

    expect([...bytes.slice(0, 4)]).toEqual(ZIP_MAGIC);
    expect(normalizeOoxml(await readDocxPart(bytes, "word/document.xml"))).toBe(
      normalizeOoxml(await partOf("# Hello", "word/document.xml")),
    );
  });
});

describe("convertToDocument", () => {
  it("returns a docx Document the caller can pack themselves", async () => {
    const { Packer } = await import("docx");
    const file = await convertToDocument("# Hello");

    // "UEsDB" is "PK\x03\x04" in base64.
    expect((await Packer.toBase64String(file)).startsWith("UEsDB")).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* Errors                                                                      */
/* -------------------------------------------------------------------------- */

describe("errors", () => {
  it("rejects non-string input with invalid-input", async () => {
    // The type system already forbids this; JavaScript callers do not have it.
    await expect(convert(42 as unknown as string)).rejects.toThrow(DownwordError);
    await expect(convert(null as unknown as string)).rejects.toMatchObject({
      code: "invalid-input",
      message: expect.stringContaining("null"),
    });
  });

  it.each([
    ["an unknown theme name", { theme: "midnight" } as unknown as ConvertOptions],
    ["a negative margin", { margins: -1 } satisfies ConvertOptions],
    ["a non-finite margin", { margins: { left: Number.NaN } } satisfies ConvertOptions],
    ["a zero page dimension", { pageSize: { width: 0, height: 100 } } satisfies ConvertOptions],
    ["an unknown html mode", { html: "sanitise" } as unknown as ConvertOptions],
    ["an unknown lineBreaks mode", { lineBreaks: "double" } as unknown as ConvertOptions],
    ["an unknown orientation", { orientation: "sideways" } as unknown as ConvertOptions],
  ])("rejects %s with invalid-options", async (_label, options) => {
    await expect(convert("# hi", options)).rejects.toMatchObject({ code: "invalid-options" });
  });

  it("wraps a throwing markdown-it plugin as parse-failed, keeping the cause", async () => {
    const boom = new Error("plugin exploded");
    const plugin: MarkdownItPlugin = () => {
      throw boom;
    };

    const error: unknown = await convert("# hi", { plugins: [plugin] }).catch((e: unknown) => e);

    expect(isDownwordError(error)).toBe(true);
    expect(error).toMatchObject({ code: "parse-failed", cause: boom });
    expect((error as DownwordError).message).toContain("plugin exploded");
  });

  it("isDownwordError is structural, so it survives a duplicated class", () => {
    class Impostor extends Error {
      override readonly name = "DownwordError";
      readonly code = "pack-failed";
    }

    expect(isDownwordError(new DownwordError("pack-failed", "x"))).toBe(true);
    expect(isDownwordError(new Impostor())).toBe(true);
    expect(isDownwordError(new Error("plain"))).toBe(false);
    expect(isDownwordError({ name: "DownwordError", code: "pack-failed" })).toBe(false);
    expect(isDownwordError(null)).toBe(false);
  });

  it("never throws for markdown a person could plausibly type", async () => {
    const nasty = [
      "```unclosed",
      "| a | b |\n| - |",
      "[^dangling]",
      "![](../../nope.png)",
      "<div><span>",
      "***",
      "1) a\n1) b",
      "| |\n|-|\n| |",
    ].join("\n\n");

    await expect(convert(nasty, { html: "keep" })).resolves.toBeInstanceOf(Uint8Array);
  });
});

/* -------------------------------------------------------------------------- */
/* Warnings                                                                    */
/* -------------------------------------------------------------------------- */

describe("warnings", () => {
  it("reports parse warnings with a source line", async () => {
    const parse = (await warningsOf("intro\n\n<div>raw</div>\n", { html: "keep" })).filter(
      (warning) => warning.stage === "parse",
    );

    expect(parse.length).toBeGreaterThan(0);
    expect(parse[0]).toMatchObject({ stage: "parse", code: "raw-html", line: 3 });
  });

  it("reports render warnings", async () => {
    expect(await warningsOf("[jump](#nowhere)")).toContainEqual(
      expect.objectContaining({ stage: "render", code: "link-unresolved" }),
    );
  });

  it("reports image diagnostics, including the resolver's own reason", async () => {
    const images = (await warningsOf("![remote](https://example.invalid/a.png)")).filter(
      (warning) => warning.stage === "image",
    );

    expect(images).toContainEqual(
      expect.objectContaining({
        stage: "image",
        code: "remote-blocked",
        severity: "error",
        src: "https://example.invalid/a.png",
      }),
    );
  });

  it("stays silent when nothing is wrong", async () => {
    expect(await warningsOf("# Fine\n\nJust prose.")).toEqual([]);
  });

  it("survives a handler that throws", async () => {
    const onWarning = vi.fn(() => {
      throw new Error("logger exploded");
    });

    await expect(convert("[jump](#nowhere)", { onWarning })).resolves.toBeInstanceOf(Uint8Array);
    expect(onWarning).toHaveBeenCalled();
  });
});

/* -------------------------------------------------------------------------- */
/* Options                                                                     */
/* -------------------------------------------------------------------------- */

describe("options: page geometry", () => {
  it("defaults to A4 portrait with one-inch margins", async () => {
    const xml = await partOf("hi", "word/document.xml");

    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>');
    expect(xml).toContain('w:top="1440"');
  });

  it("honours pageSize, orientation and margins", async () => {
    const xml = await partOf("hi", "word/document.xml", {
      pageSize: "Letter",
      orientation: "landscape",
      margins: { top: 720, right: 600, bottom: 720, left: 600 },
    });

    // docx swaps the two dimensions itself for a landscape page, so Letter's
    // long edge (15840) becomes the width.
    expect(xml).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>');
    expect(xml).toContain('w:left="600"');
    expect(xml).toContain('w:top="720"');
  });

  it("accepts one number for all four margins", async () => {
    expect(await partOf("hi", "word/document.xml", { margins: 0 })).toContain(
      '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"',
    );
  });

  it("accepts explicit page dimensions in twips", async () => {
    expect(
      await partOf("hi", "word/document.xml", { pageSize: { width: 15840, height: 12240 } }),
    ).toContain('<w:pgSz w:w="15840" w:h="12240"');
  });
});

describe("options: metadata", () => {
  it("writes the core properties Word shows under File -> Info", async () => {
    const core = await partOf("hi", "docProps/core.xml", {
      metadata: {
        title: "Quarterly Report",
        author: "A. Writer",
        subject: "Revenue",
        description: "Generated by a test",
        keywords: ["docx", "markdown"],
      },
    });

    expect(core).toContain("<dc:title>Quarterly Report</dc:title>");
    expect(core).toContain("<dc:creator>A. Writer</dc:creator>");
    expect(core).toContain("<dc:subject>Revenue</dc:subject>");
    expect(core).toContain("<dc:description>Generated by a test</dc:description>");
    expect(core).toContain("<cp:keywords>docx, markdown</cp:keywords>");
  });

  it("writes custom properties", async () => {
    const custom = await partOf("hi", "docProps/custom.xml", {
      metadata: { custom: { Department: "Finance" } },
    });

    expect(custom).toContain('name="Department"');
    expect(custom).toContain("Finance");
  });

  it("does not repeat the title in the body by default", async () => {
    const withoutBlock = await partOf("body text", "word/document.xml", {
      metadata: { title: "Repeated" },
    });
    const withBlock = await partOf("body text", "word/document.xml", {
      metadata: { title: "Repeated" },
      titleBlock: true,
    });

    expect(withoutBlock).not.toContain('<w:pStyle w:val="Title"/>');
    expect(withBlock).toContain('<w:pStyle w:val="Title"/>');
  });
});

describe("options: markdown dialect", () => {
  it("escapes raw HTML by default", async () => {
    const xml = await partOf("<b>bold?</b>", "word/document.xml");

    expect(xml).toContain("&lt;b&gt;");
    expect(xml).not.toContain('<w:pStyle w:val="HtmlBlock"/>');
  });

  it('keeps raw HTML verbatim with html: "keep"', async () => {
    expect(await partOf("<div>\nraw\n</div>", "word/document.xml", { html: "keep" })).toContain(
      '<w:pStyle w:val="HtmlBlock"/>',
    );
  });

  it('removes raw HTML with html: "drop"', async () => {
    const xml = await partOf("before\n\n<div>\nDROPPED\n</div>\n\nafter", "word/document.xml", {
      html: "drop",
    });

    expect(xml).toContain("before");
    expect(xml).toContain("after");
    expect(xml).not.toContain("DROPPED");
  });

  it("linkifies bare URLs, and can be told not to", async () => {
    expect(await partOf("see https://example.com now", "word/document.xml")).toContain(
      "<w:hyperlink",
    );
    expect(
      await partOf("see https://example.com now", "word/document.xml", { linkify: false }),
    ).not.toContain("<w:hyperlink");
  });

  it("leaves punctuation alone unless typographer is on", async () => {
    expect(await partOf('"quoted"', "word/document.xml")).toContain("&quot;quoted&quot;");
    expect(await partOf('"quoted"', "word/document.xml", { typographer: true })).toContain(
      "“quoted”",
    );
  });

  it("collapses single newlines unless lineBreaks is preserve", async () => {
    expect(await partOf("one\ntwo", "word/document.xml")).not.toContain("<w:br/>");
    expect(await partOf("one\ntwo", "word/document.xml", { lineBreaks: "preserve" })).toContain(
      "<w:br/>",
    );
  });

  it("parses footnotes, and can be told not to", async () => {
    const source = "A claim[^1]\n\n[^1]: The source.";

    expect(await listDocxParts(await convert(source))).toContain("word/footnotes.xml");
    expect(await partOf(source, "word/document.xml", { footnotes: false })).toContain("[^1]");
  });

  it("applies markdown-it plugins in order", async () => {
    const seen: string[] = [];
    const first: MarkdownItPlugin = () => {
      seen.push("first");
    };
    const second: MarkdownItPlugin = () => {
      seen.push("second");
    };

    await convert("hi", { plugins: [first, second] });
    expect(seen).toEqual(["first", "second"]);
  });
});

describe("options: theme", () => {
  it("defaults to the default theme", async () => {
    expect(await partOf("hi", "word/styles.xml")).toContain('w:ascii="Calibri"');
  });

  it("merges a partial ThemeInit over the default", async () => {
    const styles = await partOf("hi", "word/styles.xml", { theme: { fonts: { body: "Georgia" } } });

    expect(styles).toContain('w:ascii="Georgia"');
    // Untouched tokens still come from the default theme.
    expect(styles).toContain('w:ascii="Calibri Light"');
  });

  it("accepts a built-in theme by name", async () => {
    await expect(convert("hi", { theme: "print" })).resolves.toBeInstanceOf(Uint8Array);
    await expect(convert("hi", { theme: "default" })).resolves.toBeInstanceOf(Uint8Array);
  });

  it('the "print" theme differs from the default only in its code palette', () => {
    expect(THEMES.print.codePalette["keyword"]).not.toBe(THEMES.default.codePalette["keyword"]);
    expect(THEMES.print.fonts).toEqual(THEMES.default.fonts);
    expect(THEMES.print.headings).toEqual(THEMES.default.headings);
  });

  it("accepts a fully resolved theme as a starting point", async () => {
    expect(
      await partOf("hi", "word/styles.xml", {
        theme: { ...THEMES.print, fonts: { ...THEMES.print.fonts, mono: "IBM Plex Mono" } },
      }),
    ).toContain('w:ascii="IBM Plex Mono"');
  });
});

/* -------------------------------------------------------------------------- */
/* Extension points                                                            */
/* -------------------------------------------------------------------------- */

describe("options: images", () => {
  it("embeds data: URIs with no resolver configured", async () => {
    const bytes = await convert(`![logo](${PNG_DATA_URI})`);

    expect((await listDocxParts(bytes)).some((part) => part.startsWith("word/media/"))).toBe(true);
    expect(await readDocxPart(bytes, "word/document.xml")).toContain("<w:drawing>");
  });

  it("cannot reach the network by default", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      const warnings = await warningsOf("![remote](https://example.invalid/a.png)");

      expect(fetchSpy).not.toHaveBeenCalled();
      expect(
        warnings.some((warning) => warning.stage === "image" && warning.code === "remote-blocked"),
      ).toBe(true);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("uses a caller-supplied resolver instead of the default", async () => {
    const resolver: ImageResolver = { resolve: vi.fn(async () => null) };

    await convert("![a](whatever.png)", { imageResolver: resolver });
    expect(resolver.resolve).toHaveBeenCalledWith("whatever.png");
  });

  it("renders a placeholder rather than failing when an image cannot resolve", async () => {
    expect(await partOf("![missing](nope.png)", "word/document.xml")).toContain(
      '<w:pStyle w:val="ImagePlaceholder"/>',
    );
  });
});

describe("options: highlighter", () => {
  const redEverything: Highlighter = {
    highlight: (code) => [{ text: code, color: "FF0000" }],
  };

  it("is not called when absent, and code renders plain", async () => {
    const xml = await partOf("```ts\nconst x = 1;\n```", "word/document.xml");

    expect(xml).toContain('<w:pStyle w:val="CodeBlock"/>');
    expect(xml).not.toContain('<w:color w:val="FF0000"/>');
  });

  it("colours code blocks when supplied", async () => {
    expect(
      await partOf("```ts\nconst x = 1;\n```", "word/document.xml", {
        highlighter: redEverything,
      }),
    ).toContain('<w:color w:val="FF0000"/>');
  });

  it("loses only the block when a highlighter throws", async () => {
    const flaky: Highlighter = {
      highlight: (code, lang) => {
        if (lang === "ts") throw new Error("no grammar");
        return [{ text: code, color: "00FF00" }];
      },
    };

    const xml = await partOf("```ts\nalpha\n```\n\n```sh\nbravo\n```", "word/document.xml", {
      highlighter: flaky,
    });

    expect(xml).toContain('<w:color w:val="00FF00"/>');
    expect(xml).toContain("alpha");
  });
});

/* -------------------------------------------------------------------------- */
/* XML safety                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * Every character in `xml` that XML 1.0's `Char` production forbids.
 *
 * A part containing one of these is not merely ugly - it is not well-formed, so
 * Word declines to open the **whole document**. Nothing may reach a part, from
 * any source: the markdown, the metadata, or a custom property.
 *
 * Note `charCodeAt`, not `codePointAt`: `codePointAt` on a high surrogate
 * returns the combined astral code point, which makes the low half of every
 * emoji look orphaned.
 */
function xmlIllegal(xml: string): string[] {
  const found: string[] = [];
  for (let i = 0; i < xml.length; i += 1) {
    const unit = xml.charCodeAt(i);
    const at = `U+${unit.toString(16).toUpperCase().padStart(4, "0")}@${i}`;
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = xml.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) i += 1;
      else found.push(`lone-high-surrogate ${at}`);
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      found.push(`lone-low-surrogate ${at}`);
    } else if (unit === 0xfffe || unit === 0xffff) {
      found.push(`noncharacter ${at}`);
    } else if (unit < 0x20 && unit !== 0x09 && unit !== 0x0a && unit !== 0x0d) {
      found.push(`control ${at}`);
    }
  }
  return found;
}

describe("XML safety", () => {
  /** C0 controls, unpaired surrogates and the two noncharacters, all in one paste. */
  const HOSTILE =
    "# Title\u0001 here\n\n" +
    "Body with NUL\u0000, backspace\u0008, vtab\u000B, FF\u000C, US\u001F,\n" +
    "a lone high surrogate \uD800, a lone low one \uDC00 and \uFFFE\uFFFF.\n\n" +
    "```ts\nconst x = \u0002;\n```\n\n" +
    "| a\u0003 | b |\n| --- | --- |\n| c\u0004 | d |\n";

  it("never writes a character XML cannot represent into any part", async () => {
    const bytes = await convert(HOSTILE, {
      metadata: {
        title: "T\u0001",
        author: "A\u000B",
        subject: "S\uFFFF",
        description: "D\uD800",
        keywords: ["k\u0007"],
        custom: { "key\u0001": "value\u0002" },
      },
    });

    for (const part of await listDocxParts(bytes)) {
      if (part.endsWith("/") || part.startsWith("word/media/")) continue;
      expect({ part, illegal: xmlIllegal(await readDocxPart(bytes, part)) }).toEqual({
        part,
        illegal: [],
      });
    }
  });

  it("substitutes U+FFFD and says so, once, rather than dropping the text", async () => {
    const warnings = await warningsOf("Body with a\u0001b and c\u000Bd.");

    expect(warnings).toEqual([
      {
        stage: "parse",
        code: "unrepresentable-character",
        severity: "notice",
        message: "replaced 2 characters that XML cannot represent with U+FFFD",
        line: null,
      },
    ]);

    const xml = await partOf("Body with a\u0001b and c\u000Bd.", "word/document.xml");
    expect(xml).toContain("a�b");
    expect(xml).toContain("c�d");
  });

  it("keeps every character XML *can* represent, astral ones included", async () => {
    const xml = await partOf(
      "Emoji \u{1F600}, CJK 中文, a tab\tand DEL\u007F.",
      "word/document.xml",
    );

    expect(xml).toContain("\u{1F600}");
    expect(xml).toContain("中文");
    expect(xmlIllegal(xml)).toEqual([]);
  });

  it("says nothing when there is nothing to say", async () => {
    expect(await warningsOf("# Clean\n\nJust prose.\n")).toEqual([]);
  });

  it("sanitises metadata, which never passes through the parser", async () => {
    const core = await partOf("# hi", "docProps/core.xml", {
      metadata: { title: "Ti\u0001tle", author: "Au\u000Bthor", subject: "Sub\uFFFEject" },
    });

    expect(core).toContain("Ti�tle");
    expect(core).toContain("Au�thor");
    expect(core).toContain("Sub�ject");
    expect(xmlIllegal(core)).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* The snapshot helper                                                         */
/* -------------------------------------------------------------------------- */

describe("normalizeOoxml", () => {
  it("strips revision save ids", () => {
    const input = '<w:p w:rsidR="00A12B34" w:rsidRDefault="00A12B34"><w:r/></w:p>';
    expect(normalizeOoxml(input)).toBe("<w:p>\n<w:r/>\n</w:p>\n");
  });

  it("stabilises numeric ids and timestamps", () => {
    const input =
      '<w:bookmarkStart w:id="7"/><dcterms:created>2026-08-02T10:11:12Z</dcterms:created>';
    const out = normalizeOoxml(input);

    expect(out).not.toContain('w:id="7"');
    expect(out).not.toContain("2026-08-02");
    expect(out).toContain("__NORMALIZED__");
  });

  it("is idempotent", () => {
    const input = '<w:p w:rsidR="00A12B34"><w:bookmarkStart w:id="3"/></w:p>';
    const once = normalizeOoxml(input);
    expect(normalizeOoxml(once)).toBe(once);
  });

  it("documents every scrubber", () => {
    expect(describeScrubbers().length).toBeGreaterThan(0);
    for (const description of describeScrubbers()) {
      expect(description).not.toHaveLength(0);
    }
  });
});
