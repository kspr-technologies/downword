import { describe, expect, it } from "vitest";

import { convert, DOCX_MIME_TYPE, VERSION } from "../src/index.js";
import { mathPlugin } from "../src/plugins/math.js";
import { mermaidPlugin } from "../src/plugins/mermaid.js";
import { describeScrubbers, normalizeOoxml } from "./helpers/normalize.js";
import { listDocxParts } from "./helpers/docx.js";

/** PK\x03\x04 - the local file header magic every zip (and therefore docx) starts with. */
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];

describe("convert", () => {
  it("produces a zip container with the OOXML mime type", async () => {
    const result = await convert("# Hello\n\nWorld.");

    expect(result.mimeType).toBe(DOCX_MIME_TYPE);
    expect(Array.from(result.bytes.slice(0, 4))).toEqual(ZIP_MAGIC);
    expect(result.bytes.byteLength).toBeGreaterThan(1000);
  });

  it("emits the mandatory OOXML parts", async () => {
    const { bytes } = await convert("hello");
    const parts = await listDocxParts(bytes);

    expect(parts).toContain("[Content_Types].xml");
    expect(parts).toContain("word/document.xml");
    expect(parts).toContain("word/styles.xml");
    expect(parts).toContain("docProps/core.xml");
  });

  it("handles empty input without throwing", async () => {
    const { bytes } = await convert("");
    expect(Array.from(bytes.slice(0, 4))).toEqual(ZIP_MAGIC);
  });

  it("always reports the not-implemented-yet warning (T0 scaffold)", async () => {
    const { warnings } = await convert("hello");
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatch(/renderer is not implemented yet/);
  });

  it("runs plugin setup hooks and collects their warnings", async () => {
    const { warnings } = await convert("hello", {
      plugins: [mathPlugin(), mermaidPlugin({ width: 640 })],
    });

    expect(warnings).toHaveLength(3);
    expect(warnings.some((w) => w.startsWith("downword/plugins/math:"))).toBe(true);
    expect(warnings.some((w) => w.startsWith("downword/plugins/mermaid:"))).toBe(true);
  });

  it("exposes the package version as a build-time constant", () => {
    expect(VERSION).toMatch(/^\d+\.\d+\.\d+/);
  });
});

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
