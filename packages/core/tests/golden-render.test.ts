import { describe, expect, it } from "vitest";

import { listDocxParts, renderParts, writeFixture } from "./helpers/docx.js";
import {
  footnoteEdgeCases,
  kitchenSink,
  minimalDocument,
  outlineDocument,
} from "./helpers/fixtures.js";
import { normalizeOoxml } from "./helpers/normalize.js";
import { concreteNums, footnoteIds, footnoteReferenceIds, textOf } from "./helpers/xml.js";

/**
 * Golden tests for the model -> docx renderer.
 *
 * Each case renders a model document, packs it into a real `.docx`, unzips it
 * and snapshots the **normalised** OOXML. Three parts are covered, because a
 * style-driven renderer can only be judged on all three together:
 *
 *  - `word/document.xml` — that the body references styles and numbering.
 *  - `word/styles.xml`   — that those styles actually exist and are correct.
 *  - `word/numbering.xml`— that the list definitions behind them are correct.
 *
 * Snapshots are a *regression* net, not the assertion of record: the positive
 * structural assertions live in `render.test.ts`, because a snapshot will
 * happily immortalise wrong output.
 *
 * The generated files are also written to `tests/__fixtures__/out/`, where CI's
 * LibreOffice gate opens them - so every golden case is additionally checked
 * for real-world openability, not just XML shape.
 */

describe("golden: kitchen sink", () => {
  it("matches the committed word/document.xml", async () => {
    const parts = await renderParts(kitchenSink());
    await writeFixture("golden-render-kitchen-sink.docx", parts.bytes);

    expect(normalizeOoxml(parts.document)).toMatchSnapshot("word/document.xml");
  });

  it("matches the committed word/numbering.xml", async () => {
    const parts = await renderParts(kitchenSink());
    expect(parts.numbering).not.toBeNull();
    expect(normalizeOoxml(parts.numbering ?? "")).toMatchSnapshot("word/numbering.xml");
  });

  it("matches the committed word/footnotes.xml", async () => {
    const parts = await renderParts(kitchenSink());
    expect(parts.footnotes).not.toBeNull();
    expect(normalizeOoxml(parts.footnotes ?? "")).toMatchSnapshot("word/footnotes.xml");
  });

  it("reports exactly the expected warnings", async () => {
    const parts = await renderParts(kitchenSink());
    expect(parts.warnings).toMatchSnapshot("warnings");
  });
});

describe("golden: styles", () => {
  it("matches the committed word/styles.xml", async () => {
    // Rendered from a one-paragraph document: styles.xml must not depend on
    // what the body happens to contain.
    const parts = await renderParts(minimalDocument());
    await writeFixture("golden-render-minimal.docx", parts.bytes);

    expect(normalizeOoxml(parts.styles)).toMatchSnapshot("word/styles.xml");
  });

  it("registers no numbering definitions for a document with no lists", async () => {
    const parts = await renderParts(minimalDocument());

    // docx always writes a numbering part - it carries its own unused
    // default-bullet abstract - so what matters is that we contributed nothing
    // to it: one abstract, one concrete instance, both docx's own.
    expect(concreteNums(parts.numbering ?? "")).toHaveLength(1);
    expect([...(parts.numbering ?? "").matchAll(/<w:abstractNum /g)]).toHaveLength(1);
  });
});

describe("golden: footnotes", () => {
  it("matches the committed word/document.xml", async () => {
    const parts = await renderParts(footnoteEdgeCases());
    await writeFixture("golden-render-footnotes.docx", parts.bytes);

    expect(normalizeOoxml(parts.document)).toMatchSnapshot("word/document.xml");
  });

  it("matches the committed word/footnotes.xml", async () => {
    const parts = await renderParts(footnoteEdgeCases());
    expect(parts.footnotes).not.toBeNull();
    expect(normalizeOoxml(parts.footnotes ?? "")).toMatchSnapshot("word/footnotes.xml");
  });

  it("keeps every reference pointed at a note, and every note out of the body", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    // The acceptance criterion, asserted rather than snapshotted: the snapshot
    // normaliser scrubs `w:id`, so only this can prove the two parts agree.
    const defined = new Set(footnoteIds(parts.footnotes ?? ""));
    expect(defined.size).toBeGreaterThan(0);
    for (const id of footnoteReferenceIds(parts.document)) expect(defined).toContain(id);

    const body = textOf(parts.document);
    for (const note of [
      "Referenced from two places.",
      "A note with",
      "The inner note.",
      "Nothing points at me.",
    ]) {
      expect(body).not.toContain(note);
    }
  });

  it("matches the committed inlined rendering", async () => {
    const parts = await renderParts(footnoteEdgeCases(), { footnotes: false });
    await writeFixture("golden-render-footnotes-inline.docx", parts.bytes);

    expect(footnoteIds(parts.footnotes ?? "")).toEqual([]);
    expect(normalizeOoxml(parts.document)).toMatchSnapshot("word/document.xml (footnotes: false)");
  });

  it("reports exactly the expected warnings", async () => {
    const parts = await renderParts(footnoteEdgeCases());
    expect(parts.warnings).toMatchSnapshot("warnings");
  });
});

describe("golden: document niceties", () => {
  const niceties = {
    toc: { minLevel: 1, maxLevel: 3 },
    pageNumbers: { format: "page-x-of-y", alignment: "center" },
    titleBlock: true,
    subject: "Golden niceties",
  } as const;

  it("matches the committed word/document.xml", async () => {
    const parts = await renderParts(outlineDocument(), niceties);
    await writeFixture("golden-render-document-niceties.docx", parts.bytes);

    expect(normalizeOoxml(parts.document)).toMatchSnapshot("word/document.xml");
  });

  it("matches the committed word/footer1.xml", async () => {
    const parts = await renderParts(outlineDocument(), niceties);
    expect(parts.footer).not.toBeNull();
    expect(normalizeOoxml(parts.footer ?? "")).toMatchSnapshot("word/footer1.xml");
  });

  it("matches the committed word/settings.xml", async () => {
    const parts = await renderParts(outlineDocument(), niceties);
    expect(normalizeOoxml(parts.settings ?? "")).toMatchSnapshot("word/settings.xml");
  });

  it("matches the committed docProps/core.xml", async () => {
    const parts = await renderParts(outlineDocument(), niceties);
    expect(normalizeOoxml(parts.coreProperties ?? "")).toMatchSnapshot("docProps/core.xml");
  });

  it("reports exactly the expected warnings", async () => {
    const parts = await renderParts(outlineDocument(), niceties);
    expect(parts.warnings).toMatchSnapshot("warnings");
  });
});

describe("golden: determinism", () => {
  it("produces identical XML for repeated renders in one process", async () => {
    const first = await renderParts(kitchenSink());
    const second = await renderParts(kitchenSink());

    expect(normalizeOoxml(second.document)).toBe(normalizeOoxml(first.document));
    expect(normalizeOoxml(second.styles)).toBe(normalizeOoxml(first.styles));
    expect(normalizeOoxml(second.numbering ?? "")).toBe(normalizeOoxml(first.numbering ?? ""));
  });

  it("produces an identical package layout for repeated renders", async () => {
    // Not just equal XML: the same set of zip entries, which also covers media
    // parts (docx names them by content hash, so an unstable image byte stream
    // would show up here as a renamed entry).
    const first = await renderParts(kitchenSink());
    const second = await renderParts(kitchenSink());

    expect(await listDocxParts(second.bytes)).toEqual(await listDocxParts(first.bytes));
    expect(normalizeOoxml(second.footnotes ?? "")).toBe(normalizeOoxml(first.footnotes ?? ""));
  });
});
