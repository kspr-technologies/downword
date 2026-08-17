/**
 * Phase 2 cross-integration: every feature in **one** document, under every theme.
 *
 * The three Phase 2 workstreams (footnotes/TOC, maths, mermaid/themes) were
 * built concurrently against the same renderer without seeing each other's
 * code. Each is well covered in isolation; nothing covered the seams *between*
 * them. This file drives one document that exercises all of them at once and
 * asserts what can only break in combination:
 *
 * - the id spaces (bookmarks, footnotes, numbering instances, style ids,
 *   drawing ids, relationships) stay collision-free while three features
 *   allocate into them at once;
 * - **no visible word of the input is lost** — the check that catches a node
 *   silently dropped by a pass that did not know about another pass's output;
 * - a theme swaps `styles.xml` and nothing structural in `document.xml`.
 *
 * The pipeline is driven stage by stage rather than through `convertWithMath`,
 * because mermaid is a document transform that has to run *between* the parse
 * and the render, and no single-call helper composes all three.
 */

import { describe, expect, it } from "vitest";

import { createImageResolver, resolveDocumentImages } from "../src/images/index.js";
import type { ImageRasterizer } from "../src/images/types.js";
import { convertDocumentMath } from "../src/math/document.js";
import { mathMarkdownIt } from "../src/math/markdown-it.js";
import type { DocumentNode } from "../src/model.js";
import { renderMermaid } from "../src/mermaid/index.js";
import type { MermaidRenderer } from "../src/mermaid/types.js";
import { THEMES, type ThemeName } from "../src/options.js";
import { parseMarkdown } from "../src/parse/index.js";
import { renderDocument, type ImageMap } from "../src/render/index.js";
import { packDocument, readDocxPart, writeFixture } from "./helpers/docx.js";
import {
  bookmarkStarts,
  concreteNums,
  footnoteIds,
  footnoteReferenceIds,
  internalAnchors,
  numPr,
  paragraphs,
  styleIds,
  textOf,
} from "./helpers/xml.js";

/* -------------------------------------------------------------------------- */
/* The combined document                                                       */
/* -------------------------------------------------------------------------- */

/** A real 8x4 PNG (solid #2F5496), as a `data:` URI: no network, no filesystem. */
const PNG_DATA_URI =
  "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAEUlEQVR42mPQD5mGFTFQTwIA0B4jIb0iwL8AAAAASUVORK5CYII=";

/** A second, visibly different 4x8 PNG, so the two drawings cannot dedupe into one media part. */
const PNG2_DATA_URI =
  "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAICAIAAADwWiEfAAAAEUlEQVR42mNQZ2BgYGBgYGAAAAUcAB0K1PjnAAAAAElFTkSuQmCC";

function bytesOf(dataUri: string): Uint8Array {
  return Uint8Array.from(Buffer.from(dataUri.slice("data:image/png;base64,".length), "base64"));
}

/**
 * Every Phase 2 feature in one source, laid out so the risky adjacencies are
 * real: the display equation sits between two tables (the table separator and
 * the equation paragraph have to coexist), the footnote bodies mix inline maths
 * with inline code (two run-level features inside a part that is *not*
 * `document.xml`), and the three ordered lists are true siblings — the `1.` /
 * `1)` delimiter change is what stops CommonMark merging them into one list.
 */
function combinedMarkdown(): string {
  return [
    "# Cross Integration Report",
    "",
    "Jump to [Alpha](#alpha-section), [Beta](#beta-section) and [Gamma](#gamma-section).",
    "",
    "## Alpha Section",
    "",
    "Alpha prose with a marker[^one] and a second marker[^two] in one sentence.",
    "",
    "[^one]: Note one pairs inline maths $e^{i\\pi}+1=0$ with inline code `npm install downword`.",
    "[^two]: Note two cites `convertWithMath()` beside $\\sum_{k=1}^{n} k^2$ inline.",
    "",
    "## Beta Section",
    "",
    "| Region | Revenue |",
    "| --- | --- |",
    "| North | 120 |",
    "",
    "$$",
    "\\int_0^1 x^2 \\, dx = \\frac{1}{3}",
    "$$",
    "",
    "| Quarter |",
    "| --- |",
    "| Q3 |",
    "",
    "## Gamma Section",
    "",
    "```mermaid",
    "flowchart LR",
    "  Browser --> Worker --> Word",
    "```",
    "",
    "1. Ordered alpha one",
    "2. Ordered alpha two",
    "",
    "1) Ordered beta one",
    "2) Ordered beta two",
    "",
    "1. Ordered gamma one",
    "2. Ordered gamma two",
    "",
    "- [ ] Task pending review",
    "- [x] Task already finished",
    "",
    `![Chart of quarterly totals](${PNG_DATA_URI})`,
    "",
  ].join("\n");
}

/**
 * Every word the reader must still be able to find in the output.
 *
 * Deliberately *not* derived from the markdown by a stripper: a stripper that
 * shared a bug with the parser would agree with it. This is the list a human
 * reading the source expects to see in Word.
 *
 * The one deliberate omission is the mermaid fence's own source (`flowchart`,
 * `Browser`, `Worker`): when the diagram renders, the picture replaces it by
 * design. `keeps the diagram's source when mermaid cannot run` below pins the
 * other half of that contract — in Node, with no diagram, the text stays.
 */
const EXPECTED_WORDS: readonly string[] = [
  // headings and the intra-document links
  "Cross",
  "Integration",
  "Report",
  "Jump",
  "Alpha",
  "Beta",
  "Gamma",
  "Section",
  // prose carrying the footnote markers
  "prose",
  "marker",
  "sentence",
  // footnote bodies, including the inline code inside them
  "Note",
  "pairs",
  "maths",
  "npm",
  "install",
  "downword",
  "cites",
  "convertWithMath()",
  "beside",
  // both tables, on either side of the display equation
  "Region",
  "Revenue",
  "North",
  "120",
  "Quarter",
  "Q3",
  // the three sibling ordered lists
  "Ordered",
  "alpha",
  "beta",
  "gamma",
  // the task list
  "Task",
  "pending",
  "review",
  "already",
  "finished",
  // the image's alt text, which lives in `wp:docPr`, not in a `w:t`
  "Chart",
  "quarterly",
  "totals",
];

/* -------------------------------------------------------------------------- */
/* Seams: a mermaid renderer and rasteriser that need no DOM                   */
/* -------------------------------------------------------------------------- */

/** A believable flowchart SVG, so the pass sees a real intrinsic size. */
const FAKE_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="480" height="120" viewBox="0 0 480 120">' +
  '<rect width="480" height="120" fill="#ffffff"/><text x="8" y="60">Browser</text></svg>';

/** Stands in for mermaid: no DOM, deterministic bytes. */
const stubRenderer: MermaidRenderer = {
  render: () => Promise.resolve({ svg: FAKE_SVG, width: 480, height: 120 }),
};

/** Stands in for a canvas, returning a PNG that is *not* the document's own image. */
const stubRasterizer: ImageRasterizer = {
  rasterize: (request) =>
    Promise.resolve({
      format: "png" as const,
      data: bytesOf(PNG2_DATA_URI),
      width: request.targetWidth ?? 480,
      height: request.targetHeight ?? 120,
    }),
};

/* -------------------------------------------------------------------------- */
/* The pipeline                                                                */
/* -------------------------------------------------------------------------- */

interface CombinedParts {
  readonly document: string;
  readonly styles: string;
  readonly numbering: string;
  readonly footnotes: string;
  readonly ommlCount: number;
  readonly mermaidRendered: number;
}

/**
 * parse -> mermaid -> maths -> images -> render -> pack -> unzip.
 *
 * The order is the documented one and it matters: mermaid rewrites `codeBlock`
 * nodes into `image` nodes, so it has to run before the image pass resolves
 * them, and maths fills the model's `omml` slots, which the renderer reads.
 */
async function buildCombined(theme: ThemeName): Promise<CombinedParts> {
  const parsed: DocumentNode = parseMarkdown(combinedMarkdown(), { plugins: [mathMarkdownIt()] });

  const mermaid = await renderMermaid(parsed, {
    renderer: stubRenderer,
    rasterizer: stubRasterizer,
  });
  const maths = await convertDocumentMath(mermaid.document);

  const { images }: { images: ImageMap } = await resolveDocumentImages(
    maths.document,
    createImageResolver({ allowRemote: false }),
  );

  const file = renderDocument(maths.document, {
    theme: THEMES[theme],
    images,
    toc: true,
    footnotes: true,
  });
  const bytes = await packDocument(file);

  // Written where `scripts/docx-validity.mjs` looks, so CI's LibreOffice pass
  // actually opens the one document that has every feature in it at once —
  // the case most likely to produce a file Word offers to repair.
  await writeFixture(`golden-cross-${theme}.docx`, bytes);

  return {
    document: await readDocxPart(bytes, "word/document.xml"),
    styles: await readDocxPart(bytes, "word/styles.xml"),
    numbering: await readDocxPart(bytes, "word/numbering.xml"),
    footnotes: await readDocxPart(bytes, "word/footnotes.xml"),
    ommlCount: maths.ommlCount,
    mermaidRendered: mermaid.rendered,
  };
}

const THEME_NAMES: readonly ThemeName[] = ["default", "github", "academic"];

/** Cached: the pipeline is the slow part and every test wants the same output. */
const built = new Map<ThemeName, Promise<CombinedParts>>();
function combined(theme: ThemeName): Promise<CombinedParts> {
  const existing = built.get(theme);
  if (existing !== undefined) return existing;
  const fresh = buildCombined(theme);
  built.set(theme, fresh);
  return fresh;
}

/* -------------------------------------------------------------------------- */
/* Structural helpers                                                          */
/* -------------------------------------------------------------------------- */

/**
 * The `<w:tbl>` / `<w:p>` skeleton of a body, in document order.
 *
 * **Top-level children only.** Every table cell holds a `<w:p>` of its own, so
 * a flat scan counts a table's contents as siblings of the table and reports
 * one extra `P` per cell. What matters here is what sits *between* two `w:tbl`
 * elements, so this tracks table depth and records nothing inside one.
 */
function bodySkeleton(documentXml: string): string[] {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(documentXml)?.[1] ?? "";
  const out: string[] = [];
  let depth = 0;

  for (const match of body.matchAll(/<w:tbl>|<\/w:tbl>|<w:p(?:\s[^>]*)?>|<w:p\/>|<w:sectPr>/g)) {
    const token = match[0];
    if (token === "<w:tbl>") {
      if (depth === 0) out.push("TBL");
      depth += 1;
    } else if (token === "</w:tbl>") {
      depth -= 1;
    } else if (depth === 0) {
      out.push(token === "<w:sectPr>" ? "SECT" : "P");
    }
  }
  return out;
}

/** The ordered sequence of element names in a part — its shape, without values. */
function elementSkeleton(xml: string): string[] {
  return [...xml.matchAll(/<([a-zA-Z][\w:.-]*)(?=[\s/>])/g)].map((m) => m[1] ?? "");
}

/** Every `wp:docPr/@id` in the part, in order. */
function drawingIds(xml: string): string[] {
  return [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1] ?? "");
}

/** Alt text and titles, which live in attributes rather than in `<w:t>`. */
function drawingText(xml: string): string {
  return [...xml.matchAll(/<wp:docPr\s([^/>]*)\/>/g)]
    .flatMap((m) => [...(m[1] ?? "").matchAll(/(?:name|descr|title)="([^"]*)"/g)])
    .map((m) => m[1] ?? "")
    .join(" ");
}

/** `cx`/`cy` of every `wp:extent`, in EMU. */
function extents(xml: string): { cx: number; cy: number }[] {
  return [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => ({
    cx: Number(m[1]),
    cy: Number(m[2]),
  }));
}

/** Duplicated members of a list, as a set. */
function duplicates<T>(values: readonly T[]): T[] {
  const seen = new Set<T>();
  const dupes = new Set<T>();
  for (const value of values) {
    if (seen.has(value)) dupes.add(value);
    seen.add(value);
  }
  return [...dupes];
}

/* -------------------------------------------------------------------------- */
/* 1. Every feature engaged                                                    */
/* -------------------------------------------------------------------------- */

describe("cross-integration: the combined document", () => {
  it.each(THEME_NAMES)("engages every Phase 2 feature under the %s theme", async (theme) => {
    const parts = await combined(theme);

    // Asserted first so that nothing below can pass by virtue of a feature
    // having quietly done nothing at all.
    expect(parts.mermaidRendered).toBe(1);
    expect(parts.ommlCount).toBe(3);
    expect(parts.document).toContain("<m:oMath");
    expect(parts.document).toContain('<w:instrText xml:space="preserve">TOC');
    expect(parts.document).toContain("<w:footnoteReference");
    expect(parts.footnotes).toContain('<w:footnote w:id="1">');
    expect(drawingIds(parts.document)).toHaveLength(2);
  });
});

/* -------------------------------------------------------------------------- */
/* 2. Nothing collides                                                         */
/* -------------------------------------------------------------------------- */

describe("cross-integration: the id spaces stay disjoint", () => {
  it.each(THEME_NAMES)("allocates unique bookmark ids under %s", async (theme) => {
    const { document } = await combined(theme);
    const marks = bookmarkStarts(document);

    // One per heading, and the ids are what pair start with end.
    expect(marks.map((m) => m.name)).toEqual([
      "cross_integration_report",
      "alpha_section",
      "beta_section",
      "gamma_section",
    ]);
    expect(duplicates(marks.map((m) => m.id))).toEqual([]);
    // Every `bookmarkEnd` answers exactly one `bookmarkStart`.
    const ends = [...document.matchAll(/<w:bookmarkEnd w:id="(\d+)"\/>/g)].map((m) => Number(m[1]));
    expect(ends.sort()).toEqual(marks.map((m) => m.id).sort());
  });

  it.each(THEME_NAMES)("numbers footnotes consistently under %s", async (theme) => {
    const { document, footnotes } = await combined(theme);
    const refs = footnoteReferenceIds(document);
    const defs = footnoteIds(footnotes);

    expect(refs).toEqual([1, 2]);
    expect(defs).toEqual([1, 2]);
    expect(duplicates(defs)).toEqual([]);
    // The pairing is what stops Word reporting a corrupt part.
    for (const ref of refs) expect(defs).toContain(ref);
  });

  it.each(THEME_NAMES)("gives each sibling list its own numbering instance under %s", async (t) => {
    const { document, numbering } = await combined(t);
    const nums = concreteNums(numbering);

    expect(duplicates(nums.map((n) => n.numId))).toEqual([]);

    // Three ordered lists, three *different* instances, each restarting at 1 —
    // this is the invariant that stops the second list continuing "3. 4.".
    const listed = paragraphs(document)
      .map((p) => ({ text: textOf(p), num: numPr(p) }))
      .filter((entry) => entry.num !== null && entry.text.startsWith("Ordered "));
    expect(listed).toHaveLength(6);

    const orderedIds = [...new Set(listed.map((entry) => entry.num?.numId))];
    expect(orderedIds).toHaveLength(3);
    for (const numId of orderedIds) {
      const concrete = nums.find((n) => n.numId === numId);
      expect(concrete?.startOverride).toBe(1);
    }
  });

  it.each(THEME_NAMES)("declares every style id exactly once under %s", async (theme) => {
    const { styles } = await combined(theme);
    expect(duplicates(styleIds(styles))).toEqual([]);
  });

  it.each(THEME_NAMES)("gives every drawing a unique wp:docPr id under %s", async (theme) => {
    const { document } = await combined(theme);
    const ids = drawingIds(document);

    // `ST_DrawingElementId` is unique per document. docx 9.7.1 would stamp
    // `id="1"` on both (its generator is built per instance), so this is the
    // renderer's own allocator being exercised across two *different* features:
    // one picture from the mermaid pass, one from the markdown.
    expect(ids).toEqual(["1", "2"]);
    expect(duplicates(ids)).toEqual([]);
  });

  it.each(THEME_NAMES)("points every image relationship at a real part under %s", async (theme) => {
    const { document } = await combined(theme);
    const embeds = [...document.matchAll(/r:embed="(rId\d+)"/g)].map((m) => m[1]);

    // Two visually different pictures must not collapse onto one relationship.
    expect(embeds).toHaveLength(2);
    expect(duplicates(embeds)).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* 3. No text loss                                                             */
/* -------------------------------------------------------------------------- */

describe("cross-integration: nothing is silently dropped", () => {
  it.each(THEME_NAMES)("keeps every visible word of the input under %s", async (theme) => {
    const { document, footnotes } = await combined(theme);

    // The reader's whole text: body runs, footnote runs, and the alt text that
    // only ever appears as a `wp:docPr` attribute.
    const visible = [
      textOf(document),
      textOf(footnotes),
      drawingText(document),
      drawingText(footnotes),
    ].join(" ");

    const missing = EXPECTED_WORDS.filter((word) => !visible.includes(word));
    expect(missing).toEqual([]);
  });

  it("keeps the diagram's source when mermaid cannot run", async () => {
    // The other half of the contract: a fence that did not become a picture is
    // still a code block, so the words survive rather than vanishing.
    // No seams supplied: the real default path, in a real Node process with no
    // DOM, which is where `detectDomSupport` stops before mermaid is loaded.
    const parsed = parseMarkdown(combinedMarkdown(), { plugins: [mathMarkdownIt()] });
    const mermaid = await renderMermaid(parsed);
    const maths = await convertDocumentMath(mermaid.document);
    const { images } = await resolveDocumentImages(
      maths.document,
      createImageResolver({ allowRemote: false }),
    );
    const bytes = await packDocument(
      renderDocument(maths.document, { images, toc: true, footnotes: true }),
    );
    const documentXml = await readDocxPart(bytes, "word/document.xml");

    expect(mermaid.rendered).toBe(0);
    expect(mermaid.warnings.map((w) => w.code)).toEqual(["no-dom"]);
    for (const word of ["flowchart", "Browser", "Worker"]) {
      expect(textOf(documentXml)).toContain(word);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* 4. Themes swap styles, not structure                                        */
/* -------------------------------------------------------------------------- */

describe("cross-integration: a theme only restyles", () => {
  it("keeps document.xml byte-identical between default and github", async () => {
    const [a, b] = await Promise.all([combined("default"), combined("github")]);

    // Neither theme changes a font *size*, so nothing downstream of layout
    // moves and the body should be the very same bytes.
    expect(b.document).toBe(a.document);
    expect(b.numbering).toBe(a.numbering);
  });

  it("keeps the body structure identical between default and academic", async () => {
    const [a, b] = await Promise.all([combined("default"), combined("academic")]);

    // academic sets a 12pt body, which feeds `computeColumnWidths`, so the
    // table's `w:gridCol`/`w:tcW` values legitimately differ. The *shape* must
    // not: same elements, same order, same block skeleton.
    expect(elementSkeleton(b.document)).toEqual(elementSkeleton(a.document));
    expect(bodySkeleton(b.document)).toEqual(bodySkeleton(a.document));
    expect(textOf(b.document)).toBe(textOf(a.document));

    // And the difference really is only those width attributes.
    const scrub = (xml: string): string =>
      xml.replace(/w:w="\d+"/g, 'w:w="N"').replace(/<w:gridCol w:w="\d+"\/>/g, "<w:gridCol/>");
    expect(scrub(b.document)).toBe(scrub(a.document));
  });

  it.each(THEME_NAMES)("changes styles.xml per theme, keeping its skeleton (%s)", async (theme) => {
    const [base, other] = await Promise.all([combined("default"), combined(theme)]);

    // Same styles declared, in the same order — only their values move.
    expect(styleIds(other.styles)).toEqual(styleIds(base.styles));
    if (theme !== "default") expect(other.styles).not.toBe(base.styles);
  });
});

/* -------------------------------------------------------------------------- */
/* 5. Phase 1 invariants, re-checked on the combined document                  */
/* -------------------------------------------------------------------------- */

describe("cross-integration: Phase 1 invariants still hold", () => {
  it.each(THEME_NAMES)("separates the two tables with exactly one paragraph (%s)", async (t) => {
    const { document } = await combined(t);
    const skeleton = bodySkeleton(document);

    // Word merges two `w:tbl` that touch. The display equation sits between
    // these two, so the run is TBL, separator, equation, TBL.
    expect(document).not.toContain("</w:tbl><w:tbl>");
    const firstTable = skeleton.indexOf("TBL");
    expect(skeleton.slice(firstTable, firstTable + 5)).toEqual(["TBL", "P", "P", "TBL", "P"]);
  });

  it.each(THEME_NAMES)("resolves every internal link to a real bookmark (%s)", async (theme) => {
    const { document } = await combined(theme);
    const names = new Set(bookmarkStarts(document).map((m) => m.name));

    const anchors = internalAnchors(document);
    expect(anchors).toEqual(["alpha_section", "beta_section", "gamma_section"]);
    // Resolution is by *name*: `#alpha-section` in markdown, `alpha_section`
    // in OOXML, and the link must land on a bookmark that exists.
    for (const anchor of anchors) expect(names).toContain(anchor);
  });

  it.each(THEME_NAMES)("keeps every picture's aspect ratio exact (%s)", async (theme) => {
    const { document } = await combined(theme);
    const sizes = extents(document);
    expect(sizes).toHaveLength(2);

    // 480x120 (the diagram) and 8x4 (the PNG), both 1 px = 9525 EMU with no
    // whole-pixel re-quantisation on the way.
    expect(sizes[0]).toEqual({ cx: 480 * 9525, cy: 120 * 9525 });
    expect(sizes[1]).toEqual({ cx: 8 * 9525, cy: 4 * 9525 });
    for (const size of sizes) {
      expect(size.cx % 9525).toBe(0);
      expect(size.cy % 9525).toBe(0);
    }
  });

  it.each(THEME_NAMES)("never puts a footnote's body in the document body (%s)", async (theme) => {
    const { document, footnotes } = await combined(theme);

    // The note text belongs in `word/footnotes.xml` and nowhere else; leaking
    // it into the body is how a "footnote" becomes an inline parenthetical.
    expect(textOf(document)).not.toContain("Note one pairs");
    expect(textOf(document)).not.toContain("Note two cites");
    expect(textOf(footnotes)).toContain("Note one pairs");
    expect(textOf(footnotes)).toContain("Note two cites");
  });
});
