import { describe, expect, it } from "vitest";

import {
  blockquote,
  bulletList,
  codeBlock,
  doc,
  footnoteDefinition,
  footnoteReference,
  hardBreak,
  headerRow,
  heading,
  htmlBlock,
  htmlInline,
  image,
  inlineCode,
  bold,
  italic,
  link,
  list,
  listItem,
  mathBlock,
  mathInline,
  metadata,
  orderedList,
  paragraph,
  softBreak,
  table,
  tableCell,
  tableRow,
  taskList,
  text,
  thematicBreak,
  type BlockNode,
  type DocumentNode,
} from "../src/model.js";
import {
  buildStyles,
  DEFAULT_THEME,
  headingStyleId,
  prepareHighlights,
  prepareImages,
  renderDocument,
  resolveTheme,
  scopeColor,
  STYLE_IDS,
  type HighlightSpan,
  type RenderWarning,
} from "../src/render/index.js";
import { packDocument, readDocxPart, renderParts } from "./helpers/docx.js";
import { listBlocks, resolvedPng, resolvedSvg } from "./helpers/fixtures.js";
import {
  abstractForNumId,
  cells,
  concreteNums,
  indent,
  numberingLevels,
  numPr,
  paragraphs,
  paragraphWithText,
  pStyle,
  rows,
  rStyles,
  styleById,
  styleIds,
  tables,
  textOf,
} from "./helpers/xml.js";

/**
 * Structural assertions on the renderer's OOXML.
 *
 * These are the tests of record. Every one asserts *positively* on the emitted
 * XML - that a heading carries `<w:pStyle w:val="Heading2"/>`, that a list
 * paragraph joins a specific numbering instance, that a header row carries
 * `<w:tblHeader/>` - rather than diffing a blob. A snapshot cannot tell the
 * difference between "correct" and "consistently wrong"; these can.
 */

function render(blocks: readonly BlockNode[]): Promise<Awaited<ReturnType<typeof renderParts>>> {
  return renderParts(doc(blocks));
}

/**
 * The concrete numberings this renderer created.
 *
 * `numId` 1 is docx's own `default-bullet-numbering`, written into every
 * document whether or not anything references it, so counting instances means
 * excluding it.
 */
function ourConcreteNums(numberingXml: string | null): ReturnType<typeof concreteNums> {
  return concreteNums(numberingXml ?? "").filter((num) => num.numId !== 1);
}

/* -------------------------------------------------------------------------- */
/* styles.xml                                                                  */
/* -------------------------------------------------------------------------- */

describe("styles.xml", () => {
  it("declares every style the renderer references", async () => {
    const { styles } = await render([paragraph([text("x")])]);
    const ids = styleIds(styles);

    for (const required of [
      "Title",
      "Heading1",
      "Heading2",
      "Heading3",
      "Heading4",
      "Heading5",
      "Heading6",
      "Quote",
      "ListParagraph",
      "Hyperlink",
      "CodeChar",
      "CodeBlock",
      "FootnoteReference",
      "FootnoteText",
      "HorizontalRule",
      "TableText",
      "TableHeading",
      "Figure",
      "ImagePlaceholder",
      "HtmlBlock",
      "HtmlChar",
    ]) {
      expect(ids, `styles.xml is missing ${required}`).toContain(required);
    }
  });

  it("declares Quote as a real paragraph style, not a docx default", async () => {
    // The API reference is explicit that `Quote` is NOT a styles.default key;
    // getting this wrong silently drops the style.
    const { styles } = await render([blockquote([paragraph([text("q")])])]);
    const quote = styleById(styles, "Quote");

    expect(quote).toContain('<w:style w:type="paragraph" w:styleId="Quote">');
    expect(quote).toContain('<w:name w:val="Quote"/>');
    expect(quote).toContain('<w:pBdr><w:left w:val="single"');
  });

  it("gives headings the built-in names and outline levels Word expects", async () => {
    const { styles } = await render([heading(2, [text("h")], { id: "h" })]);
    const h2 = styleById(styles, "Heading2");

    // Word writes "heading 2", lowercase; that spelling is how a reader
    // matches the style to its built-in identity when the style set changes.
    expect(h2).toContain('<w:name w:val="heading 2"/>');
    expect(h2).toContain('<w:outlineLvl w:val="1"/>');
    expect(h2).toContain("<w:keepNext/>");
  });

  it("gives CodeBlock shading, a monospace fallback chain and collapsed spacing", async () => {
    const { styles } = await render([codeBlock("x\n")]);
    const style = styleById(styles, "CodeBlock");

    expect(style).toContain('<w:shd w:fill="F6F8FA"');
    // Consolas on ascii/hAnsi, Courier New on cs: OOXML's only fallback chain.
    expect(style).toContain('w:ascii="Consolas"');
    expect(style).toContain('w:hAnsi="Consolas"');
    expect(style).toContain('w:cs="Courier New"');
    // Contextual spacing is what makes consecutive code lines airtight while
    // the block as a whole keeps its outer spacing.
    expect(style).toContain("<w:contextualSpacing/>");
    expect(style).toContain("<w:keepLines/>");
  });

  it("gives CodeChar a monospace face and subtle shading", async () => {
    const { styles } = await render([paragraph([text("x", [inlineCode()])])]);
    const style = styleById(styles, "CodeChar");

    expect(style).toContain('<w:style w:type="character" w:styleId="CodeChar">');
    expect(style).toContain('w:ascii="Consolas"');
    expect(style).toContain("<w:shd");
  });

  it("gives Hyperlink a colour and an underline", async () => {
    const { styles } = await render([paragraph([text("x", [link("https://example.com")])])]);
    expect(styleById(styles, "Hyperlink")).toContain('<w:color w:val="0563C1"/>');
    expect(styleById(styles, "Hyperlink")).toContain('<w:u w:val="single"/>');
  });

  it("threads theme overrides through to the style definitions", async () => {
    const parts = await renderParts(doc([heading(1, [text("h")], { id: "h" })]), {
      theme: { colors: { link: "AA0000" }, headings: { 1: { size: 44 } } },
    });

    expect(styleById(parts.styles, "Hyperlink")).toContain('<w:color w:val="AA0000"/>');
    expect(styleById(parts.styles, "Heading1")).toContain('<w:sz w:val="44"/>');
    // Untouched tokens keep their defaults.
    expect(styleById(parts.styles, "Heading2")).toContain('<w:sz w:val="28"/>');
  });

  it("buildStyles is usable standalone and is idempotent over a resolved theme", () => {
    const fromPartial = buildStyles({ colors: { link: "AA0000" } });
    const fromResolved = buildStyles(resolveTheme({ colors: { link: "AA0000" } }));

    expect(JSON.stringify(fromResolved)).toBe(JSON.stringify(fromPartial));
    expect(JSON.stringify(buildStyles(DEFAULT_THEME))).toBe(JSON.stringify(buildStyles()));
  });
});

/* -------------------------------------------------------------------------- */
/* Headings                                                                    */
/* -------------------------------------------------------------------------- */

describe("headings", () => {
  it("uses the built-in style id rather than direct formatting", async () => {
    const { document } = await render([
      heading(1, [text("One")], { id: "one" }),
      heading(2, [text("Two")], { id: "two" }),
      heading(6, [text("Six")], { id: "six" }),
    ]);

    expect(pStyle(paragraphWithText(document, "One"))).toBe("Heading1");
    expect(pStyle(paragraphWithText(document, "Two"))).toBe("Heading2");
    expect(pStyle(paragraphWithText(document, "Six"))).toBe("Heading6");
  });

  it("puts no direct character formatting on a heading's runs", async () => {
    // The whole point of the style: a user who switches Word's Style Set must
    // get their heading appearance, not ours baked into the run.
    const { document } = await render([heading(2, [text("Two")], { id: "two" })]);
    const heading2 = paragraphWithText(document, "Two");

    expect(heading2).toContain('<w:pStyle w:val="Heading2"/>');
    expect(heading2).not.toContain("<w:sz ");
    expect(heading2).not.toContain("<w:b/>");
    expect(heading2).not.toContain("<w:color ");
    expect(heading2).not.toContain("<w:rFonts ");
  });

  it("bookmarks headings and resolves internal links to them", async () => {
    const { document } = await render([
      paragraph([text("go", [link("#target-heading")])]),
      heading(2, [text("Target")], { id: "target-heading" }),
    ]);

    expect(document).toMatch(/<w:bookmarkStart w:name="target_heading"/);
    expect(document).toContain('<w:hyperlink w:history="1" w:anchor="target_heading">');
  });

  it("keeps colliding bookmark names distinct", async () => {
    const { document } = await render([
      heading(2, [text("A-B")], { id: "a-b" }),
      heading(2, [text("A_B")], { id: "a_b" }),
      paragraph([text("first", [link("#a-b")]), text("second", [link("#a_b")])]),
    ]);

    expect(document).toContain('w:anchor="a_b"');
    expect(document).toContain('w:anchor="a_b_2"');
  });

  it("renders an unresolvable anchor as plain, unstyled text", async () => {
    const parts = await render([paragraph([text("nowhere", [link("#missing")])])]);

    expect(parts.document).not.toContain("<w:hyperlink");
    expect(rStyles(paragraphWithText(parts.document, "nowhere"))).toEqual([]);
    expect(parts.warnings.map((w) => w.code)).toContain("link-unresolved");
  });
});

/* -------------------------------------------------------------------------- */
/* Inline marks                                                                */
/* -------------------------------------------------------------------------- */

describe("inline marks", () => {
  it("maps each mark to its OOXML run property", async () => {
    const { document } = await render([
      paragraph([text("b", [bold()])]),
      paragraph([text("i", [italic()])]),
    ]);

    expect(paragraphWithText(document, "b")).toContain("<w:b/>");
    expect(paragraphWithText(document, "i")).toContain("<w:i/>");
  });

  it("renders composed marks on a single run, independent of mark order", async () => {
    const a = await render([paragraph([text("x", [bold(), italic()])])]);
    const b = await render([paragraph([text("x", [italic(), bold()])])]);

    expect(paragraphWithText(a.document, "x")).toContain("<w:b/><w:bCs/><w:i/><w:iCs/>");
    expect(paragraphWithText(b.document, "x")).toBe(paragraphWithText(a.document, "x"));
  });

  it("uses the CodeChar character style for inline code", async () => {
    const { document } = await render([paragraph([text("npm i", [inlineCode()])])]);
    expect(rStyles(paragraphWithText(document, "npm i"))).toEqual(["CodeChar"]);
  });

  it("coalesces adjacent runs that share a link into one hyperlink", async () => {
    const { document } = await render([
      paragraph([
        text("one ", [link("https://example.com")]),
        text("two", [link("https://example.com")]),
        text(" three", [link("https://example.com/other")]),
      ]),
    ]);

    // Two destinations -> exactly two <w:hyperlink> elements, not three.
    expect([...document.matchAll(/<w:hyperlink /g)]).toHaveLength(2);
  });

  it("keeps CodeChar and re-applies link styling when a link is also code", async () => {
    // A run may carry only one w:rStyle, so the colour and underline that
    // would have come from Hyperlink are applied directly.
    const { document } = await render([
      paragraph([text("npm i", [link("https://example.com"), inlineCode()])]),
    ]);
    const run = paragraphWithText(document, "npm i");

    expect(rStyles(run)).toEqual(["CodeChar"]);
    expect(run).toContain(`<w:color w:val="${DEFAULT_THEME.colors.link}"/>`);
    expect(run).toContain('<w:u w:val="single"/>');
  });

  it("renders a hard break inside the paragraph, not as a new one", async () => {
    const { document } = await render([paragraph([text("a"), hardBreak(), text("b")])]);
    const only = paragraphs(document).filter((p) => textOf(p) === "ab");

    expect(only).toHaveLength(1);
    expect(only[0]).toContain("<w:br/>");
  });

  it("honours the soft-break policy", async () => {
    const block = [paragraph([text("a"), softBreak(), text("b")])];

    const asSpace = await renderParts(doc(block));
    const asBreak = await renderParts(doc(block), { softBreak: "break" });
    const dropped = await renderParts(doc(block), { softBreak: "ignore" });

    expect(textOf(asSpace.document)).toContain("a b");
    expect(asBreak.document).toContain("<w:br/>");
    expect(textOf(dropped.document)).toContain("ab");
  });
});

/* -------------------------------------------------------------------------- */
/* Lists                                                                       */
/* -------------------------------------------------------------------------- */

describe("lists", () => {
  it("styles every list paragraph with ListParagraph and a numbering reference", async () => {
    const { document } = await render(listBlocks());

    for (const p of paragraphs(document)) {
      if (pStyle(p) !== "ListParagraph") continue;
      expect(numPr(p), `list paragraph without numPr: ${textOf(p)}`).not.toBeNull();
    }
  });

  it("never uses docx's shared built-in bullet numbering", async () => {
    // numId 1 is docx's own always-present instance. Every list in the
    // document would share it, so nothing could ever restart.
    const { document, numbering } = await render(listBlocks());
    const used = paragraphs(document)
      .map(numPr)
      .filter((n): n is NonNullable<typeof n> => n !== null)
      .map((n) => n.numId);

    expect(used.length).toBeGreaterThan(0);
    expect(used).not.toContain(1);

    // The bullets in use are ours: docx's built-in abstract declares no font
    // for its glyphs, and cannot, because it is not configurable.
    const bullet = used.find(
      (numId) => numberingLevels(abstractForNumId(numbering ?? "", numId))[0]?.format === "bullet",
    );
    expect(numberingLevels(abstractForNumId(numbering ?? "", bullet ?? -1))[0]?.font).toBe(
      "Symbol",
    );
  });

  it("nests three deep on one numbering instance, one level per depth", async () => {
    const { document } = await render([
      bulletList([
        listItem([
          paragraph([text("l0")]),
          bulletList([
            listItem([paragraph([text("l1")]), bulletList([listItem([paragraph([text("l2")])])])]),
          ]),
        ]),
      ]),
    ]);

    const l0 = numPr(paragraphWithText(document, "l0"));
    const l1 = numPr(paragraphWithText(document, "l1"));
    const l2 = numPr(paragraphWithText(document, "l2"));

    expect(l0?.ilvl).toBe(0);
    expect(l1?.ilvl).toBe(1);
    expect(l2?.ilvl).toBe(2);
    // Same concrete instance throughout: that is what makes level 1 restart
    // under each level-0 item instead of counting globally.
    expect(new Set([l0?.numId, l1?.numId, l2?.numId]).size).toBe(1);
  });

  it("defines nine bullet levels with rotating glyphs and a proper indent ladder", async () => {
    const { document, numbering } = await render([
      bulletList([listItem([paragraph([text("x")])])]),
    ]);
    const numId = numPr(paragraphWithText(document, "x"))?.numId ?? -1;
    const levels = numberingLevels(abstractForNumId(numbering ?? "", numId));

    expect(levels).toHaveLength(9);
    expect(levels.map((l) => l.indentLeft)).toEqual([
      720, 1440, 2160, 2880, 3600, 4320, 5040, 5760, 6480,
    ]);
    expect(levels.every((l) => l.hanging === 360)).toBe(true);
    expect(levels.every((l) => l.format === "bullet")).toBe(true);
    // Word's own three-glyph rotation, with the fonts that make them render.
    expect(levels.slice(0, 4).map((l) => l.font)).toEqual([
      "Symbol",
      "Courier New",
      "Wingdings",
      "Symbol",
    ]);
  });

  it("rotates decimal / lower-letter / lower-roman for ordered levels", async () => {
    const { document, numbering } = await render([
      orderedList([listItem([paragraph([text("x")])])]),
    ]);
    const numId = numPr(paragraphWithText(document, "x"))?.numId ?? -1;
    const levels = numberingLevels(abstractForNumId(numbering ?? "", numId));

    expect(levels.slice(0, 6).map((l) => l.format)).toEqual([
      "decimal",
      "lowerLetter",
      "lowerRoman",
      "decimal",
      "lowerLetter",
      "lowerRoman",
    ]);
    // %N refers to level N-1, so each level numbers itself.
    expect(levels.slice(0, 3).map((l) => l.text)).toEqual(["%1.", "%2.", "%3."]);
  });

  it("restarts sibling ordered lists instead of continuing them", async () => {
    const { document, numbering } = await render([
      orderedList([listItem([paragraph([text("first list")])])]),
      paragraph([text("between")]),
      orderedList([listItem([paragraph([text("second list")])])]),
    ]);

    const a = numPr(paragraphWithText(document, "first list"));
    const b = numPr(paragraphWithText(document, "second list"));

    // Different concrete instances...
    expect(a?.numId).not.toBe(b?.numId);
    const nums = concreteNums(numbering ?? "");
    const first = nums.find((n) => n.numId === a?.numId);
    const second = nums.find((n) => n.numId === b?.numId);

    // ...sharing one abstract definition, each overriding the start back to 1.
    expect(first?.abstractNumId).toBe(second?.abstractNumId);
    expect(first?.startOverride).toBe(1);
    expect(second?.startOverride).toBe(1);
  });

  it("gives sibling bullet lists one shared instance, because nothing counts", async () => {
    // Each concrete numbering costs docx's packer one whole-document string
    // replace (NumberingReplacer), so minting one per list makes packing
    // quadratic in document size. Bullets have no counter to restart, which
    // makes sharing invisible - and this is the assertion that keeps it so.
    const { document, numbering } = await render([
      bulletList([listItem([paragraph([text("first")])])]),
      paragraph([text("between")]),
      bulletList([listItem([paragraph([text("second")])])]),
      taskList([listItem([paragraph([text("task")])], { checked: false })]),
      bulletList([listItem([paragraph([text("third")])])]),
    ]);

    const idOf = (label: string): number | undefined =>
      numPr(paragraphWithText(document, label))?.numId;

    expect(idOf("first")).toBe(idOf("second"));
    expect(idOf("second")).toBe(idOf("third"));
    // A task list keeps its own *reference* - a different glyph - but joins the
    // same instance, so the document still holds one concrete numbering per shape.
    expect(idOf("task")).not.toBe(idOf("first"));
    expect(ourConcreteNums(numbering)).toHaveLength(2);
  });

  it("still gives a bullet list its own instance when an ordered list is inside it", async () => {
    // The nested counter is exactly the thing that has to restart between the
    // two top-level lists, so these subtrees cannot share an instance.
    const nested = (label: string): BlockNode =>
      bulletList([
        listItem([
          paragraph([text(label)]),
          orderedList([listItem([paragraph([text(`${label} one`)])])]),
        ]),
      ]);
    const { document } = await render([nested("a"), paragraph([text("between")]), nested("b")]);

    expect(numPr(paragraphWithText(document, "a"))?.numId).not.toBe(
      numPr(paragraphWithText(document, "b"))?.numId,
    );
    expect(numPr(paragraphWithText(document, "a one"))?.numId).not.toBe(
      numPr(paragraphWithText(document, "b one"))?.numId,
    );
  });

  it("keeps the count of concrete numberings flat as bullet lists multiply", async () => {
    const many = (count: number): BlockNode[] =>
      Array.from({ length: count }, (_unused, index) =>
        bulletList([listItem([paragraph([text(`item ${index}`)])])]),
      );

    const few = await render(many(5));
    const lots = await render(many(200));

    expect(ourConcreteNums(few.numbering)).toHaveLength(1);
    expect(ourConcreteNums(lots.numbering)).toHaveLength(1);
  });

  it("starts an ordered list at an arbitrary number", async () => {
    const { document, numbering } = await render([
      orderedList([listItem([paragraph([text("five")])])], { start: 5 }),
    ]);

    const numId = numPr(paragraphWithText(document, "five"))?.numId ?? -1;
    const concrete = concreteNums(numbering ?? "").find((n) => n.numId === numId);
    const levels = numberingLevels(abstractForNumId(numbering ?? "", numId));

    // startOverride is only ever written for ilvl 0 and reads from the
    // abstract's level 0, so an arbitrary start needs its own abstract.
    expect(concrete?.startOverride).toBe(5);
    expect(levels[0]?.start).toBe(5);
    expect(levels[1]?.start).toBe(1);
  });

  it("gives a nested ordered list its own abstract when it starts at N", async () => {
    const { document, numbering } = await render([
      orderedList([
        listItem([
          paragraph([text("outer")]),
          orderedList([listItem([paragraph([text("inner")])])], { start: 7 }),
        ]),
      ]),
    ]);

    const inner = numPr(paragraphWithText(document, "inner"));
    expect(inner?.ilvl).toBe(1);

    const levels = numberingLevels(abstractForNumId(numbering ?? "", inner?.numId ?? -1));
    // The start has to live on level 1, because the ilvl-0 override cannot
    // reach a nested level.
    expect(levels[1]?.start).toBe(7);
    expect(levels[0]?.start).toBe(1);
  });

  it("renders task items with checkbox glyphs chosen per item", async () => {
    const { document, numbering } = await render([
      taskList([
        listItem([paragraph([text("done")])], { checked: true }),
        listItem([paragraph([text("todo")])], { checked: false }),
      ]),
    ]);

    const done = numPr(paragraphWithText(document, "done"))?.numId ?? -1;
    const todo = numPr(paragraphWithText(document, "todo"))?.numId ?? -1;

    expect(done).not.toBe(todo);
    expect(numberingLevels(abstractForNumId(numbering ?? "", done))[0]?.text).toBe(
      DEFAULT_THEME.taskGlyphs.checked,
    );
    expect(numberingLevels(abstractForNumId(numbering ?? "", todo))[0]?.text).toBe(
      DEFAULT_THEME.taskGlyphs.unchecked,
    );
  });

  it("keeps a tight list tight and a loose list loose", async () => {
    const items = [listItem([paragraph([text("one")])]), listItem([paragraph([text("two")])])];
    const tight = await render([list("bullet", items, { tight: true })]);
    const loose = await render([list("bullet", items, { tight: false })]);

    // Tight lists rely entirely on ListParagraph's contextualSpacing.
    expect(paragraphWithText(tight.document, "one")).not.toContain("<w:spacing");
    expect(paragraphWithText(loose.document, "one")).toContain('<w:spacing w:after="120"/>');
    expect(paragraphWithText(loose.document, "one")).toContain(
      '<w:contextualSpacing w:val="false"/>',
    );
  });

  it("indents continuation blocks to the item's text edge", async () => {
    const { document } = await render([
      bulletList([listItem([paragraph([text("item")]), paragraph([text("continuation")])])]),
    ]);

    expect(indent(paragraphWithText(document, "item"))).toBeNull();
    expect(indent(paragraphWithText(document, "continuation"))).toEqual({ "w:left": "720" });
  });

  it("still emits a marker when an item does not start with a paragraph", async () => {
    const { document } = await render([bulletList([listItem([codeBlock("x\n")])])]);
    const marker = paragraphs(document).find(
      (p) => pStyle(p) === "ListParagraph" && textOf(p) === "",
    );

    expect(marker).toBeDefined();
    expect(numPr(marker ?? "")?.ilvl).toBe(0);
  });
});

/* -------------------------------------------------------------------------- */
/* Tables                                                                      */
/* -------------------------------------------------------------------------- */

describe("tables", () => {
  const sample = (): BlockNode =>
    table(
      [
        headerRow([tableCell([text("Key")]), tableCell([text("Value")])]),
        tableRow([
          tableCell([text("a", [inlineCode()])]),
          tableCell([text("bold", [bold()]), text(" and plain")]),
        ]),
        tableRow([tableCell([text("ragged")])]),
      ],
      { align: ["right", "center"] },
    );

  it("marks the header row so it repeats across pages", async () => {
    const { document } = await render([sample()]);
    const header = rows(tables(document)[0] ?? "")[0] ?? "";

    expect(header).toContain("<w:tblHeader/>");
    expect(header).toContain("<w:cantSplit/>");
    expect(header).toContain(`w:fill="${DEFAULT_THEME.colors.tableHeaderBackground}"`);
  });

  it("applies per-column alignment to the cell paragraphs", async () => {
    const { document } = await render([sample()]);
    const body = rows(tables(document)[0] ?? "")[1] ?? "";
    const [first, second] = cells(body);

    expect(first).toContain('<w:jc w:val="right"/>');
    expect(second).toContain('<w:jc w:val="center"/>');
  });

  it("styles cell paragraphs rather than formatting them inline", async () => {
    const { document } = await render([sample()]);
    const [header, body] = rows(tables(document)[0] ?? "");

    expect(pStyle(cells(header ?? "")[0] ?? "")).toBe("TableHeading");
    expect(pStyle(cells(body ?? "")[0] ?? "")).toBe("TableText");
  });

  it("keeps full inline content inside cells", async () => {
    const { document } = await render([sample()]);
    const body = rows(tables(document)[0] ?? "")[1] ?? "";

    expect(rStyles(cells(body)[0] ?? "")).toEqual(["CodeChar"]);
    expect(cells(body)[1]).toContain("<w:b/>");
    expect(textOf(cells(body)[1] ?? "")).toBe("bold and plain");
  });

  it("emits a fixed grid whose columns sum to the table width", async () => {
    const { document } = await render([sample()]);
    const tableXml = tables(document)[0] ?? "";
    const widths = [...tableXml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((m) => Number(m[1]));
    const declared = Number(/<w:tblW [^>]*w:w="(\d+)"/.exec(tableXml)?.[1]);

    expect(tableXml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(widths).toHaveLength(2);
    expect(widths.reduce((a, b) => a + b, 0)).toBe(declared);
    // A4 minus one-inch margins.
    expect(declared).toBe(9026);
  });

  it("pads ragged rows so every row fills the grid", async () => {
    const { document } = await render([sample()]);
    const all = rows(tables(document)[0] ?? "");

    expect(all).toHaveLength(3);
    for (const row of all) expect(cells(row)).toHaveLength(2);
  });

  it("has visible borders on every edge", async () => {
    const { document } = await render([sample()]);
    const tableXml = tables(document)[0] ?? "";

    for (const edge of ["top", "bottom", "left", "right", "insideH", "insideV"]) {
      expect(tableXml).toContain(`<w:${edge} w:val="single"`);
    }
  });

  it("skips a table with no cells and warns", async () => {
    const parts = await render([table([tableRow([])])]);

    expect(tables(parts.document)).toHaveLength(0);
    expect(parts.warnings.map((w) => w.code)).toContain("table-empty");
  });
});

/* -------------------------------------------------------------------------- */
/* Block quotes                                                                */
/* -------------------------------------------------------------------------- */

describe("block quotes", () => {
  it("uses the Quote style and indents by nesting depth", async () => {
    const { document } = await render([
      blockquote([paragraph([text("outer")]), blockquote([paragraph([text("inner")])])]),
    ]);

    expect(pStyle(paragraphWithText(document, "outer"))).toBe("Quote");
    expect(pStyle(paragraphWithText(document, "inner"))).toBe("Quote");
    expect(indent(paragraphWithText(document, "outer"))).toEqual({ "w:left": "720" });
    expect(indent(paragraphWithText(document, "inner"))).toEqual({ "w:left": "1440" });
  });

  it("takes the quote rule from the style, not from direct formatting", async () => {
    const { document } = await render([blockquote([paragraph([text("q")])])]);
    expect(paragraphWithText(document, "q")).not.toContain("<w:pBdr>");
  });

  it("continues the rule across blocks that carry their own style", async () => {
    // A code block inside a quote cannot use the Quote style, so it gets the
    // left rule directly - otherwise the quote's vertical line breaks.
    const { document } = await render([blockquote([codeBlock("inside();\n")])]);
    const code = paragraphWithText(document, "inside();");

    expect(pStyle(code)).toBe("CodeBlock");
    expect(code).toContain('<w:pBdr><w:left w:val="single"');
    expect(indent(code)).toEqual({ "w:left": "720" });
  });

  it("indents a list inside a quote past both", async () => {
    const { document } = await render([
      blockquote([bulletList([listItem([paragraph([text("bulleted")])])])]),
    ]);

    expect(indent(paragraphWithText(document, "bulleted"))).toEqual({
      "w:left": "1440",
      "w:hanging": "360",
    });
  });
});

/* -------------------------------------------------------------------------- */
/* Code blocks                                                                 */
/* -------------------------------------------------------------------------- */

describe("code blocks", () => {
  it("emits one CodeBlock paragraph per line", async () => {
    const { document } = await render([codeBlock("one\ntwo\n\nfour\n")]);
    const lines = paragraphs(document).filter((p) => pStyle(p) === "CodeBlock");

    // Four lines; the trailing newline terminates, it does not add a fifth.
    expect(lines.map(textOf)).toEqual(["one", "two", "", "four"]);
  });

  it("expands tabs so indentation survives", async () => {
    const { document } = await render([codeBlock("\tindented\n")]);
    expect(textOf(document)).toContain("    indented");

    const narrow = await renderParts(doc([codeBlock("\tindented\n")]), { tabSize: 2 });
    expect(textOf(narrow.document)).toContain("  indented");
  });

  it("preserves leading whitespace with xml:space", async () => {
    const { document } = await render([codeBlock("  two spaces\n")]);
    expect(document).toContain('<w:t xml:space="preserve">  two spaces</w:t>');
  });

  it("colours runs from the model's scope spans", async () => {
    const { document } = await render([
      codeBlock("SELECT 1;\n", {
        lang: "sql",
        highlights: [{ start: 0, end: 6, scope: "keyword" }],
      }),
    ]);

    expect(document).toContain(
      `<w:color w:val="${DEFAULT_THEME.codePalette["keyword"] ?? ""}"/></w:rPr><w:t xml:space="preserve">SELECT</w:t>`,
    );
  });

  it("colours runs from a synchronous highlighter", async () => {
    const parts = await renderParts(doc([codeBlock("let x;\n", { lang: "ts" })]), {
      highlighter: {
        highlight: (code) => [
          { text: "let", color: "#CF222E", bold: true },
          { text: code.slice(3), italic: true },
        ],
      },
    });

    // A leading "#" is a very common adapter mistake; docx needs bare RRGGBB.
    expect(parts.document).toContain('<w:color w:val="CF222E"/>');
    expect(parts.document).toContain("<w:b/>");
    expect(parts.document).toContain("<w:i/>");
  });

  it("warns and falls back when a highlighter is asynchronous", async () => {
    const parts = await renderParts(doc([codeBlock("x\n")]), {
      highlighter: { highlight: () => Promise.resolve([{ text: "x" }]) },
    });

    expect(parts.warnings.map((w) => w.code)).toContain("highlighter-async");
    expect(pStyle(paragraphWithText(parts.document, "x"))).toBe("CodeBlock");
  });

  it("prefers prepared spans over everything else", async () => {
    const block = codeBlock("x\n");
    const spans: HighlightSpan[] = [{ text: "x\n", color: "112233" }];
    const parts = await renderParts(doc([block]), { highlights: new Map([[block, spans]]) });

    expect(parts.document).toContain('<w:color w:val="112233"/>');
  });

  it("never loses source text to a malformed span list", async () => {
    const { document } = await render([
      codeBlock("abcdef\n", {
        highlights: [
          { start: 4, end: 2, scope: "keyword" },
          { start: 0, end: 99, scope: "string" },
        ],
      }),
    ]);

    expect(textOf(document)).toBe("abcdef");
  });
});

/* -------------------------------------------------------------------------- */
/* Thematic break, HTML, images                                                */
/* -------------------------------------------------------------------------- */

describe("other blocks", () => {
  it("renders a thematic break as an empty HorizontalRule paragraph", async () => {
    const { document, styles } = await render([thematicBreak()]);
    const rule = paragraphs(document).find((p) => pStyle(p) === "HorizontalRule");

    expect(rule).toBeDefined();
    expect(textOf(rule ?? "")).toBe("");
    expect(styleById(styles, "HorizontalRule")).toContain("<w:pBdr><w:bottom");
  });

  it("renders raw HTML in monospace under the default policy, and warns", async () => {
    const parts = await render([
      htmlBlock("<details>\n</details>\n"),
      paragraph([text("a "), htmlInline("<br>"), text(" b")]),
    ]);

    expect(paragraphs(parts.document).filter((p) => pStyle(p) === "HtmlBlock")).toHaveLength(2);
    expect(rStyles(paragraphWithText(parts.document, "a <br> b"))).toEqual(["HtmlChar"]);
    expect(parts.warnings.map((w) => w.code)).toEqual(
      expect.arrayContaining(["html-block", "html-inline"]),
    );
  });

  it("drops raw HTML under the drop policy, and still warns", async () => {
    const parts = await renderParts(
      doc([htmlBlock("<details>\n</details>\n"), paragraph([text("a "), htmlInline("<br>")])]),
      { html: "drop" },
    );

    expect(parts.document).not.toContain("HtmlBlock");
    expect(textOf(parts.document)).toBe("a ");
    expect(parts.warnings.map((w) => w.code)).toEqual(
      expect.arrayContaining(["html-block", "html-inline"]),
    );
  });

  it("embeds resolved image bytes at their intrinsic size", async () => {
    const { document } = await render([
      paragraph([image("logo.png", { alt: "Logo", resolved: resolvedPng() })]),
    ]);
    const figure = paragraphs(document).find((p) => pStyle(p) === "Figure") ?? "";

    // 8 x 4 px, and docx multiplies pixels by 9525 to get EMU.
    expect(figure).toContain('<wp:extent cx="76200" cy="38100"/>');
    expect(figure).toContain('descr="Logo"');
  });

  it("scales an oversized image down to the text column", async () => {
    const { document } = await render([
      paragraph([image("wide.png", { alt: "Wide", resolved: resolvedPng(4000, 1000) })]),
    ]);

    // 9026 twips of text column / 15 twips per pixel = 601 px, so 4000x1000
    // scales to 601 x 150.25 px. The quarter pixel is kept rather than rounded
    // away: `<wp:extent>` is in EMU (9525 to the pixel), so `round(150.25 *
    // 9525) = 1431131` is 4.25 EMU from the exact 4:1 ratio, where rounding to
    // 150 px first would have been 2381 EMU out.
    expect(document).toContain('<wp:extent cx="5724525" cy="1431131"/>');

    const extent = /<wp:extent cx="(\d+)" cy="(\d+)"\/>/.exec(document);
    expect(Number(extent?.[1]) / Number(extent?.[2])).toBeCloseTo(4, 5);
  });

  it("emits the SVG plus its raster twin", async () => {
    const { document, bytes } = await render([
      paragraph([image("d.svg", { alt: "Diagram", resolved: resolvedSvg() })]),
    ]);

    expect(document).toContain("asvg:svgBlip");
    expect(bytes.byteLength).toBeGreaterThan(0);
  });

  it("falls back to a placeholder when an image has no bytes", async () => {
    const parts = await render([
      paragraph([image("missing.png", { alt: "Missing" })]),
      paragraph([text("before "), image("also.png", { alt: "Also" })]),
    ]);

    const placeholder = paragraphs(parts.document).find((p) => pStyle(p) === "ImagePlaceholder");
    expect(textOf(placeholder ?? "")).toBe("[image: Missing]");
    expect(rStyles(paragraphWithText(parts.document, "before [image: Also]"))).toEqual([
      "ImagePlaceholderChar",
    ]);
    expect(parts.warnings.filter((w) => w.code === "image-unresolved")).toHaveLength(2);
  });

  it("injects pre-converted OMML for math, and falls back when it is missing", async () => {
    const omml =
      '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
      "<m:r><m:t>x+1</m:t></m:r></m:oMath>";

    // ImportedXmlComponent.fromXmlString cannot be used as-is: it wraps the
    // xml-js *document* node and the packer emits a literal <undefined> tag.
    const converted = await render([mathBlock("x+1", { omml })]);
    expect(converted.document).toContain("<m:oMath");
    expect(converted.document).not.toContain("<undefined");
    expect(converted.warnings).toHaveLength(0);

    const raw = await render([paragraph([mathInline("x+1")])]);
    expect(rStyles(paragraphWithText(raw.document, "x+1"))).toEqual(["CodeChar"]);
    expect(raw.warnings.map((w) => w.code)).toEqual(["math-unconverted"]);
  });

  it("takes image bytes from a prepared map", async () => {
    const node = image("late.png", { alt: "Late" });
    const parts = await renderParts(doc([paragraph([node])]), {
      images: new Map([[node, resolvedPng()]]),
    });

    expect(parts.document).toContain("<w:drawing>");
    expect(parts.warnings).toHaveLength(0);
  });
});

/* -------------------------------------------------------------------------- */
/* Footnotes                                                                   */
/* -------------------------------------------------------------------------- */

describe("footnotes", () => {
  const withFootnote = (): DocumentNode =>
    doc([
      paragraph([text("claim"), footnoteReference("1", 1, { label: "1" })]),
      footnoteDefinition("1", 1, [paragraph([text("the note")])], { label: "1" }),
    ]);

  it("hoists definitions out of the body into word/footnotes.xml", async () => {
    const parts = await renderParts(withFootnote());

    expect(textOf(parts.document)).toBe("claim");
    expect(parts.footnotes).toContain("the note");
    expect(parts.document).toContain('<w:rStyle w:val="FootnoteReference"/>');
    expect(parts.document).toMatch(/<w:footnoteReference w:id="1"\/>/);
  });

  it("styles footnote bodies with FootnoteText", async () => {
    const parts = await renderParts(withFootnote());
    expect(parts.footnotes).toContain('<w:pStyle w:val="FootnoteText"/>');
  });

  it("drops content a footnote part cannot hold, and warns", async () => {
    const parts = await renderParts(
      doc([
        paragraph([footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition("1", 1, [table([tableRow([tableCell([text("nope")])])])], {
          label: "1",
        }),
      ]),
    );

    expect(parts.warnings.map((w) => w.code)).toContain("footnote-content-dropped");
    expect(parts.footnotes).not.toContain("nope");
  });
});

/* -------------------------------------------------------------------------- */
/* Document-level options                                                      */
/* -------------------------------------------------------------------------- */

describe("document options", () => {
  it("writes metadata to the OOXML core and custom properties", async () => {
    const bytes = await packDocument(
      renderDocument(
        doc(
          [paragraph([text("x")])],
          metadata({
            title: "T",
            author: "A",
            description: "D",
            keywords: ["k1", "k2"],
            date: "2026-08-02",
            custom: { Team: "KSPR" },
          }),
        ),
      ),
    );

    const core = await readDocxPart(bytes, "docProps/core.xml");
    expect(core).toContain("<dc:title>T</dc:title>");
    expect(core).toContain("<dc:creator>A</dc:creator>");
    expect(core).toContain("<dc:description>D</dc:description>");
    expect(core).toContain("<cp:keywords>k1, k2</cp:keywords>");

    // OOXML core properties have no authored-date field, so the frontmatter
    // date becomes a custom property rather than being dropped.
    const custom = await readDocxPart(bytes, "docProps/custom.xml");
    expect(custom).toContain('name="Team"><vt:lpwstr>KSPR</vt:lpwstr>');
    expect(custom).toContain('name="Date"><vt:lpwstr>2026-08-02</vt:lpwstr>');
  });

  it("emits a Title-styled block for a frontmatter title, and can be told not to", async () => {
    const source = doc([paragraph([text("body")])], metadata({ title: "My Title" }));

    const withTitle = await renderParts(source);
    const without = await renderParts(source, { titleBlock: false });

    expect(pStyle(paragraphWithText(withTitle.document, "My Title"))).toBe("Title");
    expect(textOf(without.document)).toBe("body");
  });

  it("honours page size and margins", async () => {
    const parts = await renderParts(
      doc([table([headerRow([tableCell([text("a")]), tableCell([text("b")])])])]),
      { page: { size: "Letter", margin: { left: 720, right: 720 } } },
    );

    expect(parts.document).toContain('<w:pgSz w:w="12240" w:h="15840"');
    // 12240 - 720 - 720 = 10800 twips of text column.
    expect(parts.document).toMatch(/<w:tblW w:type="dxa" w:w="10800"\/>/);
  });

  it("widens the text column in landscape", async () => {
    const parts = await renderParts(
      doc([table([headerRow([tableCell([text("a")]), tableCell([text("b")])])])]),
      { page: { orientation: "landscape" } },
    );

    // docx swaps w/h itself, so the usable width is the long edge.
    expect(parts.document).toContain('w:orient="landscape"');
    expect(parts.document).toMatch(/<w:tblW w:type="dxa" w:w="13958"\/>/);
  });

  it("produces a valid body for an empty document", async () => {
    const parts = await renderParts(doc([]));
    expect(paragraphs(parts.document).length).toBeGreaterThanOrEqual(1);
  });

  it("reports every warning through onWarning", async () => {
    const seen: RenderWarning[] = [];
    renderDocument(doc([paragraph([image("missing.png")])]), {
      onWarning: (warning) => seen.push(warning),
    });

    expect(seen.map((w) => w.code)).toEqual(["image-unresolved"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Extension points                                                            */
/* -------------------------------------------------------------------------- */

describe("prepare helpers", () => {
  it("prepareHighlights maps every code block, in document order", async () => {
    const a = codeBlock("a\n", { lang: "ts" });
    const b = codeBlock("b\n");
    const seen: (string | null)[] = [];

    const map = await prepareHighlights(doc([a, b]), {
      highlight: (code, lang) => {
        seen.push(lang);
        return Promise.resolve([{ text: code }]);
      },
    });

    expect(seen).toEqual(["ts", null]);
    expect(map.get(a)?.[0]?.text).toBe("a\n");
    expect(map.get(b)?.[0]?.text).toBe("b\n");
  });

  it("prepareHighlights survives a throwing highlighter", async () => {
    const block = codeBlock("a\n");
    const map = await prepareHighlights(doc([block]), {
      highlight: () => {
        throw new Error("boom");
      },
    });

    expect(map.size).toBe(0);
  });

  it("prepareImages resolves each source once and skips already-resolved nodes", async () => {
    const first = image("same.png");
    const second = image("same.png");
    const already = image("done.png", { resolved: resolvedPng() });
    const calls: string[] = [];

    const map = await prepareImages(doc([paragraph([first, second, already])]), {
      resolve: (src) => {
        calls.push(src);
        return Promise.resolve(resolvedPng());
      },
    });

    expect(calls).toEqual(["same.png"]);
    expect(map.get(first)).toBeDefined();
    expect(map.get(second)).toBeDefined();
    expect(map.get(already)).toBeUndefined();
  });

  it("prepareImages turns a rejecting resolver into a placeholder", async () => {
    const node = image("bad.png", { alt: "Bad" });
    const map = await prepareImages(doc([paragraph([node])]), {
      resolve: () => Promise.reject(new Error("404")),
    });

    expect(map.size).toBe(0);

    const parts = await renderParts(doc([paragraph([node])]), { images: map });
    expect(parts.warnings.map((w) => w.code)).toEqual(["image-unresolved"]);
  });
});

describe("theme", () => {
  it("resolveTheme is idempotent", () => {
    const once = resolveTheme({ colors: { link: "AA0000" } });
    expect(resolveTheme(once)).toEqual(once);
  });

  it("scopeColor walks dotted scopes back to their prefix", () => {
    const theme = resolveTheme();
    expect(scopeColor(theme, "title.function")).toBe(theme.codePalette["title.function"]);
    expect(scopeColor(theme, "title.class.inherited")).toBe(theme.codePalette["title.class"]);
    expect(scopeColor(theme, "totally.unknown")).toBeNull();
  });

  it("exposes the style ids it emits", () => {
    expect(STYLE_IDS.quote).toBe("Quote");
    expect(STYLE_IDS.heading[1]).toBe("Heading2");
    expect(headingStyleId(3)).toBe("Heading3");
  });
});
