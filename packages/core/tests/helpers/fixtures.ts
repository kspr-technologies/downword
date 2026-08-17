/**
 * Model documents used by the renderer's tests.
 *
 * Built with the model's own builders rather than parsed from markdown: the
 * renderer's contract is with the *model*, and keeping the parser out of these
 * tests means a renderer regression can never be masked (or caused) by a parser
 * change.
 */

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
  highlight,
  link,
  list,
  listItem,
  metadata,
  orderedList,
  paragraph,
  softBreak,
  strikethrough,
  subscript,
  superscript,
  table,
  tableCell,
  tableRow,
  taskList,
  text,
  thematicBreak,
  type BlockNode,
  type DocumentNode,
  type ResolvedImage,
} from "../../src/model.js";

/* -------------------------------------------------------------------------- */
/* Image bytes                                                                 */
/* -------------------------------------------------------------------------- */

/** A real 8x4 PNG (solid #2F5496). Small, valid, and stable byte-for-byte. */
const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAEUlEQVR42mPQD5mGFTFQTwIA0B4jIb0iwL8AAAAASUVORK5CYII=";

function decodeBase64(value: string): Uint8Array {
  return Uint8Array.from(Buffer.from(value, "base64"));
}

/** An already-resolved raster image, as an image-resolution pass would supply. */
export function resolvedPng(width = 8, height = 4): ResolvedImage {
  return { format: "png", data: decodeBase64(PNG_BASE64), width, height, fallback: null };
}

/**
 * An already-resolved SVG, with the raster twin OOXML requires.
 *
 * Word stores both parts and shows the twin in readers that cannot draw SVG;
 * the model makes `fallback` non-null exactly when `format === "svg"`.
 */
export function resolvedSvg(): ResolvedImage {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="8" height="4"></svg>';
  return {
    format: "svg",
    data: new TextEncoder().encode(svg),
    width: 8,
    height: 4,
    fallback: { format: "png", data: decodeBase64(PNG_BASE64), width: 8, height: 4 },
  };
}

/* -------------------------------------------------------------------------- */
/* Inline coverage                                                             */
/* -------------------------------------------------------------------------- */

/** Every mark, alone and composed, plus breaks and both kinds of link. */
export function inlineBlocks(): readonly BlockNode[] {
  return [
    heading(1, [text("Inline marks")], { id: "inline-marks" }),
    paragraph([
      text("plain "),
      text("bold", [bold()]),
      text(" "),
      text("italic", [italic()]),
      text(" "),
      text("struck", [strikethrough()]),
      text(" "),
      text("code", [inlineCode()]),
      text(" "),
      text("sup", [superscript()]),
      text(" "),
      text("sub", [subscript()]),
      text(" "),
      text("marked", [highlight()]),
    ]),
    paragraph([
      text("composed: "),
      // Deliberately unsorted: the model normalises mark order, so this must
      // render identically to bold-then-italic.
      text("bold italic struck", [italic(), strikethrough(), bold()]),
      text(" and "),
      text("linked bold code", [link("https://example.com", "Example"), bold(), inlineCode()]),
    ]),
    paragraph([
      text("a "),
      // Three adjacent nodes sharing one link: must coalesce into a single
      // <w:hyperlink>, not three.
      text("split ", [link("https://example.com/split")]),
      text("across ", [link("https://example.com/split")]),
      text("runs", [link("https://example.com/split")]),
      text(" then "),
      text("an internal jump", [link("#tables")]),
      text(" and a "),
      text("dangling anchor", [link("#nowhere")]),
    ]),
    paragraph([
      text("hard break here"),
      hardBreak(),
      text("second line"),
      softBreak(),
      text("soft break joined"),
    ]),
    paragraph([text("raw "), htmlInline("<kbd>"), text("Esc"), htmlInline("</kbd>"), text(" tag")]),
  ];
}

/* -------------------------------------------------------------------------- */
/* Heading coverage                                                            */
/* -------------------------------------------------------------------------- */

export function headingBlocks(): readonly BlockNode[] {
  return [
    heading(1, [text("Level one")], { id: "level-one" }),
    heading(2, [text("Level two")], { id: "level-two" }),
    heading(3, [text("Level three")], { id: "level-three" }),
    heading(4, [text("Level four")], { id: "level-four" }),
    heading(5, [text("Level five")], { id: "level-five" }),
    heading(6, [text("Level six")], { id: "level-six" }),
    // Two headings that collapse to the same Word bookmark name once `-` is
    // replaced by `_`; the renderer must keep them distinct.
    heading(2, [text("A-B")], { id: "a-b" }),
    heading(2, [text("A_B")], { id: "a_b" }),
  ];
}

/* -------------------------------------------------------------------------- */
/* List coverage                                                               */
/* -------------------------------------------------------------------------- */

/** Bullets, ordered lists and task lists, all nested three deep. */
export function listBlocks(): readonly BlockNode[] {
  return [
    heading(2, [text("Bullets")], { id: "bullets" }),
    bulletList([
      listItem([
        paragraph([text("level 0")]),
        bulletList([
          listItem([
            paragraph([text("level 1")]),
            bulletList([listItem([paragraph([text("level 2")])])]),
          ]),
        ]),
      ]),
      listItem([
        paragraph([text("level 0, second item")]),
        paragraph([text("a continuation paragraph inside the item")]),
      ]),
    ]),

    heading(2, [text("Ordered")], { id: "ordered" }),
    orderedList([
      listItem([
        paragraph([text("first")]),
        orderedList([
          listItem([
            paragraph([text("nested a")]),
            orderedList([listItem([paragraph([text("nested i")])])]),
          ]),
          listItem([paragraph([text("nested b")])]),
        ]),
      ]),
      listItem([paragraph([text("second")])]),
    ]),

    paragraph([text("A separating paragraph.")]),

    // The classic bug: this must restart at 1 rather than continuing above.
    orderedList([
      listItem([paragraph([text("restarted first")])]),
      listItem([paragraph([text("restarted second")])]),
    ]),

    // An arbitrary start needs its own abstract definition.
    orderedList([listItem([paragraph([text("starts at five")])])], { start: 5 }),

    heading(2, [text("Tasks")], { id: "tasks" }),
    taskList([
      listItem([paragraph([text("done")])], { checked: true }),
      listItem([paragraph([text("not done")])], { checked: false }),
      listItem([paragraph([text("not a checkbox")])]),
      listItem(
        [
          paragraph([text("nested tasks")]),
          taskList([listItem([paragraph([text("deep done")])], { checked: true })]),
        ],
        { checked: false },
      ),
    ]),

    heading(2, [text("Loose")], { id: "loose" }),
    list(
      "bullet",
      [listItem([paragraph([text("loose one")])]), listItem([paragraph([text("loose two")])])],
      {
        tight: false,
      },
    ),

    // A list whose first block is not a paragraph still needs its marker.
    bulletList([listItem([codeBlock("no leading paragraph\n", { lang: "text" })])]),
  ];
}

/* -------------------------------------------------------------------------- */
/* Table coverage                                                              */
/* -------------------------------------------------------------------------- */

export function tableBlocks(): readonly BlockNode[] {
  return [
    heading(2, [text("Tables")], { id: "tables" }),
    table(
      [
        headerRow([
          tableCell([text("Option")]),
          tableCell([text("Centered")]),
          tableCell([text("Numeric")]),
          tableCell([text("Description")]),
        ]),
        tableRow([
          tableCell([text("--fast", [inlineCode()])]),
          tableCell([text("yes", [bold()])]),
          tableCell([text("1.25")]),
          tableCell([
            text(
              "A deliberately long description so the width heuristic has something to weigh, with ",
            ),
            text("emphasis", [italic()]),
            text(" and a "),
            text("link", [link("https://example.com")]),
            text("."),
          ]),
        ]),
        // A ragged row: legal markdown, and Word needs the grid filled.
        tableRow([tableCell([text("--slow", [inlineCode()])]), tableCell([text("no")])]),
      ],
      { align: ["left", "center", "right", "none"] },
    ),
  ];
}

/* -------------------------------------------------------------------------- */
/* Remaining block coverage                                                    */
/* -------------------------------------------------------------------------- */

export function blockBlocks(): readonly BlockNode[] {
  return [
    heading(2, [text("Quotes")], { id: "quotes" }),
    blockquote([
      paragraph([text("first level")]),
      blockquote([
        paragraph([text("second level")]),
        bulletList([listItem([paragraph([text("a list inside a nested quote")])])]),
        codeBlock("quoted();\n", { lang: "js" }),
      ]),
      paragraph([text("back to the first level")]),
    ]),

    heading(2, [text("Code")], { id: "code" }),
    codeBlock('const answer = 42;\n\nconsole.log("hi");\n\twith a tab\n', {
      lang: "ts",
      meta: 'title="demo.ts"',
    }),
    codeBlock("SELECT 1;\n", {
      lang: "sql",
      // Offsets into `value`, as a highlighting pass would produce.
      highlights: [
        { start: 0, end: 6, scope: "keyword" },
        { start: 7, end: 8, scope: "number" },
      ],
    }),

    thematicBreak(),

    heading(2, [text("Images")], { id: "images" }),
    paragraph([image("logo.png", { alt: "The logo", resolved: resolvedPng() })]),
    paragraph([image("diagram.svg", { alt: "A diagram", resolved: resolvedSvg() })]),
    paragraph([image("missing.png", { alt: "Not resolved" })]),
    paragraph([
      text("inline "),
      image("missing-inline.png", { alt: "also missing" }),
      text(" image"),
    ]),

    heading(2, [text("Raw HTML")], { id: "raw-html" }),
    htmlBlock("<details>\n  <summary>Nope</summary>\n</details>\n"),

    heading(2, [text("Footnotes")], { id: "footnotes" }),
    paragraph([
      text("A claim"),
      footnoteReference("1", 1, { label: "1" }),
      text(" and another"),
      footnoteReference("note", 2, { label: "note" }),
      text("."),
    ]),
  ];
}

/** The footnote bodies matching {@link blockBlocks}. */
export function footnoteBlocks(): readonly BlockNode[] {
  return [
    footnoteDefinition("1", 1, [paragraph([text("The first footnote body.")])], { label: "1" }),
    footnoteDefinition(
      "note",
      2,
      [
        paragraph([text("A labelled footnote with "), text("code", [inlineCode()]), text(".")]),
        paragraph([text("And a second paragraph.")]),
      ],
      { label: "note" },
    ),
  ];
}

/* -------------------------------------------------------------------------- */
/* Footnote edge cases                                                         */
/* -------------------------------------------------------------------------- */

/**
 * Every footnote shape that is not "one marker, one paragraph".
 *
 * Deliberately hand-built rather than parsed: `markdown-it-footnote` normalises
 * three of these out of existence before a token is produced — it numbers
 * definitions by *first reference* (so source order is already lost), it drops
 * a definition nothing references, and it leaves a `[^x]` with no definition as
 * literal text. The renderer still has to survive all three, because
 * `renderDocument` is public and a consumer-built model has been through none
 * of that normalisation.
 *
 * Numbers are assigned as markdown-it would: `twice` is cited first, so it is
 * 1, and the definitions below are written in the *other* order on purpose.
 */
export function footnoteEdgeCases(): DocumentNode {
  return doc([
    heading(1, [text("Footnote edge cases")], { id: "footnote-edge-cases" }),
    paragraph([
      text("Cited twice"),
      footnoteReference("twice", 1, { label: "twice" }),
      text(", then rich"),
      footnoteReference("rich", 2, { label: "rich" }),
      text(", then the same one again"),
      footnoteReference("twice", 1, { label: "twice" }),
      text(", then a dangling marker"),
      footnoteReference("missing", 9, { label: "missing" }),
      text("."),
    ]),

    // Out of order: `rich` (2) is defined before `twice` (1).
    footnoteDefinition(
      "rich",
      2,
      [
        paragraph([
          text("A note with "),
          text("code", [inlineCode()]),
          text(", a "),
          text("link", [link("https://example.com")]),
          text(" and a nested marker"),
          footnoteReference("nested", 3, { label: "nested" }),
          text("."),
        ]),
        bulletList([
          listItem([paragraph([text("a list item inside a footnote")])]),
          listItem([paragraph([text("and a second")])]),
        ]),
        codeBlock("inside();\n", { lang: "js" }),
      ],
      { label: "rich" },
    ),
    footnoteDefinition("twice", 1, [paragraph([text("Referenced from two places.")])], {
      label: "twice",
    }),
    footnoteDefinition("nested", 3, [paragraph([text("The inner note.")])], { label: "nested" }),
    footnoteDefinition("orphan", 4, [paragraph([text("Nothing points at me.")])], {
      label: "orphan",
    }),
  ]);
}

/**
 * A document with headings whose slugs collide, plus links to each of them.
 *
 * The slugs are the ones {@link import("../../src/model.js").createSlugger}
 * hands out for three headings all reading "Setup", and the links are written
 * the way markdown-it normalises them — including the percent-encoded one,
 * which is what any anchor containing a non-ASCII letter actually looks like by
 * the time the renderer sees it.
 */
export function anchorDocument(): DocumentNode {
  return doc([
    paragraph([
      text("to the first", [link("#setup")]),
      text(", "),
      text("to the second", [link("#setup-1")]),
      text(", "),
      text("to the third", [link("#setup-2")]),
      text(", "),
      text("to the accented one", [link("#caf%C3%A9")]),
      text(", "),
      text("to nothing at all", [link("#no-such-heading")]),
    ]),
    heading(2, [text("Setup")], { id: "setup" }),
    paragraph([text("first")]),
    heading(2, [text("Setup")], { id: "setup-1" }),
    paragraph([text("second")]),
    heading(2, [text("Setup")], { id: "setup-2" }),
    paragraph([text("third")]),
    heading(2, [text("Café")], { id: "café" }),
    paragraph([text("accented")]),
  ]);
}

/** Six heading levels and some prose, for the table-of-contents cases. */
export function outlineDocument(): DocumentNode {
  return doc(
    [
      heading(1, [text("Chapter one")], { id: "chapter-one" }),
      paragraph([text("Opening prose.")]),
      heading(2, [text("Section A")], { id: "section-a" }),
      paragraph([text("More prose.")]),
      heading(3, [text("Subsection i")], { id: "subsection-i" }),
      heading(4, [text("Too deep for the default range")], { id: "too-deep" }),
      heading(1, [text("Chapter two")], { id: "chapter-two" }),
      paragraph([text("Closing prose.")]),
    ],
    metadata({
      title: "Outline Fixture",
      author: "downword",
      description: "Generated by packages/core/tests/document-features.test.ts",
      keywords: ["toc", "outline"],
      date: "2026-08-17",
      custom: { Fixture: "outline" },
    }),
  );
}

/* -------------------------------------------------------------------------- */
/* Whole documents                                                             */
/* -------------------------------------------------------------------------- */

/** Everything the renderer can emit, in one document. */
export function kitchenSink(): DocumentNode {
  return doc(
    [
      ...inlineBlocks(),
      ...headingBlocks(),
      ...listBlocks(),
      ...tableBlocks(),
      ...blockBlocks(),
      ...footnoteBlocks(),
    ],
    metadata({
      title: "Downword Golden Fixture",
      author: "downword",
      description: "Generated by packages/core/tests/golden-render.test.ts",
      keywords: ["markdown", "docx"],
      date: "2026-08-02",
      custom: { Fixture: "kitchen-sink" },
    }),
  );
}

/** A single paragraph; used to snapshot styles.xml without any body noise. */
export function minimalDocument(): DocumentNode {
  return doc([paragraph([text("Just one paragraph.")])]);
}
