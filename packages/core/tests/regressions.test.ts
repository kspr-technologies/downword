import { readFile } from "node:fs/promises";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { createHighlighter, PRINT_CODE_PALETTE } from "../src/highlight/index.js";
import {
  blockquote,
  bulletList,
  codeBlock,
  doc,
  headerRow,
  heading,
  htmlBlock,
  htmlInline,
  image,
  link,
  listItem,
  mathBlock,
  mathInline,
  metadata,
  nodeText,
  orderedList,
  paragraph,
  subscript,
  superscript,
  table,
  tableCell,
  text,
  type BlockNode,
  type ListNode,
  type TableNode,
} from "../src/model.js";
import {
  convert,
  isDownwordError,
  THEMES,
  type ConvertOptions,
  type ConvertWarning,
  type DownwordError,
  type ThemeName,
} from "../src/index.js";
import {
  MAX_BLOCK_NESTING,
  parseMarkdown,
  PARSE_WARNING_CODES,
  parseWarningSeverity,
  type ParseWarning,
} from "../src/parse/index.js";
import {
  DEFAULT_THEME,
  MAX_LIST_LEVEL,
  prepareHighlights,
  prepareImages,
  RENDER_WARNING_CODES,
  renderDocument,
  renderWarningSeverity,
  resolveTheme,
  scopeColor,
  type Highlighter,
  type HighlightSpan,
  type RenderWarning,
} from "../src/render/index.js";
import { contrast, WCAG_AA, WCAG_AAA } from "./helpers/contrast.js";
import { packDocument, readDocxPart, renderParts } from "./helpers/docx.js";
import { kitchenSink, minimalDocument, resolvedPng } from "./helpers/fixtures.js";
import {
  abstractForNumId,
  concreteNums,
  indent,
  numberingLevels,
  numPr,
  paragraphWithText,
  paragraphs,
  pStyle,
  styleById,
  styleIds,
  textOf,
} from "./helpers/xml.js";

/**
 * Regression tests for the two Phase 1 code reviews.
 *
 * One `describe` per finding, each named with the finding number so a failure
 * points straight at the report it came from. The two reviews numbered
 * independently and their numbers overlap, so the second review's findings live
 * below a banner and are prefixed `review 2`. These are deliberately kept
 * together rather than scattered into the domain suites: every one of them
 * reproduces a *specific defect that shipped*, and keeping the reproduction
 * next to the claim it falsifies is what makes the fix auditable later.
 *
 * Everything here asserts on the real OOXML (or the real warning stream), not
 * on internal state - a numbering bug is only a bug because of what Word does
 * with `word/numbering.xml`.
 */

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

function render(blocks: readonly BlockNode[]): Promise<Awaited<ReturnType<typeof renderParts>>> {
  return renderParts(doc(blocks));
}

/** Two-item ordered list whose items are labelled `${label} 1` / `${label} 2`. */
function counted(label: string): ListNode {
  return orderedList([
    listItem([paragraph([text(`${label} 1`)])]),
    listItem([paragraph([text(`${label} 2`)])]),
  ]);
}

/** The concrete `numId` of the paragraph reading exactly `label`. */
function numIdOf(documentXml: string, label: string): number {
  const found = numPr(paragraphWithText(documentXml, label));
  if (found === null) throw new Error(`paragraph ${JSON.stringify(label)} carries no <w:numPr>`);
  return found.numId;
}

/**
 * Asserts that two ordered lists are genuinely independent counters.
 *
 * Three things have to hold together for Word to print `1. 2.` twice rather
 * than `1. 2. 3. 4.`: the paragraphs must join different concrete numberings,
 * each of those must exist in `numbering.xml`, and the level they sit on must
 * declare `start=1` (with the `ilvl 0` override docx writes for every instance
 * it mints).
 */
function expectRestart(
  parts: { document: string; numbering: string | null },
  first: string,
  second: string,
): void {
  const a = numPr(paragraphWithText(parts.document, first));
  const b = numPr(paragraphWithText(parts.document, second));
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect(a?.ilvl).toBe(b?.ilvl);
  expect(a?.numId, `${first} and ${second} share a numbering instance`).not.toBe(b?.numId);

  const nums = concreteNums(parts.numbering ?? "");
  for (const label of [first, second]) {
    const numId = numIdOf(parts.document, label);
    const concrete = nums.find((num) => num.numId === numId);
    expect(concrete, `no <w:num w:numId="${numId}"> for ${label}`).toBeDefined();
    // docx writes the ilvl-0 override for every instance it mints; that is what
    // makes a *top-level* sibling restart.
    expect(concrete?.startOverride).toBe(1);
    // ...and the level actually used must itself start at 1, which is what
    // makes a *nested* sibling restart.
    const levels = numberingLevels(abstractForNumId(parts.numbering ?? "", numId));
    expect(levels[numPr(paragraphWithText(parts.document, label))?.ilvl ?? 0]?.start).toBe(1);
  }
}

/** A bullet list nested `depth` levels deep, innermost item labelled `leaf`. */
function nestedBullets(depth: number, leaf: string): ListNode {
  let inner: ListNode = bulletList([listItem([paragraph([text(leaf)])])]);
  for (let level = depth - 1; level > 0; level -= 1) {
    inner = bulletList([listItem([paragraph([text(`level ${level}`)]), inner])]);
  }
  return inner;
}

/* -------------------------------------------------------------------------- */
/* Finding 2 - sibling ordered lists under a non-ordered parent never restart  */
/* -------------------------------------------------------------------------- */

describe("finding 2: sibling ordered lists restart", () => {
  it("restarts two ordered lists nested under a bullet parent", async () => {
    // The report's exact input:
    //   - a
    //     1. one
    //     2. two
    //   - b
    //     1. one
    //     2. two
    const parts = await render([
      bulletList([
        listItem([paragraph([text("a")]), counted("a")]),
        listItem([paragraph([text("b")]), counted("b")]),
      ]),
    ]);

    expect(numPr(paragraphWithText(parts.document, "a 1"))?.ilvl).toBe(1);
    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts two ordered lists nested under a task parent", async () => {
    const parts = await render([
      bulletList([
        listItem([paragraph([text("a")]), counted("a")], { checked: true }),
        listItem([paragraph([text("b")]), counted("b")], { checked: false }),
      ]),
    ]);

    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts two ordered lists inside a block quote", async () => {
    const parts = await render([
      blockquote([counted("a"), paragraph([text("mid")]), counted("b")]),
    ]);

    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts two ordered lists nested under bullets inside a block quote", async () => {
    const parts = await render([
      blockquote([
        bulletList([
          listItem([paragraph([text("a")]), counted("a")]),
          listItem([paragraph([text("b")]), counted("b")]),
        ]),
      ]),
    ]);

    expectRestart(parts, "a 1", "b 1");
  });

  it.each([
    ["a paragraph", paragraph([text("between")])],
    ["a heading", heading(2, [text("between")], { id: "between" })],
    ["a code block", codeBlock("between\n")],
  ])("restarts two top-level ordered lists separated by %s", async (_what, separator) => {
    const parts = await render([counted("a"), separator, counted("b")]);

    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts two ordered lists that follow each other with nothing between", async () => {
    const parts = await render([counted("a"), counted("b")]);

    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts sibling ordered lists inside one bullet item", async () => {
    // Reachable from markdown by changing the delimiter (`1.` then `1)`), which
    // starts a second list rather than continuing the first.
    const parts = await render([
      bulletList([listItem([paragraph([text("item")]), counted("a"), counted("b")])]),
    ]);

    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts sibling ordered lists inside one ordered item", async () => {
    const parts = await render([
      orderedList([listItem([paragraph([text("item")]), counted("a"), counted("b")])]),
    ]);

    expectRestart(parts, "a 1", "b 1");
  });

  it("restarts nested ordered lists that start at N", async () => {
    const started = (label: string): ListNode =>
      orderedList([listItem([paragraph([text(label)])])], { start: 5 });
    const parts = await render([
      bulletList([
        listItem([paragraph([text("a")]), started("a 1")]),
        listItem([paragraph([text("b")]), started("b 1")]),
      ]),
    ]);

    const a = numIdOf(parts.document, "a 1");
    const b = numIdOf(parts.document, "b 1");
    expect(a).not.toBe(b);
    // Both point at the one abstract whose *level 1* declares start=5.
    for (const numId of [a, b]) {
      expect(numberingLevels(abstractForNumId(parts.numbering ?? "", numId))[1]?.start).toBe(5);
    }
  });

  it("still shares one instance down a pure ordered nest", async () => {
    // The correct case, which the fix must not break: the enclosing ilvl-0
    // paragraphs are what restart level 1 under each item.
    const parts = await render([
      orderedList([
        listItem([paragraph([text("one")]), counted("a")]),
        listItem([paragraph([text("two")]), counted("b")]),
      ]),
    ]);

    expect(numIdOf(parts.document, "one")).toBe(numIdOf(parts.document, "a 1"));
    expect(numIdOf(parts.document, "a 1")).toBe(numIdOf(parts.document, "b 1"));
  });

  it("still shares one instance across bullet-only documents", async () => {
    // The performance property: bullets have no counter, so they must not mint
    // a concrete numbering each.
    const parts = await render(
      Array.from({ length: 50 }, (_unused, index) =>
        bulletList([listItem([paragraph([text(`bullet ${index}`)])])]),
      ),
    );

    expect(concreteNums(parts.numbering ?? "").filter((num) => num.numId !== 1)).toHaveLength(1);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 3 - `[x]` markers stripped from ordered list items                  */
/* -------------------------------------------------------------------------- */

describe("finding 3: checkboxes on ordered list items", () => {
  it("renders the checkbox glyph inside a numbered item", async () => {
    const parts = await render([
      orderedList([
        listItem([paragraph([text("done")])], { checked: true }),
        listItem([paragraph([text("todo")])], { checked: false }),
        listItem([paragraph([text("plain")])]),
      ]),
    ]);

    const texts = paragraphs(parts.document)
      .filter((p) => pStyle(p) === "ListParagraph")
      .map(textOf);

    expect(texts).toEqual([
      `${DEFAULT_THEME.taskGlyphs.checked} done`,
      `${DEFAULT_THEME.taskGlyphs.unchecked} todo`,
      "plain",
    ]);
  });

  it("keeps the item numbered, in the symbol font, and warns about nothing", async () => {
    const parts = await render([
      orderedList([listItem([paragraph([text("done")])], { checked: true })]),
    ]);

    const marker = paragraphWithText(parts.document, `${DEFAULT_THEME.taskGlyphs.checked} done`);
    const numId = numPr(marker)?.numId ?? -1;
    expect(numberingLevels(abstractForNumId(parts.numbering ?? "", numId))[0]?.format).toBe(
      "decimal",
    );
    expect(marker).toContain(`w:ascii="${DEFAULT_THEME.fonts.symbol}"`);
    expect(parts.warnings).toEqual([]);
  });

  it("puts the glyph on an item whose first block is not a paragraph", async () => {
    const parts = await render([orderedList([listItem([codeBlock("x\n")], { checked: false })])]);

    expect(
      paragraphs(parts.document).some(
        (p) => pStyle(p) === "ListParagraph" && textOf(p) === DEFAULT_THEME.taskGlyphs.unchecked,
      ),
    ).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 4 - nesting past nine levels collapses silently                     */
/* -------------------------------------------------------------------------- */

describe("finding 4: list nesting past nine levels", () => {
  it("warns once per level it had to truncate", async () => {
    const parts = await render([nestedBullets(11, "deepest")]);

    const truncations = parts.warnings.filter((warning) => warning.code === "list-depth-truncated");
    expect(truncations).toHaveLength(2);
    expect(truncations[0]?.severity).toBe("notice");
    expect(truncations[0]?.message).toContain("9");
  });

  it("keeps indenting past the ninth level even though ilvl cannot", async () => {
    const parts = await render([nestedBullets(11, "deepest")]);

    const levels = paragraphs(parts.document)
      .filter((p) => pStyle(p) === "ListParagraph")
      .map((p) => ({ ilvl: numPr(p)?.ilvl ?? -1, left: indent(p)?.["w:left"] ?? null }));

    // OOXML has nine levels and no more, so the last three all sit on ilvl 8...
    expect(levels.map((l) => l.ilvl)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 8]);
    // ...but the indent ladder continues, so the nesting is still visible.
    expect(levels.slice(9).map((l) => l.left)).toEqual(["7200", "7920"]);
    // The nine levels the numbering definition can express carry no direct
    // indent at all: it already puts level 9 at 6480.
    expect(levels.slice(0, 9).every((l) => l.left === null)).toBe(true);
  });

  it("says nothing for a list that fits", async () => {
    const parts = await render([nestedBullets(MAX_LIST_LEVEL + 1, "deepest")]);

    expect(parts.warnings).toEqual([]);
    expect(numPr(paragraphWithText(parts.document, "deepest"))?.ilvl).toBe(MAX_LIST_LEVEL);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 5 - `theme` and `highlighter` do not compose                        */
/* -------------------------------------------------------------------------- */

/** `<w:color>` of the first `CodeBlock` run reading exactly `text`. */
function codeInk(documentXml: string, text: string): string | null {
  for (const p of paragraphs(documentXml)) {
    if (pStyle(p) !== "CodeBlock") continue;
    for (const run of p.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)) {
      const body = run[1] ?? "";
      const runText = [...body.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
        .map((t) => t[1] ?? "")
        .join("");
      if (runText !== text) continue;
      const props = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(body)?.[1] ?? "";
      return /<w:color w:val="([^"]*)"\/>/.exec(props)?.[1] ?? null;
    }
  }
  throw new Error(`no CodeBlock run reading ${JSON.stringify(text)}`);
}

/** Converts one TypeScript fence with the real highlight.js adapter. */
async function highlightedInk(
  theme: ThemeName,
  text: string,
): Promise<{ ink: string | null; bold: boolean }> {
  const bytes = await convert("```ts\nconst answer = 42; // why\n```\n", {
    theme,
    highlighter: createHighlighter(),
  });
  const documentXml = await readDocxPart(bytes, "word/document.xml");
  return {
    ink: codeInk(documentXml, text),
    bold: /<w:b\/>/.test(
      paragraphs(documentXml).find((p) => pStyle(p) === "CodeBlock" && textOf(p).includes(text)) ??
        "",
    ),
  };
}

describe("finding 5: theme and highlighter compose", () => {
  it("colours a highlighted fence from the theme, not from the adapter", async () => {
    const fromDefault = await highlightedInk("default", "const");
    const fromPrint = await highlightedInk("print", "const");

    expect(fromDefault.ink).toBe(DEFAULT_THEME.codePalette["keyword"]);
    expect(fromPrint.ink).toBe(THEMES.print.codePalette["keyword"]);
    // The bug: both used to come out as the print ink whatever the theme said.
    expect(fromDefault.ink).not.toBe(fromPrint.ink);
  });

  it("keeps the greyscale signals whichever palette is in force", async () => {
    // Weight and slant are not colour, so they are not the theme's to choose:
    // a .docx is a print artefact under either palette.
    expect((await highlightedInk("default", "const")).bold).toBe(true);
    expect((await highlightedInk("print", "const")).bold).toBe(true);
  });

  it("lets a caller take the palette back with scopeStyles", async () => {
    const bytes = await convert("```ts\nconst answer = 42;\n```\n", {
      theme: "print",
      highlighter: createHighlighter({ scopeStyles: { keyword: { color: "112233" } } }),
    });

    expect(codeInk(await readDocxPart(bytes, "word/document.xml"), "const")).toBe("112233");
  });

  it("colours model-carried scope spans from the same palette", async () => {
    // The renderer's other highlighting path must agree with the adapter's, or
    // the same document renders two different ways depending on how it was built.
    const model = doc([
      codeBlock("SELECT 1;\n", {
        lang: "sql",
        highlights: [{ start: 0, end: 6, scope: "keyword" }],
      }),
    ]);
    const asDefault = await renderParts(model, { theme: THEMES.default });
    const asPrint = await renderParts(model, { theme: THEMES.print });

    expect(codeInk(asDefault.document, "SELECT")).toBe(DEFAULT_THEME.codePalette["keyword"]);
    expect(codeInk(asPrint.document, "SELECT")).toBe(THEMES.print.codePalette["keyword"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 6 - THEMES.print merged rather than replaced                        */
/* -------------------------------------------------------------------------- */

const SHADING = DEFAULT_THEME.colors.codeBackground;

describe("finding 6: the print theme is really the print theme", () => {
  it("replaces the default palette instead of merging over it", () => {
    // These are default (screen) inks with no print counterpart. Merging let
    // them survive into `print`, and `tag`/`selector` (116329, 6.94:1) were
    // what broke the AAA floor the docblock promises.
    for (const scope of ["tag", "selector", "operator", "punctuation", "subst", "params"]) {
      expect(
        THEMES.print.codePalette[scope],
        `print theme still carries "${scope}"`,
      ).toBeUndefined();
    }
    expect(Object.keys(THEMES.print.codePalette).sort()).toEqual(
      Object.keys(PRINT_CODE_PALETTE).sort(),
    );
  });

  it("clears WCAG AAA for every ink actually reachable through the print theme", () => {
    const failures = Object.entries(THEMES.print.codePalette)
      .map(([scope, ink]) => ({
        scope,
        ink,
        onShading: contrast(ink, SHADING),
        onWhite: contrast(ink, "FFFFFF"),
      }))
      .filter((row) => row.onShading < WCAG_AAA || row.onWhite < WCAG_AAA);

    expect(failures).toEqual([]);
    // The documented floor, recomputed from the palette rather than restated.
    const floor = Math.min(
      ...Object.values(THEMES.print.codePalette).map((ink) => contrast(ink, SHADING)),
    );
    expect(floor).toBeCloseTo(7.13, 2);
  });

  it("clears WCAG AA for every ink reachable through the default theme", () => {
    // Not AAA - the default palette is GitHub's screen one - but a scope that
    // cannot be read on the shading it is painted on is a defect either way,
    // and finding 5 is what made this palette reachable at all.
    const failures = Object.entries(THEMES.default.codePalette)
      .map(([scope, ink]) => ({ scope, ink, onShading: contrast(ink, SHADING) }))
      .filter((row) => row.onShading < WCAG_AA);

    expect(failures).toEqual([]);
  });

  it("replaces rather than merges for any caller-supplied palette", () => {
    const theme = resolveTheme({ codePalette: { keyword: "010203" } });

    expect(theme.codePalette).toEqual({ keyword: "010203" });
    expect(scopeColor(theme, "string")).toBeNull();
    // Omitting it still inherits the whole default palette.
    expect(resolveTheme({}).codePalette).toEqual(DEFAULT_THEME.codePalette);
  });

  it("does not pay for the finer spans with extra runs in document.xml", async () => {
    // Reporting scopes splits spans the highlighter used to merge; the renderer
    // merges them again once it knows the theme. If it stopped, a big fence
    // would quietly double its `<w:r>` count.
    const block = codeBlock("a1 b\n");
    const parts = await renderParts(doc([block]), {
      highlights: new Map([
        [
          block,
          [
            { text: "a", scope: "attr" },
            { text: "1", scope: "number" },
            { text: " ", scope: undefined },
            { text: "b", scope: undefined },
          ],
        ],
      ]),
    });

    const code = paragraphs(parts.document).filter((p) => pStyle(p) === "CodeBlock");
    // `attr` and `number` share an ink in both built-in themes, so the four
    // spans collapse back to two runs: one coloured, one plain.
    expect([...(code[0] ?? "").matchAll(/<w:r>[\s\S]*?<\/w:r>/g)]).toHaveLength(2);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 7 - the documented purity/determinism guarantee                     */
/* -------------------------------------------------------------------------- */

/** Every zip entry of a packed document, as bytes keyed by path. */
async function zipEntries(bytes: Uint8Array): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, string>();
  for (const [path, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    out.set(path, await entry.async("string"));
  }
  return out;
}

/**
 * The two values `docx` mints from the ambient environment rather than from the
 * model, and the only reason two packs of one document are not byte-identical.
 *
 * `renderDocument`'s docblock has to name both; the test below is what keeps
 * that list honest, by *measuring* the difference rather than trusting it.
 */
const AMBIENT = [
  { part: "docProps/core.xml", pattern: /<dcterms:(created|modified)[^<]*<\/dcterms:\1>/g },
  { part: "word/_rels/document.xml.rels", pattern: /rId[A-Za-z0-9_-]{21}/g },
  { part: "word/document.xml", pattern: /rId[A-Za-z0-9_-]{21}/g },
] as const;

function scrubAmbient(part: string, xml: string): string {
  let out = xml;
  for (const ambient of AMBIENT) {
    if (ambient.part === part) out = out.replace(ambient.pattern, "AMBIENT");
  }
  return out;
}

describe("finding 7: what 'deterministic' actually means", () => {
  it("differs between two packs in exactly the entries the docblock names", async () => {
    const model = kitchenSink();
    const first = await zipEntries(await packDocument(renderDocument(model)));
    const second = await zipEntries(await packDocument(renderDocument(model)));

    expect([...second.keys()].sort()).toEqual([...first.keys()].sort());

    const differing = [...first.keys()]
      .filter((path) => first.get(path) !== second.get(path))
      .sort();

    // Nothing downword itself writes is in this list - only the parts carrying
    // a value docx generates: two wall-clock timestamps and a nanoid rId.
    expect(differing).toEqual([...new Set(AMBIENT.map((a) => a.part))].sort());
  });

  it("is byte-identical once those two values are substituted", async () => {
    const model = kitchenSink();
    const first = await zipEntries(await packDocument(renderDocument(model)));
    const second = await zipEntries(await packDocument(renderDocument(model)));

    for (const [path, xml] of first) {
      expect(scrubAmbient(path, second.get(path) ?? ""), `${path} is not reproducible`).toBe(
        scrubAmbient(path, xml),
      );
    }
  });

  it("names all of them in the renderer's own docblock", async () => {
    // A documentation regression test: the claim this module used to make
    // ("byte-identical document out") was false, and prose has no other guard.
    const source = await readFile(new URL("../src/render/index.ts", import.meta.url), "utf8");
    const docblock = source.slice(0, source.indexOf("*/"));

    for (const part of new Set(AMBIENT.map((a) => a.part))) {
      expect(docblock, `the docblock does not mention ${part}`).toContain(part);
    }
    expect(docblock).not.toContain("byte-identical document out");
    // The third source is the container, not a part - see the test below.
    expect(docblock, "the docblock does not mention the zip container").toContain(
      "the zip container",
    );
  });

  it("is not byte-reproducible even so, because the zip stamps every entry", async () => {
    // Substituting the two ambient values does *not* buy a reproducible
    // artefact, which the docblock used to imply: Packer adds each part to
    // JSZip with no `date`, so every local file header carries the clock. This
    // measures the mechanism rather than sleeping through a DOS-timestamp
    // boundary, so it is fast and cannot flake.
    const bytes = await packDocument(renderDocument(minimalDocument()));
    const zip = await JSZip.loadAsync(bytes);
    const entries = Object.values(zip.files).filter((file) => !file.dir);

    expect(entries.length).toBeGreaterThan(0);
    for (const entry of entries) {
      const age = Date.now() - entry.date.getTime();
      // Written from the wall clock at pack time, not from a fixed epoch.
      expect(Math.abs(age), `${entry.name} carries a fixed timestamp`).toBeLessThan(60_000);
    }
  });

  it("does not invent an author for a document that has none", async () => {
    // docx defaults `dc:creator` and `cp:lastModifiedBy` to "Un-named", which
    // Word then shows as the Author of every document downword produces.
    const bytes = await packDocument(renderDocument(minimalDocument()));
    const core = await readDocxPart(bytes, "docProps/core.xml");

    expect(core).not.toContain("Un-named");
    expect(core).not.toContain("<dc:creator>");
    expect(core).not.toContain("<cp:lastModifiedBy>");
  });

  it("still writes the author it was given, on both fields", async () => {
    const model = doc([paragraph([text("x")])], metadata({ author: "Ada" }));
    const core = await readDocxPart(await packDocument(renderDocument(model)), "docProps/core.xml");

    expect(core).toContain("<dc:creator>Ada</dc:creator>");
    expect(core).toContain("<cp:lastModifiedBy>Ada</cp:lastModifiedBy>");
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 8 - `html: "escape"` raises no warnings                             */
/* -------------------------------------------------------------------------- */

const HTML_SAMPLE = "a <br> b\n\n<div>x</div>\n";

async function convertWithWarnings(
  markdown: string,
  options: Parameters<typeof convert>[1] = {},
): Promise<{ xml: string; warnings: ConvertWarning[] }> {
  const warnings: ConvertWarning[] = [];
  const bytes = await convert(markdown, { ...options, onWarning: (w) => warnings.push(w) });
  return { xml: await readDocxPart(bytes, "word/document.xml"), warnings };
}

describe("finding 8: the html policies say what they do", () => {
  it("escape parses no HTML, so it has nothing to warn about", async () => {
    const { xml, warnings } = await convertWithWarnings(HTML_SAMPLE, { html: "escape" });

    expect(warnings).toEqual([]);
    // Nothing is lost, which is *why* there is nothing to report: the markup
    // survives verbatim as text.
    expect(textOf(xml)).toContain("<br>");
    expect(textOf(xml)).toContain("<div>x</div>");
  });

  it("keep and drop warn once per construct", async () => {
    const kept = await convertWithWarnings(HTML_SAMPLE, { html: "keep" });
    const dropped = await convertWithWarnings(HTML_SAMPLE, { html: "drop" });

    expect(kept.warnings.map((w) => w.code).sort()).toEqual([
      "html-block",
      "html-inline",
      "raw-html",
      "raw-html",
    ]);
    expect(dropped.warnings.map((w) => w.code).sort()).toEqual(
      kept.warnings.map((w) => w.code).sort(),
    );
    expect(textOf(kept.xml)).toContain("<div>x</div>");
    expect(textOf(dropped.xml)).not.toContain("<div>x</div>");
  });

  it("says so in the TSDoc for HtmlHandling", async () => {
    const source = await readFile(new URL("../src/options.ts", import.meta.url), "utf8");
    expect(source).not.toContain("All three raise a warning per construct");
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 9 - superscript + subscript emit two w:vertAlign                    */
/* -------------------------------------------------------------------------- */

describe("finding 9: conflicting vertical-alignment marks", () => {
  it("emits one w:vertAlign, not two, and says which mark lost", async () => {
    const parts = await render([paragraph([text("x", [superscript(), subscript()])])]);
    const run = paragraphWithText(parts.document, "x");

    expect([...run.matchAll(/<w:vertAlign /g)]).toHaveLength(1);
    expect(run).toContain('<w:vertAlign w:val="superscript"/>');
    expect(parts.warnings).toEqual([
      {
        code: "conflicting-marks",
        severity: "notice",
        message: expect.stringContaining("subscript"),
      },
    ]);
  });

  it("leaves either mark alone on its own", async () => {
    const parts = await render([
      paragraph([text("up", [superscript()])]),
      paragraph([text("down", [subscript()])]),
    ]);

    expect(paragraphWithText(parts.document, "up")).toContain('<w:vertAlign w:val="superscript"/>');
    expect(paragraphWithText(parts.document, "down")).toContain('<w:vertAlign w:val="subscript"/>');
    expect(parts.warnings).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 10 - two model fields dropped with no warning                       */
/* -------------------------------------------------------------------------- */

describe("finding 10: fields OOXML has nowhere to put", () => {
  it("warns that a link title cannot become a tooltip", async () => {
    const parts = await render([paragraph([text("t", [link("https://x.test", "hover tooltip")])])]);

    expect(parts.document).not.toContain("w:tooltip");
    expect(parts.warnings).toEqual([
      {
        code: "link-title-dropped",
        severity: "notice",
        message: expect.stringContaining("hover tooltip"),
      },
    ]);
  });

  it("says nothing for a link with no title", async () => {
    const parts = await render([paragraph([text("t", [link("https://x.test")])])]);
    expect(parts.warnings).toEqual([]);
  });

  it("warns that fence metadata has nowhere to go", async () => {
    const parts = await render([codeBlock("x\n", { lang: "ts", meta: 'title="app.ts"' })]);

    expect(parts.warnings).toEqual([
      {
        code: "code-meta-dropped",
        severity: "notice",
        message: expect.stringContaining('title=\\"app.ts\\"'),
      },
    ]);
  });

  it("reaches a caller of convert(), with its stage", async () => {
    const { warnings } = await convertWithWarnings(
      '[t](https://x.test "hover")\n\n```ts title="app.ts"\nx\n```\n',
    );

    expect(warnings.map((w) => `${w.stage}:${w.code}`)).toEqual([
      "render:link-title-dropped",
      "render:code-meta-dropped",
    ]);
  });
});

/* -------------------------------------------------------------------------- */
/* Finding 11 - the async prepare passes had no diagnostic channel             */
/* -------------------------------------------------------------------------- */

describe("finding 11: the prepare passes report what they lost", () => {
  it("prepareHighlights reports every block a highlighter threw on", async () => {
    const warnings: RenderWarning[] = [];
    const model = doc([codeBlock("a\n", { lang: "ts" }), codeBlock("b\n", { lang: "py" })]);

    const map = await prepareHighlights(
      model,
      {
        highlight: () => {
          throw new Error("boom");
        },
      },
      { onWarning: (w) => warnings.push(w) },
    );

    expect(map.size).toBe(0);
    expect(warnings.map((w) => w.code)).toEqual(["highlighter-failed", "highlighter-failed"]);
    expect(warnings[0]?.severity).toBe("notice");
    expect(warnings[0]?.message).toContain("boom");
  });

  it("prepareImages reports a resolver that throws", async () => {
    const warnings: RenderWarning[] = [];
    const model = doc([paragraph([image("a.png")])]);

    const map = await prepareImages(
      model,
      {
        resolve: () => Promise.reject(new Error("nope")),
      },
      { onWarning: (w) => warnings.push(w) },
    );

    expect(map.size).toBe(0);
    expect(warnings).toEqual([
      {
        code: "image-resolver-failed",
        severity: "error",
        message: expect.stringContaining("nope"),
      },
    ]);
  });

  it("still degrades quietly when no handler is supplied", async () => {
    const model = doc([codeBlock("a\n")]);
    await expect(
      prepareHighlights(model, {
        highlight: () => {
          throw new Error("boom");
        },
      }),
    ).resolves.toEqual(new Map());
  });

  it("folds a throwing highlighter into convert()'s warning stream", async () => {
    const { warnings } = await convertWithWarnings("```ts\nx\n```\n", {
      highlighter: {
        highlight: () => Promise.reject(new Error("boom")),
      },
    });

    expect(warnings).toEqual([
      {
        stage: "render",
        code: "highlighter-failed",
        severity: "notice",
        message: expect.stringContaining("boom"),
      },
    ]);
  });
});

/* -------------------------------------------------------------------------- */
/* Findings 4, 8, 10, 11 - one channel, one shape                              */
/* -------------------------------------------------------------------------- */

describe("the diagnostics channel", () => {
  it("gives every stage's warning the same three fields", async () => {
    const { warnings } = await convertWithWarnings(
      // One warning from each of the three stages, in one document.
      `[t](https://x.test "hover")\n\n![m](./nope.png)\n\n<div>x</div>\n`,
      { html: "keep" },
    );

    expect(warnings.length).toBeGreaterThanOrEqual(3);
    expect(new Set(warnings.map((w) => w.stage))).toEqual(new Set(["parse", "image", "render"]));
    for (const warning of warnings) {
      expect(typeof warning.code).toBe("string");
      expect(["error", "notice"]).toContain(warning.severity);
      expect(warning.message.length).toBeGreaterThan(0);
    }
  });

  it("answers the severity question from a total table, per stage", () => {
    // Every code a stage can raise must have an answer, and it must be the same
    // answer everywhere - which is only true while these are lookups, not
    // decisions made at the call site.
    for (const code of RENDER_WARNING_CODES) {
      expect(["error", "notice"]).toContain(renderWarningSeverity(code));
    }
    for (const code of PARSE_WARNING_CODES) {
      expect(["error", "notice"]).toContain(parseWarningSeverity(code));
    }
  });

  it("cannot be broken by a handler that throws", async () => {
    const model = doc([codeBlock("a\n", { lang: "ts", meta: "x" })]);
    expect(() =>
      renderDocument(model, {
        onWarning: () => {
          throw new Error("logger exploded");
        },
      }),
    ).not.toThrow();
  });
});

/* -------------------------------------------------------------------------- */
/* Integration verification: the block nesting ceiling                         */
/* -------------------------------------------------------------------------- */

describe("nesting limit: markdown-it discards blocks past MAX_BLOCK_NESTING", () => {
  /** A blockquote nested `depth` deep with `payload` as its only content. */
  const quote = (depth: number): string => `${">".repeat(depth)} payload\n`;

  /** A bullet list nested `depth` deep, each item labelled `level N`. */
  const list = (depth: number): string =>
    `${Array.from({ length: depth }, (_, i) => `${"  ".repeat(i)}- level ${i + 1}`).join("\n")}\n`;

  function parseWith(source: string): { text: string; warnings: ParseWarning[] } {
    const warnings: ParseWarning[] = [];
    const model = parseMarkdown(source, { onWarning: (warning) => warnings.push(warning) });
    return { text: nodeText(model), warnings };
  }

  const nesting = (warnings: readonly ParseWarning[]): ParseWarning[] =>
    warnings.filter((warning) => warning.code === "nesting-limit");

  it("reports the lines it loses instead of dropping them in silence", () => {
    // This is the one place content can disappear before a node ever exists:
    // ParserBlock.tokenize bails with `state.line = endLine`, pushing nothing.
    // Before the guard this produced a completely empty document, no warning.
    const { text: lost, warnings } = parseWith(quote(MAX_BLOCK_NESTING + 20));

    expect(lost).toBe("");
    expect(nesting(warnings)).toHaveLength(1);
    const [warning] = nesting(warnings);
    expect(warning?.severity).toBe("error");
    expect(warning?.line).toBe(1);
    expect(warning?.message).toContain(String(MAX_BLOCK_NESTING));
    expect(warning?.message).toContain("1 source line");
  });

  it("fires at the depth where content is actually lost, and not one level before", () => {
    // The guard mirrors tokenize's own first-iteration test, so the boundary it
    // reports has to be exactly the boundary markdown-it enforces.
    const under = parseWith(quote(MAX_BLOCK_NESTING - 1));
    expect(under.text).toBe("payload");
    expect(nesting(under.warnings)).toHaveLength(0);

    const at = parseWith(quote(MAX_BLOCK_NESTING));
    expect(at.text).toBe("");
    expect(nesting(at.warnings)).toHaveLength(1);
  });

  it("counts a list level as two containers, the way markdown-it does", () => {
    // Each level is a list *and* an item, so the ceiling arrives at half the
    // depth a blockquote reaches.
    const deep = MAX_BLOCK_NESTING / 2;

    const under = parseWith(list(deep - 1));
    expect(under.text).toContain(`level ${deep - 1}`);
    expect(nesting(under.warnings)).toHaveLength(0);

    const at = parseWith(list(deep));
    expect(at.text).toContain(`level ${deep - 1}`);
    expect(at.text).not.toContain(`level ${deep}`);
    expect(nesting(at.warnings)).toHaveLength(1);
    expect(nesting(at.warnings)[0]?.line).toBe(deep);
  });

  it("stays quiet on markdown anyone would actually write", () => {
    // A guard that cried wolf on ordinary nesting would be worse than none.
    const ordinary = parseWith(
      "# Title\n\n> quoted\n>\n> > deeper\n\n- a\n  - b\n    - c\n      - d\n\n1. one\n\n   ```ts\n   const x = 1;\n   ```\n\n| a | b |\n| - | - |\n| 1 | 2 |\n",
    );
    expect(nesting(ordinary.warnings)).toHaveLength(0);
  });

  it("reaches convert()'s shared warning stream, tagged parse", async () => {
    const warnings: ConvertWarning[] = [];
    await convert(quote(MAX_BLOCK_NESTING + 5), {
      onWarning: (warning) => warnings.push(warning),
    });

    const found = warnings.filter((warning) => warning.code === "nesting-limit");
    expect(found).toHaveLength(1);
    expect(found[0]?.stage).toBe("parse");
    expect(found[0]?.severity).toBe("error");
  });

  it("still degrades quietly, and still produces a document, with no handler", async () => {
    expect(() => parseMarkdown(quote(MAX_BLOCK_NESTING + 5))).not.toThrow();
    const bytes = await convert(quote(MAX_BLOCK_NESTING + 5));
    expect(bytes.byteLength).toBeGreaterThan(0);
  });
});

/* ========================================================================== */
/* SECOND REVIEW                                                              */
/*                                                                            */
/* The findings below come from the independent second review of Phase 1,     */
/* which renumbered from 1. Its numbers collide with the first review's, so   */
/* every describe here is prefixed `review 2` - a failure has to point at one */
/* report, not at two.                                                        */
/* ========================================================================== */

/* -------------------------------------------------------------------------- */
/* Review 2, finding 1 - a Highlighter could silently replace a block's text   */
/* -------------------------------------------------------------------------- */

/** A `Highlighter` that returns exactly `spans`, whatever it is given. */
function fixedHighlighter(spans: readonly HighlightSpan[]): Highlighter {
  return { highlight: () => spans };
}

/** Every `<w:t>` of every `CodeBlock` paragraph, joined with newlines. */
function codeText(documentXml: string): string {
  return paragraphs(documentXml)
    .filter((p) => pStyle(p) === "CodeBlock")
    .map(textOf)
    .join("\n");
}

describe("review 2, finding 1: a highlighter cannot change what the document says", () => {
  const SOURCE = "const secret = 1;\nconsole.log(secret);\n";
  const FENCE = "```js\nconst secret = 1;\nconsole.log(secret);\n```\n";

  it("restores the source when an adapter substitutes different text", async () => {
    // The report's exact reproduction: two lines of code became the word OOPS,
    // with no warning at all.
    const warnings: ConvertWarning[] = [];
    const bytes = await convert(FENCE, {
      highlighter: fixedHighlighter([{ text: "OOPS" }]),
      onWarning: (warning) => warnings.push(warning),
    });

    const documentXml = await readDocxPart(bytes, "word/document.xml");
    expect(codeText(documentXml)).toBe("const secret = 1;\nconsole.log(secret);");
    expect(documentXml).not.toContain("OOPS");
    expect(warnings.map((warning) => warning.code)).toContain("highlighter-mismatch");
  });

  it("does not emit the block twice when an adapter duplicates it", async () => {
    const warnings: ConvertWarning[] = [];
    const bytes = await convert(FENCE, {
      highlighter: fixedHighlighter([{ text: SOURCE }, { text: SOURCE }]),
      onWarning: (warning) => warnings.push(warning),
    });

    expect(codeText(await readDocxPart(bytes, "word/document.xml"))).toBe(
      "const secret = 1;\nconsole.log(secret);",
    );
    expect(warnings.map((warning) => warning.code)).toContain("highlighter-mismatch");
  });

  it("keeps the colours and restores the tail when the spans are a proper prefix", async () => {
    // What a grammar that hit an `illegal` match leaves behind: losing the
    // colour on the tail is much better than losing the tail.
    const block = codeBlock(SOURCE, { lang: "js" });
    const parts = await renderParts(doc([block]), {
      highlights: new Map([[block, [{ text: "const", scope: "keyword" }]]]),
    });

    expect(codeText(parts.document)).toBe("const secret = 1;\nconsole.log(secret);");
    expect(codeInk(parts.document, "const")).toBe(DEFAULT_THEME.codePalette["keyword"]);
    const mismatch = parts.warnings.filter((w) => w.code === "highlighter-mismatch");
    expect(mismatch).toHaveLength(1);
    expect(mismatch[0]?.message).toContain("short of the end");
  });

  it("catches a same-length substitution too, not just a length change", async () => {
    const block = codeBlock("abcd\n", { lang: "text" });
    const parts = await renderParts(doc([block]), {
      highlights: new Map([[block, [{ text: "xy" }, { text: "zw\n" }]]]),
    });

    expect(codeText(parts.document)).toBe("abcd");
    expect(parts.warnings.map((w) => w.code)).toEqual(["highlighter-mismatch"]);
  });

  it("says nothing, and keeps every colour, for an adapter that honours the contract", async () => {
    const parts = await renderParts(doc([codeBlock("const x = 1;\n", { lang: "js" })]), {
      highlighter: fixedHighlighter([
        { text: "const", scope: "keyword" },
        { text: " x = " },
        { text: "1", scope: "number" },
        { text: ";\n" },
      ]),
    });

    expect(codeText(parts.document)).toBe("const x = 1;");
    expect(codeInk(parts.document, "const")).toBe(DEFAULT_THEME.codePalette["keyword"]);
    expect(parts.warnings).toEqual([]);
  });

  it("holds on both paths: the prepared map and the synchronous highlighter", async () => {
    const block = codeBlock(SOURCE, { lang: "js" });
    const prepared = await renderParts(doc([block]), {
      highlights: new Map([[block, [{ text: "OOPS" }]]]),
    });
    const inline = await renderParts(doc([block]), {
      highlighter: fixedHighlighter([{ text: "OOPS" }]),
    });

    for (const parts of [prepared, inline]) {
      expect(codeText(parts.document)).toBe("const secret = 1;\nconsole.log(secret);");
      expect(parts.warnings.map((w) => w.code)).toEqual(["highlighter-mismatch"]);
    }
  });

  it("keeps downword's own highlighter free of the warning", async () => {
    const warnings: ConvertWarning[] = [];
    await convert(FENCE, {
      highlighter: createHighlighter(),
      onWarning: (warning) => warnings.push(warning),
    });

    expect(warnings.map((warning) => warning.code)).not.toContain("highlighter-mismatch");
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 2 - two adjacent tables are one table to Word             */
/* -------------------------------------------------------------------------- */

/**
 * The `<w:tbl>` / `<w:p>` skeleton of a body, in document order.
 *
 * **Top-level children only.** Every cell of every table holds a `<w:p>` of its
 * own (`renderTableCell` — a `w:tc` with no paragraph is invalid), so a flat
 * scan for `<w:p>` counts the table's contents as siblings of the table and
 * reports one extra `P` per cell. What this assertion is about is what sits
 * *between* two `w:tbl` elements in the body, so the scan tracks table depth and
 * records nothing inside one.
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

/** A one-row table with `count` columns, its first cell reading `label`. */
function tinyTable(label: string, count: number): TableNode {
  return table([
    headerRow(
      Array.from({ length: count }, (_, index) =>
        tableCell([text(index === 0 ? label : `c${index}`)]),
      ),
    ),
  ]);
}

describe("review 2, finding 2: adjacent tables are kept apart", () => {
  it("never emits </w:tbl><w:tbl>", async () => {
    // Two markdown tables split by a blank line - ordinary LLM output. Word
    // merges two adjacent `w:tbl` into one, re-gridding the second table's
    // rows into the first's columns.
    const bytes = await convert("| a | b |\n| - | - |\n| 1 | 2 |\n\n| c |\n| - |\n| 3 |\n");
    const documentXml = await readDocxPart(bytes, "word/document.xml");

    expect(documentXml).not.toContain("</w:tbl><w:tbl>");
    expect(bodySkeleton(documentXml)).toEqual(["TBL", "P", "TBL", "P", "SECT"]);
  });

  it("separates two tables the model puts side by side, with no blank line to help", async () => {
    const parts = await render([tinyTable("first", 2), tinyTable("second", 1)]);

    expect(parts.document).not.toContain("</w:tbl><w:tbl>");
    expect(bodySkeleton(parts.document)).toEqual(["TBL", "P", "TBL", "P", "SECT"]);
    // The two grids stay different, which is the corruption the report saw:
    // the second table's single column was absorbed into the first's two.
    const grids = [...parts.document.matchAll(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/g)].map(
      (m) => [...(m[1] ?? "").matchAll(/<w:gridCol/g)].length,
    );
    expect(grids).toEqual([2, 1]);
  });

  it("closes a document whose last block is a table", async () => {
    // `[TBL][SECT]` is tolerated but repaired: Word synthesises a paragraph of
    // its own, so the file it saves is not the file we wrote.
    const parts = await render([paragraph([text("before")]), tinyTable("last", 2)]);

    expect(bodySkeleton(parts.document)).toEqual(["P", "TBL", "P", "SECT"]);
  });

  it("uses a real style for the separator rather than an unstyled paragraph", async () => {
    const parts = await render([tinyTable("only", 1)]);
    const after = paragraphs(parts.document).at(-1) ?? "";

    expect(pStyle(after)).toBe("TableSpacing");
    expect(textOf(after)).toBe("");
  });

  it("keeps three tables in a row three tables", async () => {
    const parts = await render([tinyTable("a", 1), tinyTable("b", 2), tinyTable("c", 3)]);

    expect(bodySkeleton(parts.document)).toEqual(["TBL", "P", "TBL", "P", "TBL", "P", "SECT"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 3 - renderDocument never sanitized                        */
/* -------------------------------------------------------------------------- */

/**
 * Every character XML 1.0 cannot carry, as a regular expression over UTF-16.
 *
 * `no-control-regex` exists to catch control characters written into a pattern
 * by accident. Matching them is the entire point here, so the rule is off for
 * this one declaration - as a disable/enable pair rather than
 * `disable-next-line`, which prettier detaches from the pattern as soon as the
 * declaration is long enough to wrap after the `=`.
 */
/* eslint-disable no-control-regex */
const ILLEGAL_XML =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDFFF](?![\uDC00-\uDFFF])/;
/* eslint-enable no-control-regex */

/** Reads every text part of a `.docx` and returns them keyed by path. */
async function textParts(bytes: Uint8Array): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, string>();
  for (const [path, entry] of Object.entries(zip.files)) {
    if (entry.dir || (!path.endsWith(".xml") && !path.endsWith(".rels"))) continue;
    out.set(path, await entry.async("string"));
  }
  return out;
}

/** Asserts that no part of the package carries a character XML 1.0 forbids. */
async function expectWellFormed(bytes: Uint8Array): Promise<void> {
  for (const [path, xml] of await textParts(bytes)) {
    expect(ILLEGAL_XML.test(xml), `${path} carries a character XML 1.0 cannot represent`).toBe(
      false,
    );
  }
}

describe("review 2, finding 3: renderDocument sanitizes a consumer-built model", () => {
  it("replaces a control character in a text node instead of writing it", async () => {
    // A model built from a database row, another AST or a hand-written literal
    // has been through neither parseMarkdown nor resolveConvertOptions.
    const parts = await renderParts(doc([paragraph([text("before\u0001after")])]));

    expect(textOf(parts.document)).toBe("before\uFFFDafter");
    await expectWellFormed(parts.bytes);
    expect(parts.warnings.map((w) => w.code)).toEqual(["unrepresentable-character"]);
  });

  it("sanitizes docProps/core.xml, which is a separate part and just as fatal", async () => {
    const model = doc(
      [paragraph([text("body")])],
      metadata({
        title: "Ti\u000Btle",
        author: "A\u0007uthor",
        description: "d\u0000esc",
        keywords: ["k\u001Fey"],
        custom: { "na\u000Cme": "va\u0001lue" },
      }),
    );
    const file = renderDocument(model, { titleBlock: false });
    const bytes = await packDocument(file);

    const core = await readDocxPart(bytes, "docProps/core.xml");
    expect(core).toContain("<dc:title>Ti\uFFFDtle</dc:title>");
    await expectWellFormed(bytes);
  });

  it("sanitizes every other render sink: image alt, raw HTML, math, the title block", async () => {
    const model = doc(
      [
        paragraph([image("s\u0001rc.png", { alt: "a\u0001lt", title: "t\u0001itle" })]),
        htmlBlock("<div>\u0001</div>\n"),
        mathBlock("x\u0001 = 1"),
        paragraph([htmlInline("<b\u0001>")]),
        paragraph([mathInline("y\u0001")]),
      ],
      metadata({ title: "T\u0001" }),
    );
    const bytes = await packDocument(renderDocument(model, { titleBlock: true }));

    await expectWellFormed(bytes);
    const documentXml = await readDocxPart(bytes, "word/document.xml");
    expect(documentXml).toContain("\uFFFD");
  });

  it("reports it once for a document full of them, not once per run", async () => {
    const parts = await renderParts(
      doc(Array.from({ length: 20 }, () => paragraph([text("a\u0001b")]))),
    );

    expect(parts.warnings.map((w) => w.code)).toEqual(["unrepresentable-character"]);
    expect(parts.warnings[0]?.severity).toBe("notice");
  });

  it("costs nothing and says nothing for a clean model", async () => {
    const parts = await renderParts(minimalDocument());

    expect(parts.warnings.filter((w) => w.code === "unrepresentable-character")).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 4 - print inks overrode the theme for unresolvable scopes */
/* -------------------------------------------------------------------------- */

/** One fence per language, chosen to exercise as many scopes as each grammar has. */
const LANGUAGE_SAMPLES: readonly (readonly [string, string])[] = [
  ["ts", "// note\nexport class A extends B { run(x: number): string { return `v${x}`; } }\n"],
  ["js", "const re = /a+/g; // c\nfunction f(a) { return a?.b ?? 1_000; }\n"],
  ["python", "@dec\ndef f(a: int = 3) -> str:\n    '''doc'''\n    return f\"{a!r}\"\n"],
  ["css", "a.b:hover, #id[data-x] { color: red; --v: 1px; }\n@media print { .c { top: 0 } }\n"],
  ["scss", "$v: 1px;\n.a { &:hover { color: red } }\n"],
  ["html", '<!doctype html>\n<div class="a" id="b">x<script>var y=1;</script></div>\n'],
  ["xml", '<?xml version="1.0"?>\n<a b="c"><!-- d --></a>\n'],
  ["json", '{"a": 1, "b": [true, null, "s"]}\n'],
  ["yaml", "a: 1\nb:\n  - c\n  - &anchor d\n# comment\n"],
  ["bash", '#!/bin/sh\nset -e\nfor f in *.txt; do echo "$f" | wc -l; done\n'],
  ["sql", "SELECT a, COUNT(*) FROM t WHERE b = 'x' -- note\nGROUP BY a;\n"],
  ["diff", "--- a\n+++ b\n@@ -1 +1 @@\n-old\n+new\n"],
  ["markdown", "# h\n\n**b** `c` [l](u)\n\n> q\n\n- i\n"],
  ["go", 'package main\nimport "fmt"\nfunc main() { fmt.Println("x") }\n'],
  ["rust", '#[derive(Debug)]\npub struct S { a: u32 }\nfn main() { println!("{}", 1); }\n'],
  ["java", "public class A { private static final int X = 1; /* c */ }\n"],
  ["ruby", "class A\n  def b(c) = :sym\n  # comment\nend\n"],
  ["php", '<?php\nnamespace A;\nfunction b(int $c): string { return "x$c"; }\n'],
  ["c", "#include <stdio.h>\nint main(void){ return 0; }\n"],
  ["ini", "[a]\nb = 1 ; c\n"],
  ["dockerfile", "FROM node:20\nRUN echo hi\n"],
  ["latex", "\\documentclass{article}\n\\begin{document}$x^2$\\end{document}\n"],
  ["http", "GET /a HTTP/1.1\nHost: x\n"],
  ["graphql", "query Q($a: Int!) { b(c: $a) { d } }\n"],
];

/**
 * Scopes `PRINT_SCOPE_STYLES` documents as deliberately uncoloured.
 *
 * They fall through to the `CodeBlock` style's own near-black, which is the
 * point: a printed page with more than about eight text colours reads as noise.
 * Listed here so the totality check below can tell "deliberately absent" from
 * "forgotten", which is the distinction the whole finding turns on.
 */
const PRINT_UNCOLOURED = [
  "punctuation",
  "operator",
  "subst",
  "params",
  "tag",
  "strong",
  "function",
];

describe("review 2, finding 4: the theme owns code colour, with no adapter fallback", () => {
  it("themes a stylesheet, whose scopes the prefix walk cannot reach", async () => {
    // The report's exact case. `selector-tag` and friends are dash-separated,
    // so `scopeColor`'s dot walk cannot reach `selector` from them - and while
    // a span's own ink was still a fallback, every one of them stayed printed
    // in `PRINT_INK.markup` whatever the theme said.
    const fence = "```css\na.b:hover { color: red; }\n```\n";
    const options = { highlighter: createHighlighter() } as const;
    const asDefault = await readDocxPart(
      await convert(fence, { ...options, theme: "default" }),
      "word/document.xml",
    );
    const asPrint = await readDocxPart(
      await convert(fence, { ...options, theme: "print" }),
      "word/document.xml",
    );

    expect(codeInk(asDefault, "a")).toBe(DEFAULT_THEME.codePalette["selector-tag"]);
    expect(codeInk(asPrint, "a")).toBe(THEMES.print.codePalette["selector-tag"]);
    expect(codeInk(asDefault, "a")).not.toBe(codeInk(asPrint, "a"));
    // Byte-identical output for the two themes was the measurement that failed.
    expect(asDefault).not.toBe(asPrint);
  });

  it("makes codePalette's replace-not-merge semantics observable at render time", async () => {
    // Measured before the fix: identical to `theme: "default"`, because every
    // scope the one-entry palette did not name kept the adapter's print ink.
    const fence = "```css\na.b:hover { color: red; }\n```\n";
    const documentXml = await readDocxPart(
      await convert(fence, {
        theme: { codePalette: { "selector-tag": "B00020" } },
        highlighter: createHighlighter(),
      }),
      "word/document.xml",
    );

    expect(codeInk(documentXml, "a")).toBe("B00020");
    // Nothing else in the line is in the replacement palette, so nothing else
    // may be coloured - and because uncoloured neighbours coalesce into one run
    // (`coalescePieces`), the whole remainder arriving as a single inkless run
    // *is* the assertion: `selector-class`, `selector-pseudo` and `attribute`
    // all lost the print ink the adapter attached to them.
    expect(codeInk(documentXml, ".b:hover { color: red; }")).toBeNull();
  });

  it("ignores a span's colour whenever it reports a scope", async () => {
    const block = codeBlock("ab\n");
    const parts = await renderParts(doc([block]), {
      highlights: new Map([
        [
          block,
          [
            { text: "a", color: "FF0000", scope: "keyword" },
            { text: "b", color: "00FF00" },
          ],
        ],
      ]),
    });

    // Scope wins outright...
    expect(codeInk(parts.document, "a")).toBe(DEFAULT_THEME.codePalette["keyword"]);
    // ...and an adapter that reports only colours still keeps them.
    expect(codeInk(parts.document, "b")).toBe("00FF00");
  });

  it("resolves every scope the real highlighter emits, across two dozen grammars", async () => {
    const highlighter = createHighlighter();
    const emitted = new Set<string>();
    for (const [lang, code] of LANGUAGE_SAMPLES) {
      for (const span of await highlighter.highlight(code, lang)) {
        if (span.scope !== undefined) emitted.add(span.scope);
      }
    }

    // A sanity floor: a corpus that stopped exercising the grammars would make
    // the assertion below vacuous.
    expect(emitted.size).toBeGreaterThan(30);
    expect([...emitted].filter((scope) => scopeColor(DEFAULT_THEME, scope) === null)).toEqual([]);
    expect([...emitted].filter((scope) => scopeColor(THEMES.print, scope) === null).sort()).toEqual(
      [...PRINT_UNCOLOURED].filter((scope) => emitted.has(scope)).sort(),
    );
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 5 - a tall image ran off the bottom of the page           */
/* -------------------------------------------------------------------------- */

/** `<wp:extent cx cy>` of every picture in a document, in EMU. */
function extents(documentXml: string): { cx: number; cy: number }[] {
  return [...documentXml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => ({
    cx: Number(m[1]),
    cy: Number(m[2]),
  }));
}

/** A4 with one-inch margins, in EMU: 635 EMU to the twip. */
const A4_CONTENT_WIDTH_EMU = 9026 * 635;
const A4_CONTENT_HEIGHT_EMU = (16838 - 2880) * 635;

describe("review 2, finding 5: an image is capped in both directions", () => {
  it("scales a tall screenshot to fit the text column's height", async () => {
    // The report's measurement: a 400x8000 PNG came out 4.17in x 83.33in in a
    // column 9.19in tall, so eleven twelfths of it was simply not drawn.
    const parts = await renderParts(
      doc([paragraph([image("tall.png", { resolved: resolvedPng(400, 8000) })])]),
    );

    const [extent] = extents(parts.document);
    expect(extent).toBeDefined();
    expect(extent?.cy).toBeLessThanOrEqual(A4_CONTENT_HEIGHT_EMU);
    expect(extent?.cx).toBeLessThanOrEqual(A4_CONTENT_WIDTH_EMU);
    // Aspect ratio preserved: 400/8000 = 0.05.
    expect((extent?.cx ?? 0) / (extent?.cy ?? 1)).toBeCloseTo(0.05, 3);
  });

  it("says so, naming the picture, rather than degrading in silence", async () => {
    const parts = await renderParts(
      doc([paragraph([image("tall.png", { resolved: resolvedPng(400, 8000) })])]),
    );

    const oversized = parts.warnings.filter((w) => w.code === "image-oversized");
    expect(oversized).toHaveLength(1);
    expect(oversized[0]?.severity).toBe("notice");
    expect(oversized[0]?.message).toContain("tall.png");
  });

  it("leaves the width-limited case exactly as it was, and quiet", async () => {
    const parts = await renderParts(
      doc([paragraph([image("wide.png", { resolved: resolvedPng(1200, 600) })])]),
    );

    const [extent] = extents(parts.document);
    // 9026 twips = 601 px; 601 px = 5724525 EMU.
    expect(extent?.cx).toBe(601 * 9525);
    expect(parts.warnings.filter((w) => w.code === "image-oversized")).toEqual([]);
  });

  it("never enlarges a small picture to fill the column", async () => {
    const parts = await renderParts(
      doc([paragraph([image("icon.png", { resolved: resolvedPng(40, 40) })])]),
    );

    expect(extents(parts.document)[0]).toEqual({ cx: 40 * 9525, cy: 40 * 9525 });
  });

  it("caps a picture in a table cell by the cell's width and the page's height", async () => {
    // A row grows to fit its content, but a page does not - so the height limit
    // inside a cell is still the text column's.
    const parts = await render([
      table([
        headerRow([
          tableCell([image("tall.png", { resolved: resolvedPng(400, 8000) })]),
          tableCell([text("b")]),
        ]),
      ]),
    ]);

    const [extent] = extents(parts.document);
    expect(extent?.cy).toBeLessThanOrEqual(A4_CONTENT_HEIGHT_EMU);
  });

  it("scales against the page it was actually given", async () => {
    const short = await renderParts(
      doc([paragraph([image("tall.png", { resolved: resolvedPng(400, 8000) })])]),
      { page: { size: { width: 11906, height: 5000 } } },
    );

    const [extent] = extents(short.document);
    expect(extent?.cy).toBeLessThanOrEqual((5000 - 2880) * 635);
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 6 - no vertical space after a tight list or a table       */
/* -------------------------------------------------------------------------- */

/** The `<w:spacing …/>` attributes of a style or paragraph, as raw strings. */
function spacing(xml: string): Record<string, string> | null {
  const match = /<w:spacing ([^/>]*)\/>/.exec(xml);
  if (match === null) return null;
  const out: Record<string, string> = {};
  for (const attr of (match[1] ?? "").matchAll(/([\w:]+)="([^"]*)"/g)) {
    out[attr[1] ?? ""] = attr[2] ?? "";
  }
  return out;
}

describe("review 2, finding 6: a tight list and a table both breathe", () => {
  it("does not pin ListParagraph's `after` to zero", async () => {
    // `contextualSpacing` already suppresses the gap *between* items, which is
    // the tight look. Pinning `after: 0` as well made it redundant and welded
    // the list to whatever followed. Word's own ListParagraph does not.
    const parts = await render([bulletList([listItem([paragraph([text("a")])])])]);
    const style = styleById(parts.styles, "ListParagraph") ?? "";

    expect(style).toContain("<w:contextualSpacing/>");
    expect(spacing(style)?.["w:after"]).toBeUndefined();
  });

  it("leaves a real gap between a tight list and the paragraph after it", async () => {
    const parts = await render([
      bulletList([listItem([paragraph([text("a")])]), listItem([paragraph([text("b")])])]),
      paragraph([text("Next paragraph.")]),
    ]);

    // Neither paragraph carries direct spacing, so the gap is the inherited
    // one: `after` from docDefaults, which ListParagraph no longer zeroes.
    expect(spacing(paragraphWithText(parts.document, "b"))).toBeNull();
    const defaults = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(parts.styles)?.[0] ?? "";
    expect(spacing(defaults)?.["w:after"]).toBe(String(DEFAULT_THEME.spacing.paragraphAfter));
  });

  it("reads theme.spacing.tableSpacing, which nothing used to", async () => {
    const parts = await renderParts(doc([tinyTable("a", 2)]), {
      theme: { spacing: { tableSpacing: 333 } },
    });

    expect(spacing(styleById(parts.styles, "TableSpacing") ?? "")?.["w:after"]).toBe("333");
  });

  it("keeps the separator paragraph out of the style gallery", async () => {
    // Nobody should apply it by hand: it exists for the layout engine.
    const parts = await render([tinyTable("a", 2)]);
    const style = styleById(parts.styles, "TableSpacing") ?? "";

    expect(style).toContain("<w:semiHidden/>");
    expect(style).toContain("<w:unhideWhenUsed/>");
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 7 - every bookmark was written with w:id="1"              */
/* -------------------------------------------------------------------------- */

/** Every `<w:bookmarkStart>` in a document, as `{ name, id }`, in order. */
function bookmarkStarts(documentXml: string): { name: string; id: number }[] {
  return [...documentXml.matchAll(/<w:bookmarkStart w:name="([^"]*)" w:id="(\d+)"\/>/g)].map(
    (m) => ({ name: m[1] ?? "", id: Number(m[2]) }),
  );
}

/** Every `<w:bookmarkEnd>` id, in order. */
function bookmarkEnds(documentXml: string): number[] {
  return [...documentXml.matchAll(/<w:bookmarkEnd w:id="(\d+)"\/>/g)].map((m) => Number(m[1]));
}

describe("review 2, finding 7: bookmark ids are unique within the part", () => {
  it("gives four headings four different ids", async () => {
    // The report's measurement: all four were `w:id="1"`, because docx's own
    // `Bookmark` mints its counter per instance.
    const bytes = await convert("# one\n\n# two\n\n# three\n\n# four\n");
    const documentXml = await readDocxPart(bytes, "word/document.xml");
    const starts = bookmarkStarts(documentXml);

    expect(starts.map((b) => b.name)).toEqual(["one", "two", "three", "four"]);
    expect(starts.map((b) => b.id)).toEqual([1, 2, 3, 4]);
    expect(new Set(starts.map((b) => b.id)).size).toBe(starts.length);
  });

  it("pairs every start with an end carrying the same id", async () => {
    const bytes = await convert("# one\n\n## two\n\n### three\n");
    const documentXml = await readDocxPart(bytes, "word/document.xml");

    expect(bookmarkEnds(documentXml)).toEqual(bookmarkStarts(documentXml).map((b) => b.id));
  });

  it("allocates from one counter with no gaps, so a later anchor cannot collide", async () => {
    // The property a Phase 2 `TOC` field or `REF` cross-reference depends on:
    // ids are dense from 1, so the allocator's next value is free by
    // construction rather than by luck.
    const bytes = await convert(
      Array.from({ length: 12 }, (_, index) => `## heading ${index}`).join("\n\n"),
    );
    const ids = bookmarkStarts(await readDocxPart(bytes, "word/document.xml")).map((b) => b.id);

    expect(ids).toEqual(Array.from({ length: 12 }, (_, index) => index + 1));
  });

  it("still resolves an internal link by name, not by id", async () => {
    const bytes = await convert("[jump](#target)\n\n## target\n");
    const documentXml = await readDocxPart(bytes, "word/document.xml");

    // Matched on the attribute rather than on the whole start tag: docx writes
    // `CT_Hyperlink`'s optional `w:history="1"` first and its attribute order is
    // not ours to choose (and means nothing in XML). What must hold is that the
    // link addresses the bookmark by **name** - `w:anchor`, which survives the
    // id renumbering above - and not through a relationship id, which is how an
    // *external* link is written and would leave the document.
    const startTag = /<w:hyperlink\b[^>]*>/.exec(documentXml)?.[0] ?? "";
    expect(startTag).toContain('w:anchor="target"');
    expect(startTag).not.toContain("r:id=");
    expect(bookmarkStarts(documentXml)[0]?.name).toBe("target");
  });

  it("gives two headings that slug-collide one anchor each", async () => {
    const bytes = await convert("## a b\n\n## a-b\n");
    const starts = bookmarkStarts(await readDocxPart(bytes, "word/document.xml"));

    expect(new Set(starts.map((b) => b.name)).size).toBe(2);
    expect(new Set(starts.map((b) => b.id)).size).toBe(2);
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 8 - w:ind w:left was never clamped to the text column     */
/* -------------------------------------------------------------------------- */

/** The largest `w:ind w:left` anywhere in a document. */
function maxIndent(documentXml: string): number {
  return Math.max(
    0,
    ...[...documentXml.matchAll(/<w:ind[^/>]*w:left="(\d+)"/g)].map((m) => Number(m[1])),
  );
}

/** A4's text column, less the one list-indent of headroom the clamp keeps. */
const INDENT_LIMIT = 9026 - DEFAULT_THEME.spacing.listIndent;

describe("review 2, finding 8: an indent cannot start past the right margin", () => {
  it("clamps fourteen levels of block quote to the text column", async () => {
    // Measured before the fix: `w:ind w:left="10080"`, 1054 twips past the
    // *right* margin, where Word draws a column zero characters wide.
    const parts = await renderParts(doc([{ ...blockquote([paragraph([text("deep")])]) }]), {});
    expect(parts.document).toBeDefined();

    let inner: BlockNode = paragraph([text("deep")]);
    for (let level = 0; level < 14; level += 1) inner = blockquote([inner]);
    const deep = await render([inner]);

    expect(maxIndent(deep.document)).toBeLessThanOrEqual(INDENT_LIMIT);
    expect(maxIndent(deep.document)).toBe(INDENT_LIMIT);
  });

  it("clamps a twenty-deep list, whose own warning promised it kept indenting", async () => {
    const deep = await render([nestedBullets(20, "deepest")]);

    expect(maxIndent(deep.document)).toBeLessThanOrEqual(INDENT_LIMIT);
  });

  it("reports it once per document, not once per paragraph", async () => {
    let inner: BlockNode = paragraph([text("deep")]);
    for (let level = 0; level < 20; level += 1) inner = blockquote([inner]);
    const parts = await render([inner, paragraph([text("after")])]);

    const clamped = parts.warnings.filter((w) => w.code === "indent-clamped");
    expect(clamped).toHaveLength(1);
    expect(clamped[0]?.severity).toBe("notice");
  });

  it("says nothing, and moves nothing, for nesting anyone would actually write", async () => {
    const parts = await render([
      blockquote([blockquote([paragraph([text("two deep")])])]),
      nestedBullets(4, "four deep"),
    ]);

    expect(parts.warnings.filter((w) => w.code === "indent-clamped")).toEqual([]);
    expect(maxIndent(parts.document)).toBeLessThan(INDENT_LIMIT);
  });

  it("clamps a table's own indent too", async () => {
    let inner: BlockNode = tinyTable("deep", 2);
    for (let level = 0; level < 14; level += 1) inner = blockquote([inner]);
    const parts = await render([inner]);

    const tableIndent = Number(
      /<w:tblInd w:w="(\d+)"/.exec(parts.document)?.[1] ??
        /<w:tblInd w:type="dxa" w:w="(\d+)"/.exec(parts.document)?.[1] ??
        "0",
    );
    expect(tableIndent).toBeLessThanOrEqual(INDENT_LIMIT);
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 9 - no RTL support of any kind, and no warning            */
/* -------------------------------------------------------------------------- */

const ARABIC = "مرحبا بالعالم";
const HEBREW = "שלום עולם";

describe("review 2, finding 9: right-to-left is expressible, and its absence is reported", () => {
  it("sets the base direction on paragraphs and runs", async () => {
    const parts = await renderParts(
      doc([heading(1, [text(HEBREW)], { id: "h" }), paragraph([text(ARABIC)])]),
      { direction: "rtl" },
    );

    expect(paragraphWithText(parts.document, ARABIC)).toContain("<w:bidi/>");
    expect(paragraphWithText(parts.document, ARABIC)).toContain("<w:rtl/>");
    expect(paragraphWithText(parts.document, HEBREW)).toContain("<w:bidi/>");
    expect(parts.warnings.filter((w) => w.code === "rtl-not-enabled")).toEqual([]);
  });

  it("mirrors a table's column order", async () => {
    const parts = await renderParts(doc([tinyTable(ARABIC, 2)]), { direction: "rtl" });

    expect(parts.document).toContain("<w:bidiVisual/>");
  });

  it("puts the list marker on the right", async () => {
    const parts = await renderParts(doc([bulletList([listItem([paragraph([text(ARABIC)])])])]), {
      direction: "rtl",
    });

    expect(paragraphWithText(parts.document, ARABIC)).toContain("<w:bidi/>");
  });

  it("leaves code, raw HTML and display math left-to-right", async () => {
    // A right-to-left *document* does not make a shell transcript
    // right-to-left: `<w:bidi/>` there would reverse every line's reading order.
    const parts = await renderParts(
      doc([
        codeBlock("const x = 1;\n", { lang: "js" }),
        htmlBlock("<div>x</div>\n"),
        mathBlock("x = 1"),
      ]),
      { direction: "rtl" },
    );

    for (const p of paragraphs(parts.document)) {
      if (pStyle(p) === "CodeBlock" || pStyle(p) === "HtmlBlock") {
        expect(p, `${pStyle(p) ?? "?"} carries <w:bidi/>`).not.toContain("<w:bidi/>");
      }
    }
  });

  it("warns exactly once when right-to-left prose lands in a left-to-right document", async () => {
    const parts = await renderParts(
      doc([
        paragraph([text(ARABIC)]),
        paragraph([text(HEBREW)]),
        bulletList([listItem([paragraph([text(ARABIC)])])]),
      ]),
    );

    const rtl = parts.warnings.filter((w) => w.code === "rtl-not-enabled");
    expect(rtl).toHaveLength(1);
    expect(rtl[0]?.severity).toBe("notice");
    expect(parts.document).not.toContain("<w:bidi/>");
  });

  it("says nothing for a Latin document", async () => {
    const parts = await renderParts(kitchenSink());

    expect(parts.warnings.filter((w) => w.code === "rtl-not-enabled")).toEqual([]);
  });

  it("reaches a caller of convert(), which can also set the direction", async () => {
    const warnings: ConvertWarning[] = [];
    await convert(`${ARABIC}\n`, { onWarning: (warning) => warnings.push(warning) });
    expect(warnings.filter((w) => w.code === "rtl-not-enabled")).toHaveLength(1);

    const quiet: ConvertWarning[] = [];
    const bytes = await convert(`${ARABIC}\n`, {
      direction: "rtl",
      onWarning: (warning) => quiet.push(warning),
    });
    expect(quiet.filter((w) => w.code === "rtl-not-enabled")).toEqual([]);
    expect(await readDocxPart(bytes, "word/document.xml")).toContain("<w:bidi/>");
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 10 - renderDocument's options got no validation at all    */
/* -------------------------------------------------------------------------- */

describe("review 2, finding 10: renderDocument validates its options too", () => {
  const model = doc([codeBlock("\tindented\n", { lang: "text" })]);

  it("rejects a tabSize that is an out-of-memory crash, not a tab stop", () => {
    // Measured before the fix: a bare `RangeError: Invalid string length` out
    // of `" ".repeat(tabSize)`, which is exactly the failure options.ts's own
    // docblock promises never to allow.
    let thrown: unknown;
    try {
      renderDocument(model, { tabSize: 1e9 });
    } catch (error: unknown) {
      thrown = error;
    }

    expect(isDownwordError(thrown)).toBe(true);
    expect(thrown).toMatchObject({ code: "invalid-options" });
    expect((thrown as DownwordError).message).toContain("tabSize");
  });

  it.each([
    ["a fractional tabSize", { tabSize: 2.5 }],
    ["a negative tabSize", { tabSize: -1 }],
    ["an unknown html policy", { html: "sanitise" }],
    ["an unknown softBreak policy", { softBreak: "newline" }],
    ["an unknown direction", { direction: "sideways" }],
    ["a non-boolean titleBlock", { titleBlock: "yes" }],
    ["a 1x1 page", { page: { size: { width: 1, height: 1 } } }],
    ["a page bigger than Word allows", { page: { size: { width: 999999, height: 100 } } }],
    ["margins that consume the page", { page: { margin: { left: 6000, right: 6000 } } }],
    ["a negative margin", { page: { margin: { top: -1 } } }],
    ["an unknown paper size", { page: { size: "Foolscap" } }],
    ["an unknown orientation", { page: { orientation: "sideways" } }],
  ])("rejects %s with invalid-options", (_label, options) => {
    expect(() => renderDocument(model, options as never)).toThrow(
      expect.objectContaining({ code: "invalid-options" }) as Error,
    );
  });

  it("accepts the whole documented range of tabSize", () => {
    for (const tabSize of [0, 1, 4, 64]) {
      expect(() => renderDocument(model, { tabSize })).not.toThrow();
    }
  });

  it("surfaces tabSize and direction on ConvertOptions, validated identically", async () => {
    const expanded = await readDocxPart(
      await convert("```\n\tx\n```\n", { tabSize: 2 }),
      "word/document.xml",
    );
    expect(textOf(expanded)).toBe("  x");

    await expect(convert("# hi", { tabSize: 1e9 })).rejects.toMatchObject({
      code: "invalid-options",
    });
    await expect(
      convert("# hi", { direction: "sideways" } as unknown as ConvertOptions),
    ).rejects.toMatchObject({ code: "invalid-options" });
  });

  it("still produces a document for every value it accepts", async () => {
    const parts = await renderParts(model, {
      page: { size: "Letter", orientation: "landscape", margin: { top: 720, left: 720 } },
      direction: "rtl",
      tabSize: 0,
      html: "drop",
      softBreak: "ignore",
      titleBlock: false,
    });

    expect(parts.bytes.byteLength).toBeGreaterThan(0);
  });
});

/* -------------------------------------------------------------------------- */
/* Review 2, finding 11 - no default style, and 23 dangling w:basedOn          */
/* -------------------------------------------------------------------------- */

describe("review 2, finding 11: every basedOn resolves, and body text has a style", () => {
  it("declares Normal and DefaultParagraphFont", async () => {
    const parts = await render([paragraph([text("body")])]);

    expect(styleIds(parts.styles)).toContain("Normal");
    expect(styleIds(parts.styles)).toContain("DefaultParagraphFont");
  });

  it("has no dangling basedOn anywhere in styles.xml", async () => {
    // Measured before the fix: 16 x basedOn="Normal" and 7 x
    // basedOn="DefaultParagraphFont", neither of which was declared.
    const parts = await renderParts(kitchenSink(), { toc: false });
    const declared = new Set(styleIds(parts.styles));
    const referenced = [...parts.styles.matchAll(/<w:basedOn w:val="([^"]*)"\/>/g)].map(
      (m) => m[1] ?? "",
    );

    expect(referenced.length).toBeGreaterThan(20);
    expect([...new Set(referenced)].filter((id) => !declared.has(id))).toEqual([]);
  });

  it("resolves every `next` reference too, for the same reason", async () => {
    const parts = await render([paragraph([text("body")])]);
    const declared = new Set(styleIds(parts.styles));
    const referenced = [...parts.styles.matchAll(/<w:next w:val="([^"]*)"\/>/g)].map(
      (m) => m[1] ?? "",
    );

    expect([...new Set(referenced)].filter((id) => !declared.has(id))).toEqual([]);
  });

  it("puts a pStyle on ordinary body text, so Design -> Style Set can reach it", async () => {
    const parts = await render([
      paragraph([text("body")]),
      blockquote([paragraph([text("quoted")])]),
    ]);

    expect(pStyle(paragraphWithText(parts.document, "body"))).toBe("Normal");
    // A quote keeps its own style; it was never the unreachable one.
    expect(pStyle(paragraphWithText(parts.document, "quoted"))).toBe("Quote");
  });

  it("keeps Normal free of properties, so docDefaults stays the single source", async () => {
    const parts = await render([paragraph([text("body")])]);
    const normal = styleById(parts.styles, "Normal") ?? "";

    expect(normal).not.toContain("<w:pPr>");
    expect(normal).not.toContain("<w:rPr>");
  });
});
