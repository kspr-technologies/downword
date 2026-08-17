/**
 * T2.1 — real Word footnotes.
 *
 * The thing under test is a *pair* of parts, not one: a `[^1]` in the body is
 * `<w:footnoteReference w:id="1"/>` in `word/document.xml`, the note itself is
 * `<w:footnote w:id="1">` in `word/footnotes.xml`, and the document is only
 * correct when every id in the first is present in the second. Two failure
 * modes are therefore checked everywhere below: a reference into an id that is
 * not in the footnotes part (a reference into nothing), and a definition whose
 * text also leaks into the body flow (the note printed twice).
 *
 * Three of the cases here cannot be reached through `parseMarkdown` at all,
 * because `markdown-it-footnote` normalises them away before a token exists —
 * see `footnoteEdgeCases` in `helpers/fixtures.ts`. They are still the
 * renderer's problem: `renderDocument` is exported and documented as one half
 * of the pipeline, so a consumer-built model reaches it having been through no
 * normalisation whatsoever.
 */

import { describe, expect, it } from "vitest";

import {
  blockquote,
  bulletList,
  codeBlock,
  doc,
  footnoteDefinition,
  footnoteReference,
  heading,
  inlineCode,
  link,
  listItem,
  paragraph,
  table,
  tableCell,
  tableRow,
  text,
  type DocumentNode,
} from "../src/model.js";
import { parseMarkdown } from "../src/parse/index.js";
import { renderParts } from "./helpers/docx.js";
import { footnoteEdgeCases } from "./helpers/fixtures.js";
import {
  footnoteById,
  footnoteIds,
  footnoteReferenceIds,
  paragraphs,
  pStyle,
  rStyles,
  textOf,
} from "./helpers/xml.js";

/** The simplest possible shape: one marker, one note. */
function oneFootnote(): DocumentNode {
  return doc([
    paragraph([text("A claim"), footnoteReference("1", 1, { label: "1" }), text(".")]),
    footnoteDefinition("1", 1, [paragraph([text("the note")])], { label: "1" }),
  ]);
}

/* -------------------------------------------------------------------------- */
/* The parts, and the link between them                                        */
/* -------------------------------------------------------------------------- */

describe("footnotes: the two parts", () => {
  it("writes the note to word/footnotes.xml and never to the body", async () => {
    const parts = await renderParts(oneFootnote());

    expect(parts.footnotes).not.toBeNull();
    expect(parts.footnotes).toContain("the note");
    // The definition is hoisted, so its text must not appear in the body flow.
    expect(textOf(parts.document)).not.toContain("the note");
    expect(textOf(parts.document)).toBe("A claim.");
  });

  it("uses a real w:footnoteReference, not a superscript look-alike", async () => {
    const parts = await renderParts(oneFootnote());

    expect(parts.document).toContain('<w:rStyle w:val="FootnoteReference"/>');
    expect(parts.document).toContain('<w:footnoteReference w:id="1"/>');
  });

  it("gives every reference in the body a footnote with the same id", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    const referenced = footnoteReferenceIds(parts.document);
    const defined = footnoteIds(parts.footnotes ?? "");

    expect(referenced.length).toBeGreaterThan(0);
    for (const id of referenced) expect(defined).toContain(id);
  });

  it("styles footnote bodies with the built-in FootnoteText style", async () => {
    const parts = await renderParts(oneFootnote());
    const note = footnoteById(parts.footnotes ?? "", 1) ?? "";

    expect(pStyle(paragraphs(note)[0] ?? "")).toBe("FootnoteText");
  });

  it("puts the FootnoteReference character style on the marker in both parts", async () => {
    const parts = await renderParts(oneFootnote());
    const note = footnoteById(parts.footnotes ?? "", 1) ?? "";

    // docx writes the note's own leading `<w:footnoteRef/>` run with the same
    // built-in character style the body marker carries, which is what makes the
    // two superscripts match.
    expect(rStyles(note)).toContain("FootnoteReference");
    expect(note).toContain("<w:footnoteRef/>");
  });

  it("declares both footnote styles in styles.xml", async () => {
    const parts = await renderParts(oneFootnote());

    expect(parts.styles).toContain('w:styleId="FootnoteReference"');
    expect(parts.styles).toContain('w:styleId="FootnoteText"');
  });

  it("writes no footnotes part content for a document with no footnotes", async () => {
    const parts = await renderParts(doc([paragraph([text("nothing to note")])]));

    // docx always emits the part (it carries the two separators); what matters
    // is that we contributed no notes to it.
    expect(footnoteIds(parts.footnotes ?? "")).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Referenced twice, never, and not at all                                     */
/* -------------------------------------------------------------------------- */

describe("footnotes: reference counting", () => {
  it("draws one marker per reference and still only one note", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    // `twice` is number 1 and is cited from two places.
    expect(footnoteReferenceIds(parts.document).filter((id) => id === 1)).toHaveLength(2);
    expect(footnoteIds(parts.footnotes ?? "").filter((id) => id === 1)).toHaveLength(1);
  });

  it("writes no note that nothing points at, and says whose text went", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    const warning = parts.warnings.find((w) => w.code === "footnote-unreferenced");
    expect(warning?.message).toContain("[^orphan]");
    expect(warning?.severity).toBe("error");

    // Word draws a footnote only where a marker is and drops an unreferenced
    // one on the next save, so writing it would be pretending it survived.
    expect(parts.footnotes).not.toContain("Nothing points at me.");
    expect(footnoteReferenceIds(parts.document)).not.toContain(4);
    expect(footnoteIds(parts.footnotes ?? "")).not.toContain(4);
  });

  it("says nothing about a note reached only from inside another note", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    // `nested` (3) is cited only by `rich` (2), which splices its text in. Its
    // words are in the document, so it is neither reported nor written as a
    // note of its own.
    expect(footnoteIds(parts.footnotes ?? "")).not.toContain(3);
    expect(parts.footnotes).toContain("The inner note.");

    const unreferenced = parts.warnings
      .filter((w) => w.code === "footnote-unreferenced")
      .map((w) => w.message);
    expect(unreferenced.some((message) => message.includes("[^nested]"))).toBe(false);
    expect(unreferenced.some((message) => message.includes("[^orphan]"))).toBe(true);
  });

  it("keeps the two parts a matched pair in both directions", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    const referenced = [...new Set(footnoteReferenceIds(parts.document))].sort((a, b) => a - b);
    const defined = [...footnoteIds(parts.footnotes ?? "")].sort((a, b) => a - b);

    expect(defined).toEqual(referenced);
  });

  it("never emits a reference into a footnote that is not in the part", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    // `[^missing]` is number 9 and has no definition.
    expect(footnoteReferenceIds(parts.document)).not.toContain(9);
    expect(footnoteIds(parts.footnotes ?? "")).not.toContain(9);
  });

  it("draws a dangling marker as superscript text, and says so", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    expect(parts.warnings.map((w) => w.code)).toContain("footnote-unresolved");
    expect(parts.warnings.find((w) => w.code === "footnote-unresolved")?.message).toContain(
      "[^missing]",
    );
    // The label survives on the page, superscripted like a real marker.
    expect(textOf(parts.document)).toContain("missing");
  });
});

/* -------------------------------------------------------------------------- */
/* Rich content                                                                */
/* -------------------------------------------------------------------------- */

describe("footnotes: rich content", () => {
  it("keeps code, links and lists inside the note", async () => {
    const parts = await renderParts(footnoteEdgeCases());
    const rich = footnoteById(parts.footnotes ?? "", 2) ?? "";

    expect(rStyles(rich)).toContain("CodeChar");
    expect(rich).toContain("<w:hyperlink");
    // The bullet list becomes ListParagraph paragraphs with a numbering
    // reference, exactly as it would in the body.
    expect(paragraphs(rich).map(pStyle)).toContain("ListParagraph");
    expect(rich).toContain("<w:numPr>");
    expect(textOf(rich)).toContain("a list item inside a footnote");
    expect(textOf(rich)).toContain("inside();");
  });

  it("registers a footnote list's numbering before numbering.xml is built", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    // The list only exists inside a footnote, so its abstract definition can
    // only be in numbering.xml if the note was rendered before the registry
    // was frozen.
    const rich = footnoteById(parts.footnotes ?? "", 2) ?? "";
    const numId = /<w:numId w:val="(\d+)"\/>/.exec(rich)?.[1];
    expect(numId).toBeDefined();
    expect(parts.numbering).toContain(`<w:num w:numId="${numId ?? ""}">`);
  });

  it("splices a nested marker into the outer note instead of nesting footnotes", async () => {
    const parts = await renderParts(footnoteEdgeCases());
    const rich = footnoteById(parts.footnotes ?? "", 2) ?? "";

    // Word cannot draw a footnote inside a footnote, so the inner note's text
    // joins the outer one in parentheses and no reference is written into the
    // footnotes part at all.
    expect(rich).not.toContain("<w:footnoteReference");
    expect(textOf(rich)).toContain("(The inner note.)");
  });

  it("drops a block a footnote cannot hold, and says which", async () => {
    const parts = await renderParts(
      doc([
        paragraph([footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition(
          "1",
          1,
          [
            paragraph([text("before")]),
            table([tableRow([tableCell([text("nope")])])]),
            paragraph([text("after")]),
          ],
          { label: "1" },
        ),
      ]),
    );

    expect(parts.warnings.map((w) => w.code)).toContain("footnote-content-dropped");
    expect(parts.footnotes).not.toContain("nope");
    expect(parts.footnotes).toContain("before");
    expect(parts.footnotes).toContain("after");
  });

  it("gives an empty definition a paragraph rather than an empty footnote", async () => {
    const parts = await renderParts(
      doc([
        paragraph([footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition("1", 1, [], { label: "1" }),
      ]),
    );

    const note = footnoteById(parts.footnotes ?? "", 1) ?? "";
    expect(paragraphs(note)).toHaveLength(1);
    expect(pStyle(paragraphs(note)[0] ?? "")).toBe("FootnoteText");
  });
});

/* -------------------------------------------------------------------------- */
/* Ordering                                                                    */
/* -------------------------------------------------------------------------- */

describe("footnotes: ordering", () => {
  it("keys notes by number, not by the order the definitions were written", async () => {
    const parts = await renderParts(footnoteEdgeCases());

    // The fixture defines 2 before 1; the ids in the part must still be the
    // model's numbers, and the *reference* order in the body decides what Word
    // prints beside them.
    expect(footnoteIds(parts.footnotes ?? "").sort((a, b) => a - b)).toEqual([1, 2]);
    expect(footnoteById(parts.footnotes ?? "", 1)).toContain("Referenced from two places.");
    expect(footnoteById(parts.footnotes ?? "", 2)).toContain("A note with");
  });

  it("resolves a marker that appears above its definition", async () => {
    const parts = await renderParts(oneFootnote());

    expect(footnoteReferenceIds(parts.document)).toEqual([1]);
    expect(footnoteIds(parts.footnotes ?? "")).toEqual([1]);
  });

  it("keeps the first of two definitions claiming the same number, and says so", async () => {
    const parts = await renderParts(
      doc([
        paragraph([footnoteReference("a", 1, { label: "a" })]),
        footnoteDefinition("a", 1, [paragraph([text("the first")])], { label: "a" }),
        footnoteDefinition("b", 1, [paragraph([text("the second")])], { label: "b" }),
      ]),
    );

    // `Document({ footnotes })` is keyed by id, so only one can survive; the
    // note that is kept must be the one the reference was resolved against.
    expect(footnoteIds(parts.footnotes ?? "")).toEqual([1]);
    expect(parts.footnotes).toContain("the first");
    expect(parts.footnotes).not.toContain("the second");
    expect(parts.warnings.find((w) => w.code === "footnote-content-dropped")?.message).toContain(
      "[^b]",
    );
  });
});

/* -------------------------------------------------------------------------- */
/* footnotes: false                                                            */
/* -------------------------------------------------------------------------- */

describe("footnotes: false", () => {
  it("splices the note into the sentence and writes no footnotes part content", async () => {
    const parts = await renderParts(oneFootnote(), { footnotes: false });

    expect(textOf(parts.document)).toBe("A claim (the note).");
    expect(footnoteReferenceIds(parts.document)).toEqual([]);
    expect(footnoteIds(parts.footnotes ?? "")).toEqual([]);
  });

  it("keeps the marks inside the note", async () => {
    const parts = await renderParts(
      doc([
        paragraph([text("claim"), footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition(
          "1",
          1,
          [
            paragraph([
              text("see "),
              text("npm i", [inlineCode()]),
              text(" at "),
              text("example", [link("https://example.com")]),
            ]),
          ],
          { label: "1" },
        ),
      ]),
      { footnotes: false },
    );

    expect(textOf(parts.document)).toBe("claim (see npm i at example)");
    expect(rStyles(parts.document)).toContain("CodeChar");
    expect(parts.document).toContain("<w:hyperlink");
  });

  it("splices a note twice when it is cited twice", async () => {
    const parts = await renderParts(footnoteEdgeCases(), { footnotes: false });
    const body = textOf(parts.document);

    expect(body.match(/Referenced from two places\./g)).toHaveLength(2);
  });

  it("joins a multi-block note with single spaces", async () => {
    const parts = await renderParts(
      doc([
        paragraph([text("claim"), footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition(
          "1",
          1,
          [
            paragraph([text("first para")]),
            blockquote([paragraph([text("quoted")])]),
            bulletList([listItem([paragraph([text("item")])])]),
            codeBlock("a();\nb();\n"),
          ],
          { label: "1" },
        ),
      ]),
      { footnotes: false },
    );

    expect(textOf(parts.document)).toBe("claim (first para quoted item a(); b();)");
  });

  it("drops an empty note rather than printing an empty parenthesis", async () => {
    const parts = await renderParts(
      doc([
        paragraph([text("claim"), footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition("1", 1, [], { label: "1" }),
      ]),
      { footnotes: false },
    );

    expect(textOf(parts.document)).toBe("claim");
  });

  it("still reports a dangling marker", async () => {
    const parts = await renderParts(footnoteEdgeCases(), { footnotes: false });
    expect(parts.warnings.map((w) => w.code)).toContain("footnote-unresolved");
  });

  it("stops a cycle rather than recursing forever", async () => {
    const parts = await renderParts(
      doc([
        paragraph([text("claim"), footnoteReference("a", 1, { label: "a" })]),
        footnoteDefinition(
          "a",
          1,
          [paragraph([text("A cites B"), footnoteReference("b", 2, { label: "b" })])],
          { label: "a" },
        ),
        footnoteDefinition(
          "b",
          2,
          [paragraph([text("B cites A"), footnoteReference("a", 1, { label: "a" })])],
          { label: "b" },
        ),
      ]),
      { footnotes: false },
    );

    expect(textOf(parts.document)).toBe("claim (A cites B (B cites A))");
    expect(parts.warnings.map((w) => w.code)).toContain("footnote-content-dropped");
  });

  it("rejects a non-boolean footnotes option", async () => {
    await expect(
      renderParts(oneFootnote(), { footnotes: "inline" as unknown as boolean }),
    ).rejects.toThrow(/options\.footnotes must be a boolean/);
  });
});

/* -------------------------------------------------------------------------- */
/* Misplaced definitions                                                       */
/* -------------------------------------------------------------------------- */

describe("footnotes: misplaced definitions", () => {
  it("drops a definition nested inside another block, and says so", async () => {
    const parts = await renderParts(
      doc([
        paragraph([footnoteReference("1", 1, { label: "1" })]),
        blockquote([
          footnoteDefinition("1", 1, [paragraph([text("nested away")])], { label: "1" }),
        ]),
      ]),
    );

    expect(parts.warnings.map((w) => w.code)).toContain("footnote-misplaced");
    expect(textOf(parts.document)).not.toContain("nested away");
    // Only root definitions are hoisted, so the marker now has nothing behind
    // it and must not become a reference into an empty footnote.
    expect(footnoteReferenceIds(parts.document)).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* End to end, through the parser                                              */
/* -------------------------------------------------------------------------- */

describe("footnotes: markdown in, footnotes.xml out", () => {
  it("numbers notes by first reference, not by definition order", async () => {
    const parts = await renderParts(
      parseMarkdown("Beta[^b] then alpha[^a].\n\n[^a]: the a note\n[^b]: the b note\n"),
    );

    // `[^b]` is cited first, so it is footnote 1 even though `[^a]` is defined
    // first — which is what markdown-it-footnote decides and the model carries.
    expect(footnoteReferenceIds(parts.document)).toEqual([1, 2]);
    expect(footnoteById(parts.footnotes ?? "", 1)).toContain("the b note");
    expect(footnoteById(parts.footnotes ?? "", 2)).toContain("the a note");
  });

  it("carries a note containing a fence, a link and a list all the way through", async () => {
    const parts = await renderParts(
      parseMarkdown(
        [
          "Claim[^rich].",
          "",
          "[^rich]: See [the docs](https://example.com):",
          "",
          "    - one",
          "    - two",
          "",
          "    ```js",
          "    go();",
          "    ```",
          "",
        ].join("\n"),
      ),
    );

    const note = footnoteById(parts.footnotes ?? "", 1) ?? "";
    expect(note).toContain("<w:hyperlink");
    expect(textOf(note)).toContain("one");
    expect(textOf(note)).toContain("go();");
    expect(textOf(parts.document)).toBe("Claim.");
  });

  it("reports a definition markdown-it silently discards for want of a reference", () => {
    const warnings: { code: string; message: string; line: number | null }[] = [];
    const document = parseMarkdown("Cited[^a].\n\n[^a]: used\n[^z]: never used\n", {
      onWarning: (warning) => warnings.push(warning),
    });

    // markdown-it-footnote drops an unreferenced definition before a token
    // exists, so the parser is the only place this can ever be noticed.
    expect(warnings.map((w) => w.code)).toContain("footnote-unreferenced");
    expect(warnings.find((w) => w.code === "footnote-unreferenced")?.message).toContain("[^z]");
    expect(document.children.filter((child) => child.type === "footnoteDefinition")).toHaveLength(
      1,
    );
  });

  it("says nothing when every definition is referenced", () => {
    const warnings: { code: string }[] = [];
    parseMarkdown("Cited[^a].\n\n[^a]: used\n", {
      onWarning: (warning) => warnings.push(warning),
    });

    expect(warnings.map((w) => w.code)).not.toContain("footnote-unreferenced");
  });

  it("leaves a marker with no definition as literal text", async () => {
    const document = parseMarkdown("A claim[^nope].\n");
    const parts = await renderParts(document);

    // markdown-it-footnote never produces a `footnote_ref` for a label it has
    // no definition for, so the source survives verbatim and no reference is
    // written. The renderer's own dangling-reference path is exercised by
    // `footnoteEdgeCases`, which is hand-built for exactly that reason.
    expect(textOf(parts.document)).toBe("A claim[^nope].");
    expect(footnoteReferenceIds(parts.document)).toEqual([]);
  });

  it("inlines parsed footnotes when the renderer is told to", async () => {
    const parts = await renderParts(parseMarkdown("A claim[^1].\n\n[^1]: the note\n"), {
      footnotes: false,
    });

    expect(textOf(parts.document)).toBe("A claim (the note).");
    expect(footnoteIds(parts.footnotes ?? "")).toEqual([]);
  });

  it("keeps a heading inside a footnote readable", async () => {
    const parts = await renderParts(
      doc([
        paragraph([footnoteReference("1", 1, { label: "1" })]),
        footnoteDefinition("1", 1, [heading(3, [text("A heading")], { id: "a-heading" })], {
          label: "1",
        }),
      ]),
    );

    const note = footnoteById(parts.footnotes ?? "", 1) ?? "";
    expect(textOf(note)).toContain("A heading");
    expect(pStyle(paragraphs(note)[0] ?? "")).toBe("Heading3");
  });
});
