/**
 * T2.4 — table of contents, core properties, page numbers, heading anchors.
 *
 * ## What a TOC test can honestly assert
 *
 * A `TOC` field has **no entries in the file**. OOXML stores the instruction
 * (`TOC \o "1-3" \h \u \z`) and the word processor computes the entries when
 * the field is updated, so nothing downword writes can be checked by looking
 * for the heading text inside the field. What *can* be checked, and is:
 *
 *  - the field is there, inside its `w:sdt`, with the right switches;
 *  - its `begin` `fldChar` carries `w:dirty="true"`, so a reader knows the
 *    cached result is stale;
 *  - `word/settings.xml` carries `<w:updateFields/>`, which is what makes Word
 *    offer to update on open;
 *  - the heading above it is `TOCHeading`, whose outline level 9 keeps it out
 *    of the list it introduces;
 *  - the headings the field will collect carry real `HeadingN` styles with the
 *    outline levels the switches name.
 *
 * Whether a reader then *fills it in* is the reader's behaviour, not the file's.
 * The repo's LibreOffice gate (`scripts/docx-validity.mjs`) converts every
 * generated fixture to PDF, which proves the document opens and lays out — it
 * does **not** run `Tools -> Update -> Indexes and Tables`, so it does not
 * prove the entries appear. No test here claims otherwise.
 */

import { describe, expect, it } from "vitest";

import { convert, type ConvertWarning } from "../src/index.js";
import { doc, heading, link, metadata, paragraph, text } from "../src/model.js";
import { parseMarkdown } from "../src/parse/index.js";
import { readDocxPart, renderParts, writeFixture } from "./helpers/docx.js";
import { anchorDocument, minimalDocument, outlineDocument } from "./helpers/fixtures.js";
import {
  bookmarkStarts,
  fieldInstructions,
  internalAnchors,
  paragraphs,
  pStyle,
  styleById,
  styleIds,
  textOf,
} from "./helpers/xml.js";

/* -------------------------------------------------------------------------- */
/* Table of contents                                                           */
/* -------------------------------------------------------------------------- */

describe("table of contents", () => {
  it("is off unless asked for", async () => {
    const parts = await renderParts(outlineDocument());

    expect(fieldInstructions(parts.document)).toEqual([]);
    expect(parts.document).not.toContain("<w:sdt>");
    // No field to update, so nothing may prompt the reader on open.
    expect(parts.settings).not.toContain("<w:updateFields/>");
  });

  it("emits a native TOC field with Word's own default switches", async () => {
    const parts = await renderParts(outlineDocument(), { toc: true });
    await writeFixture("golden-render-toc.docx", parts.bytes);

    // \o "1-3": collect outline levels 1-3. \h: hyperlink each entry.
    // \u: include paragraphs carrying an outline level directly.
    // \z: hide the tab and page number in web layout.
    expect(fieldInstructions(parts.document)).toEqual(['TOC \\h \\o "1-3" \\u \\z']);
    expect(parts.document).toContain("<w:sdt>");
  });

  it("marks the field dirty and asks the reader to update it", async () => {
    const parts = await renderParts(outlineDocument(), { toc: true });

    expect(parts.document).toContain('<w:fldChar w:fldCharType="begin" w:dirty="true"/>');
    expect(parts.settings).toContain("<w:updateFields/>");
  });

  it("warns that the field is empty until it is updated, and says how", async () => {
    const parts = await renderParts(outlineDocument(), { toc: true });

    const warning = parts.warnings.find((w) => w.code === "toc-needs-update");
    expect(warning).toBeDefined();
    expect(warning?.severity).toBe("notice");
    expect(warning?.message).toContain("Update Field");
    expect(warning?.message).toContain("LibreOffice");
  });

  it("honours an explicit heading-level range", async () => {
    const parts = await renderParts(outlineDocument(), { toc: { minLevel: 2, maxLevel: 4 } });

    expect(fieldInstructions(parts.document)).toEqual(['TOC \\h \\o "2-4" \\u \\z']);
  });

  it("prints a TOCHeading paragraph above the field by default", async () => {
    const parts = await renderParts(outlineDocument(), { toc: true });
    const contents = paragraphs(parts.document).find((p) => textOf(p) === "Contents");

    expect(contents).toBeDefined();
    expect(pStyle(contents ?? "")).toBe("TOCHeading");
  });

  it("keeps the contents heading out of the contents", async () => {
    const parts = await renderParts(outlineDocument(), { toc: true });
    const style = styleById(parts.styles, "TOCHeading") ?? "";

    // Outline level 9 is OOXML's "body text": without it the style's `basedOn`
    // Heading1 would put "Contents" at level 0 and the field would list itself.
    expect(style).toContain('<w:basedOn w:val="Heading1"/>');
    expect(style).toContain('<w:outlineLvl w:val="9"/>');
  });

  it("takes a custom title, or none at all", async () => {
    const named = await renderParts(outlineDocument(), { toc: { title: "Table of contents" } });
    expect(paragraphs(named.document).map(textOf)).toContain("Table of contents");

    const bare = await renderParts(outlineDocument(), { toc: { title: null } });
    expect(paragraphs(bare.document).map(pStyle)).not.toContain("TOCHeading");
    // The field itself is still there; only its heading is gone.
    expect(fieldInstructions(bare.document)).toEqual(['TOC \\h \\o "1-3" \\u \\z']);
  });

  it("sits below the title block and above the body", async () => {
    const parts = await renderParts(outlineDocument(), { toc: true, titleBlock: true });
    const styles = paragraphs(parts.document).map(pStyle);

    expect(styles[0]).toBe("Title");
    expect(styles[1]).toBe("TOCHeading");
    // Then the field's own two paragraphs, then the first real heading.
    expect(styles.indexOf("Heading1")).toBeGreaterThan(1);
    expect(textOf(paragraphs(parts.document)[0] ?? "")).toBe("Outline Fixture");
  });

  it("refuses a range that is upside down or out of bounds", async () => {
    await expect(
      renderParts(minimalDocument(), { toc: { minLevel: 4, maxLevel: 2 } }),
    ).rejects.toThrow(
      /options\.toc\.minLevel \(4\) must not be deeper than options\.toc\.maxLevel \(2\)/,
    );
    await expect(renderParts(minimalDocument(), { toc: { maxLevel: 7 } })).rejects.toThrow(
      /options\.toc\.maxLevel must be an integer between 1 and 6/,
    );
    await expect(
      renderParts(minimalDocument(), { toc: { title: 7 as unknown as string } }),
    ).rejects.toThrow(/options\.toc\.title must be a string or null/);
    await expect(
      renderParts(minimalDocument(), { toc: "yes" as unknown as boolean }),
    ).rejects.toThrow(/options\.toc must be a boolean or an object/);
  });
});

/* -------------------------------------------------------------------------- */
/* Heading outline levels                                                      */
/* -------------------------------------------------------------------------- */

describe("heading outline levels", () => {
  it("gives every heading level the outline level the navigation pane reads", async () => {
    const parts = await renderParts(outlineDocument());

    for (let level = 1; level <= 6; level += 1) {
      const style = styleById(parts.styles, `Heading${level}`) ?? "";
      expect(style).toContain(`<w:outlineLvl w:val="${level - 1}"/>`);
    }
  });

  it("references the built-in heading style ids, not direct formatting", async () => {
    const parts = await renderParts(outlineDocument());
    const styles = paragraphs(parts.document).map(pStyle);

    expect(styles).toContain("Heading1");
    expect(styles).toContain("Heading2");
    expect(styles).toContain("Heading3");
    expect(styles).toContain("Heading4");
  });

  it("keeps a heading below the TOC range in the outline anyway", async () => {
    // `\o "1-3"` excludes it from the *field*, but it is still a Heading4 with
    // outline level 3, so the Navigation pane and a wider range both find it.
    const parts = await renderParts(outlineDocument(), { toc: true });
    const deep = paragraphs(parts.document).find(
      (p) => textOf(p) === "Too deep for the default range",
    );

    expect(pStyle(deep ?? "")).toBe("Heading4");
  });
});

/* -------------------------------------------------------------------------- */
/* Heading bookmarks and intra-document links                                  */
/* -------------------------------------------------------------------------- */

describe("heading bookmarks", () => {
  it("bookmarks every heading with a unique w:id", async () => {
    const parts = await renderParts(anchorDocument());
    const marks = bookmarkStarts(parts.document);

    expect(marks).toHaveLength(4);
    expect(new Set(marks.map((m) => m.id)).size).toBe(marks.length);
    expect(new Set(marks.map((m) => m.name)).size).toBe(marks.length);
  });

  it("resolves a link to a duplicate-slug heading to the right one", async () => {
    const parts = await renderParts(anchorDocument());
    const marks = bookmarkStarts(parts.document);
    const headings = paragraphs(parts.document).filter((p) => pStyle(p) === "Heading2");

    // Three headings all reading "Setup"; the slugger gave them setup,
    // setup-1, setup-2 and each carries its own bookmark, in order.
    expect(marks.map((m) => m.name)).toEqual(["setup", "setup_1", "setup_2", "caf_"]);
    for (const [index, name] of ["setup", "setup_1", "setup_2", "caf_"].entries()) {
      expect(headings[index]).toContain(`w:name="${name}"`);
    }
    // And each link points at the bookmark of the heading it names.
    expect(internalAnchors(parts.document)).toEqual(["setup", "setup_1", "setup_2", "caf_"]);
  });

  it("resolves a percent-encoded anchor to its heading", async () => {
    const parts = await renderParts(anchorDocument());

    // markdown-it normalises every destination through `mdurl.encode`, so
    // `[x](#café)` reaches the renderer as `#caf%C3%A9` while the heading's
    // slug is still `café`. Without decoding, no internal link to a heading
    // containing a non-ASCII letter would ever resolve.
    expect(internalAnchors(parts.document)).toContain("caf_");
    expect(textOf(parts.document)).toContain("to the accented one");
  });

  it("degrades a link to a nonexistent anchor to plain text, and says so", async () => {
    const parts = await renderParts(anchorDocument());

    const warning = parts.warnings.find((w) => w.code === "link-unresolved");
    expect(warning?.message).toContain("#no-such-heading");
    // No hyperlink at all: not a blue underlined run that does nothing, and
    // certainly not a `w:anchor` pointing at a bookmark that is not there.
    expect(parts.document).not.toContain('w:anchor="no_such_heading"');
    expect(internalAnchors(parts.document)).toHaveLength(4);
    expect(textOf(parts.document)).toContain("to nothing at all");
  });

  it("agrees with the parser's slugger end to end", async () => {
    const parts = await renderParts(
      parseMarkdown(
        "[one](#setup) [two](#setup-1) [three](#accents-café)\n\n" +
          "## Setup\n\n## Setup\n\n## Accents café\n",
      ),
    );

    expect(internalAnchors(parts.document)).toEqual(["setup", "setup_1", "accents_caf_"]);
    expect(bookmarkStarts(parts.document).map((m) => m.name)).toEqual([
      "setup",
      "setup_1",
      "accents_caf_",
    ]);
    expect(parts.warnings.filter((w) => w.code === "link-unresolved")).toHaveLength(0);
  });

  it("keeps two slugs apart that collapse to the same Word bookmark name", async () => {
    const parts = await renderParts(
      doc([
        paragraph([text("to a-b", [link("#a-b")])]),
        paragraph([text("to a_b", [link("#a_b")])]),
        heading(2, [text("A-B")], { id: "a-b" }),
        heading(2, [text("A_B")], { id: "a_b" }),
      ]),
    );

    // Word bookmark names admit no `-`, so `a-b` and `a_b` would collide; both
    // the heading and the link go through the same table, so they cannot.
    expect(bookmarkStarts(parts.document).map((m) => m.name)).toEqual(["a_b", "a_b_2"]);
    expect(internalAnchors(parts.document)).toEqual(["a_b", "a_b_2"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Page numbers                                                                */
/* -------------------------------------------------------------------------- */

describe("page numbers", () => {
  it("adds no footer unless asked for", async () => {
    const parts = await renderParts(minimalDocument());

    expect(parts.footer).toBeNull();
    expect(parts.document).not.toContain("<w:footerReference");
  });

  it("puts a PAGE field in a footer the section references", async () => {
    const parts = await renderParts(outlineDocument(), { pageNumbers: true });
    await writeFixture("golden-render-page-numbers.docx", parts.bytes);

    expect(parts.document).toContain('<w:footerReference w:type="default"');
    expect(fieldInstructions(parts.footer ?? "")).toEqual(["PAGE"]);
  });

  it("needs no update step, unlike a TOC field", async () => {
    const parts = await renderParts(outlineDocument(), { pageNumbers: true });

    // PAGE and NUMPAGES are computed during layout by every reader that
    // paginates, so nothing may prompt the reader here.
    expect(parts.settings).not.toContain("<w:updateFields/>");
    expect(parts.warnings.map((w) => w.code)).not.toContain("toc-needs-update");
  });

  it("words the number the way the format asks", async () => {
    const bare = await renderParts(minimalDocument(), { pageNumbers: { format: "number" } });
    expect(fieldInstructions(bare.footer ?? "")).toEqual(["PAGE"]);
    expect(textOf(bare.footer ?? "")).toBe("");

    const pageX = await renderParts(minimalDocument(), { pageNumbers: { format: "page-x" } });
    expect(fieldInstructions(pageX.footer ?? "")).toEqual(["PAGE"]);
    expect(textOf(pageX.footer ?? "")).toBe("Page ");

    const ofY = await renderParts(minimalDocument(), { pageNumbers: { format: "page-x-of-y" } });
    expect(fieldInstructions(ofY.footer ?? "")).toEqual(["PAGE", "NUMPAGES"]);
    expect(textOf(ofY.footer ?? "")).toBe("Page  of ");
  });

  it("centres the footer by default and honours an explicit side", async () => {
    const centred = await renderParts(minimalDocument(), { pageNumbers: true });
    expect(centred.footer).toContain('<w:jc w:val="center"/>');

    const right = await renderParts(minimalDocument(), { pageNumbers: { alignment: "right" } });
    expect(right.footer).toContain('<w:jc w:val="right"/>');

    const left = await renderParts(minimalDocument(), { pageNumbers: { alignment: "left" } });
    expect(left.footer).toContain('<w:jc w:val="left"/>');
  });

  it("styles the footer paragraph with Word's own Footer style", async () => {
    const parts = await renderParts(minimalDocument(), { pageNumbers: true });

    expect(parts.footer).toContain('<w:pStyle w:val="Footer"/>');
    expect(styleIds(parts.styles)).toContain("Footer");
  });

  it("refuses a format or an alignment it cannot write", async () => {
    await expect(
      renderParts(minimalDocument(), {
        pageNumbers: { format: "roman" as unknown as "number" },
      }),
    ).rejects.toThrow(/options\.pageNumbers\.format must be one of/);
    await expect(
      renderParts(minimalDocument(), {
        pageNumbers: { alignment: "justify" as unknown as "left" },
      }),
    ).rejects.toThrow(/options\.pageNumbers\.alignment must be one of/);
    await expect(
      renderParts(minimalDocument(), { pageNumbers: 1 as unknown as boolean }),
    ).rejects.toThrow(/options\.pageNumbers must be a boolean or an object/);
  });
});

/* -------------------------------------------------------------------------- */
/* Core properties                                                             */
/* -------------------------------------------------------------------------- */

describe("core properties", () => {
  it("writes the model's metadata into docProps/core.xml", async () => {
    const parts = await renderParts(outlineDocument(), { subject: "Testing" });
    const core = parts.coreProperties ?? "";

    // The exact element names Word's File -> Info panel reads. Verified against
    // docx 9.7.1's IPropertiesOptions: title, subject, creator, keywords,
    // description, lastModifiedBy.
    expect(core).toContain("<dc:title>Outline Fixture</dc:title>");
    expect(core).toContain("<dc:creator>downword</dc:creator>");
    expect(core).toContain("<cp:lastModifiedBy>downword</cp:lastModifiedBy>");
    expect(core).toContain("<dc:subject>Testing</dc:subject>");
    expect(core).toContain("<cp:keywords>toc, outline</cp:keywords>");
    expect(core).toContain(
      "<dc:description>Generated by packages/core/tests/document-features.test.ts</dc:description>",
    );
  });

  it("omits the author rather than claiming Word's Un-named default", async () => {
    const parts = await renderParts(doc([paragraph([text("x")])], metadata({ title: "T" })));
    const core = parts.coreProperties ?? "";

    expect(core).toContain("<dc:title>T</dc:title>");
    expect(core).not.toContain("<dc:creator>Un-named</dc:creator>");
  });

  it("writes an authored date and anything unmapped as custom properties", async () => {
    const parts = await renderParts(outlineDocument());
    const custom = parts.customProperties ?? "";

    // OOXML core properties have no "date the author wrote this" that Word does
    // not rewrite, so it becomes a custom property instead of being dropped.
    expect(custom).toContain('name="Date"');
    expect(custom).toContain("<vt:lpwstr>2026-08-17</vt:lpwstr>");
    expect(custom).toContain('name="Fixture"');
    expect(custom).toContain("<vt:lpwstr>outline</vt:lpwstr>");
  });

  it("sanitizes metadata that XML cannot hold", async () => {
    const parts = await renderParts(
      doc([paragraph([text("x")])], metadata({ title: "bad\u000Btitle" })),
      { subject: "bad\u000Bsubject" },
    );

    // docProps is a separate part with the same Char production as the body: a
    // single U+000B in a pasted title makes Word reject the whole package.
    expect(parts.coreProperties).toContain("<dc:title>bad\uFFFDtitle</dc:title>");
    expect(parts.coreProperties).toContain("<dc:subject>bad\uFFFDsubject</dc:subject>");
    expect(parts.warnings.map((w) => w.code)).toContain("unrepresentable-character");
  });
});

/* -------------------------------------------------------------------------- */
/* The three together                                                          */
/* -------------------------------------------------------------------------- */

describe("toc + page numbers + metadata", () => {
  it("compose without interfering", async () => {
    const parts = await renderParts(outlineDocument(), {
      toc: { minLevel: 1, maxLevel: 2, title: "Contents" },
      pageNumbers: { format: "page-x-of-y", alignment: "center" },
      titleBlock: true,
      subject: "Everything at once",
    });
    await writeFixture("golden-render-niceties.docx", parts.bytes);

    expect(fieldInstructions(parts.document)).toEqual(['TOC \\h \\o "1-2" \\u \\z']);
    expect(fieldInstructions(parts.footer ?? "")).toEqual(["PAGE", "NUMPAGES"]);
    expect(parts.settings).toContain("<w:updateFields/>");
    expect(parts.coreProperties).toContain("<dc:subject>Everything at once</dc:subject>");
    expect(paragraphs(parts.document).map(pStyle).slice(0, 2)).toEqual(["Title", "TOCHeading"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Reachable from convert()                                                    */
/* -------------------------------------------------------------------------- */

describe("convert()", () => {
  const source = "# One\n\ntext[^n]\n\n## Two\n\n[^n]: the note\n";

  async function convertParts(
    options: Parameters<typeof convert>[1],
  ): Promise<{ document: string; footer: string | null; warnings: ConvertWarning[] }> {
    const warnings: ConvertWarning[] = [];
    const bytes = await convert(source, { ...options, onWarning: (w) => warnings.push(w) });
    let footer: string | null = null;
    try {
      footer = await readDocxPart(bytes, "word/footer1.xml");
    } catch {
      footer = null;
    }
    return { document: await readDocxPart(bytes, "word/document.xml"), footer, warnings };
  }

  it("passes toc through, warning included", async () => {
    const { document, warnings } = await convertParts({ toc: { minLevel: 1, maxLevel: 2 } });

    expect(fieldInstructions(document)).toEqual(['TOC \\h \\o "1-2" \\u \\z']);
    expect(warnings.map((w) => w.code)).toContain("toc-needs-update");
    expect(warnings.find((w) => w.code === "toc-needs-update")?.stage).toBe("render");
  });

  it("passes pageNumbers through", async () => {
    const { footer } = await convertParts({ pageNumbers: { format: "page-x" } });

    expect(fieldInstructions(footer ?? "")).toEqual(["PAGE"]);
  });

  it("adds neither by default", async () => {
    const { document, footer } = await convertParts({});

    expect(fieldInstructions(document)).toEqual([]);
    expect(footer).toBeNull();
  });

  it('splits footnotes: "inline" across the parser and the renderer', async () => {
    const notes = await convertParts({});
    expect(textOf(notes.document)).toBe("OnetextTwo");
    expect(notes.document).toContain("<w:footnoteReference");

    const inline = await convertParts({ footnotes: "inline" });
    expect(textOf(inline.document)).toBe("Onetext (the note)Two");
    expect(inline.document).not.toContain("<w:footnoteReference");

    // `false` is the third answer, and a different one: the syntax is not
    // parsed at all, so `[^n]` and its definition survive as literal text.
    const off = await convertParts({ footnotes: false });
    expect(textOf(off.document)).toContain("[^n]");
    expect(off.document).not.toContain("<w:footnoteReference");
  });

  it("refuses a footnotes value it has no meaning for", async () => {
    await expect(convert(source, { footnotes: "maybe" as unknown as boolean })).rejects.toThrow(
      /options\.footnotes must be true, false or "inline"/,
    );
  });
});
