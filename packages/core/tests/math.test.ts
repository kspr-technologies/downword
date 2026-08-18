import { fileURLToPath } from "node:url";

import { ImportedXmlComponent } from "docx";
import { build, type BuildOptions, type Metafile } from "esbuild";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { convert } from "../src/convert.js";
import { createMathConverter } from "../src/math/convert.js";
import { convertDocumentMath } from "../src/math/document.js";
import type { MathEngine, MathEngineLoader } from "../src/math/engine.js";
import { estimateMathBox, wrapMathmlInSvg } from "../src/math/image.js";
import { mathMarkdownIt } from "../src/math/markdown-it.js";
import { checkOmml, repairOmml } from "../src/math/omml.js";
import { resolveMathOptions } from "../src/math/options.js";
import {
  MATH_WARNING_CODES,
  mathWarningSeverity,
  createMathWarningSink,
  type MathWarning,
  type MathWarningCode,
} from "../src/math/warnings.js";
import {
  doc,
  mathBlock,
  mathInline,
  nodesOfType,
  paragraph,
  text,
  type DocumentNode,
  type ResolvedRasterImage,
} from "../src/model.js";
import { parseMarkdown } from "../src/parse/index.js";
import { renderDocument } from "../src/render/index.js";
import type { ImageRasterizer } from "../src/images/types.js";
import { convertToDocumentWithMath, convertWithMath, mathPlugin } from "../src/plugins/math.js";
import { packDocument, readDocxPart, writeFixture } from "./helpers/docx.js";
import { paragraphs, pStyle, textOf } from "./helpers/xml.js";

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

/** Parses markdown with the math tokenizer attached, and nothing else new. */
function parse(markdown: string): DocumentNode {
  return parseMarkdown(markdown, { plugins: [mathMarkdownIt()] });
}

/** The TeX of every equation a source produces, inline then display. */
function equations(markdown: string): { inline: string[]; block: string[] } {
  const parsed = parse(markdown);
  return {
    inline: nodesOfType(parsed, "mathInline").map((node) => node.value),
    block: nodesOfType(parsed, "mathBlock").map((node) => node.value),
  };
}

interface Converted {
  readonly document: string;
  readonly warnings: readonly MathWarning[];
  readonly bytes: Uint8Array;
}

/** The whole pipeline, for a source, with the real engines. */
async function run(
  markdown: string,
  options: Parameters<typeof convertWithMath>[1] = {},
): Promise<Converted> {
  const warnings: MathWarning[] = [];
  const bytes = await convertWithMath(markdown, {
    ...options,
    math: { ...options.math, onWarning: (warning) => warnings.push(warning) },
  });
  return { document: await readDocxPart(bytes, "word/document.xml"), warnings, bytes };
}

/** Every `<m:oMath>` element in a document part. */
function ommlElements(documentXml: string): string[] {
  return [...documentXml.matchAll(/<m:oMath[\s\S]*?<\/m:oMath>/g)].map((match) => match[0]);
}

/* -------------------------------------------------------------------------- */
/* Delimiters                                                                  */
/* -------------------------------------------------------------------------- */

describe("mathMarkdownIt: inline delimiters", () => {
  it("recognises $…$", () => {
    expect(equations("The area is $\\pi r^2$ exactly.").inline).toEqual(["\\pi r^2"]);
  });

  it("leaves prices alone", () => {
    // The closing candidate is preceded by a space *and* followed by a digit;
    // either one on its own is enough to rule it out.
    expect(equations("A dollar sign that is not math: it costs $5 and $10.")).toEqual({
      inline: [],
      block: [],
    });
  });

  it("requires the delimiters to hug their content", () => {
    expect(equations("$ x$ and $x $ and $$").inline).toEqual([]);
  });

  it("does not treat an escaped dollar as a delimiter", () => {
    expect(equations("\\$5 and \\$10").inline).toEqual([]);
    // The escape belongs to the closing dollar, so the equation runs on to the
    // next real one.
    expect(equations("$a\\$b$").inline).toEqual(["a\\$b"]);
  });

  it("leaves an unterminated $ as text", () => {
    expect(equations("an unterminated $x here").inline).toEqual([]);
    expect(
      nodesOfType(parse("an unterminated $x here"), "text")
        .map((n) => n.value)
        .join(""),
    ).toBe("an unterminated $x here");
  });

  it("carries marks onto math inside emphasis, and finds math in links", () => {
    const parsed = parse("**$E = mc^2$** and [see $\\pi$](https://example.com)");
    const [bold, linked] = nodesOfType(parsed, "mathInline");
    expect(bold?.value).toBe("E = mc^2");
    expect(bold?.marks.map((mark) => mark.type)).toEqual(["bold"]);
    expect(linked?.marks.map((mark) => mark.type)).toEqual(["link"]);
  });

  it("finds math in list items and table cells", () => {
    const parsed = parse(
      ["1. Compute $z = Wx + b$", "", "| a | b |", "| - | - |", "| $\\eta$ | rate |"].join("\n"),
    );
    expect(nodesOfType(parsed, "mathInline").map((node) => node.value)).toEqual([
      "z = Wx + b",
      "\\eta",
    ]);
  });

  it("reads $$…$$ inside a paragraph as inline double-dollar maths", () => {
    expect(equations("before $$x^2$$ after")).toEqual({ inline: ["x^2"], block: [] });
  });
});

describe("mathMarkdownIt: display delimiters", () => {
  it("recognises a fenced $$ block", () => {
    expect(equations("Display:\n\n$$\n\\frac{1}{N}\n$$\n").block).toEqual(["\\frac{1}{N}"]);
  });

  it("recognises a one-line $$…$$ block", () => {
    expect(equations("$$E = mc^2$$").block).toEqual(["E = mc^2"]);
  });

  it("closes an unterminated $$ at the end of the document rather than throwing", () => {
    expect(equations("$$\n\\frac{a}{b}\nand then some prose").block).toEqual([
      "\\frac{a}{b}\nand then some prose",
    ]);
  });

  it("treats an indented $$ as code, not maths", () => {
    expect(equations("    $$\n    x\n    $$").block).toEqual([]);
  });

  it("finds display maths inside a blockquote and a list item", () => {
    expect(equations("> $$\n> x^2\n> $$").block).toEqual(["x^2"]);
    expect(equations("- $$y^2$$").block).toEqual(["y^2"]);
  });

  it("lets $$ interrupt a paragraph", () => {
    expect(equations("prose\n$$x$$\nmore").block).toEqual(["x"]);
  });
});

/* -------------------------------------------------------------------------- */
/* OMML: the headline feature                                                  */
/* -------------------------------------------------------------------------- */

describe("math: 'omml' produces editable Word equations", () => {
  it("emits real m:oMath, not a picture, for inline and display maths", async () => {
    const { document, warnings } = await run("Inline $x^2$.\n\n$$\n\\frac{a}{b}\n$$\n");

    expect(ommlElements(document)).toHaveLength(2);
    expect(document).not.toContain("<w:drawing>");
    expect(document).not.toContain("<undefined");
    expect(warnings).toEqual([]);
  });

  it("emits the OMML element each construct needs", async () => {
    // The five element kinds that prove this is a real equation tree and not a
    // string of characters: fraction, superscript, n-ary, radical, matrix.
    const cases: ReadonlyArray<readonly [string, string]> = [
      ["$\\frac{a}{b}$", "<m:f>"],
      ["$x^2$", "<m:sSup>"],
      ["$x_i$", "<m:sSub>"],
      ["$x_i^2$", "<m:sSubSup>"],
      ["$\\sum_{i=1}^{n} i$", "<m:nary>"],
      ["$\\int_0^\\infty e^{-x}dx$", "<m:nary>"],
      ["$\\sqrt{x+1}$", "<m:rad>"],
      ["$\\begin{matrix} 1 & 2 \\\\ 3 & 4 \\end{matrix}$", "<m:m>"],
      // temml spells an overline as `menclose notation="top"` and an accent as
      // a plain `mover`, so these land on m:borderBox and m:limUpp rather than
      // OMML's m:bar/m:acc. Both are real equation structure either way.
      ["$\\overline{AB}$", "<m:borderBox>"],
      ["$\\hat{x}$", "<m:limUpp>"],
      ["$\\widehat{xy}$", "<m:groupChr>"],
      // Limits go under the operator only in display style.
      ["$$\\lim_{x \\to 0} f(x)$$", "<m:limLow>"],
    ];

    for (const [source, element] of cases) {
      const { document, warnings } = await run(source);
      expect(document, `${source} should contain ${element}`).toContain(element);
      expect(warnings.filter((warning) => warning.severity === "error")).toEqual([]);
    }
  });

  it("keeps matrices, which the docx Math* builders cannot express at all", async () => {
    const { document } = await run("$$\n\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}\n$$");
    const [omml] = ommlElements(document);

    expect(omml).toContain("<m:m>");
    // Two rows, two cells each, and the parentheses kept as run text.
    expect([...(omml ?? "").matchAll(/<m:mr>/g)]).toHaveLength(2);
    expect(omml).toContain('<m:count m:val="2"/>');
    expect(omml).toContain(">(<");
    expect(omml).toContain(">)<");
  });

  it("converts the constructs a real document uses", async () => {
    const source = [
      "Greek: $\\alpha\\beta\\Gamma\\Omega$.",
      "",
      "Stacks: $x_{i}^{2}$ and $x^{y^{z}}$ and $a_{b_{c}}$.",
      "",
      "Text with spaces: $\\text{hello world}$.",
      "",
      "$$\n\\begin{aligned} x &= 1 \\\\ y &= 2 \\end{aligned}\n$$",
      "",
      "$$\n\\begin{cases} 1 & x > 0 \\\\ 0 & \\text{otherwise} \\end{cases}\n$$",
      "",
      "$$\n\\int_a^b f(x)\\,dx = F(b) - F(a)\n$$",
      "",
      "$$\n\\prod_{k=1}^{n} \\frac{k}{k+1}\n$$",
    ].join("\n");

    const { document, warnings } = await run(source);

    expect(ommlElements(document)).toHaveLength(9);
    expect(warnings.filter((warning) => warning.severity === "error")).toEqual([]);
    // \text{} keeps its spaces, as the non-breaking spaces temml emits so that
    // nothing downstream can collapse them.
    expect(document).toContain('<m:t xml:space="preserve">hello\u00a0world</m:t>');
  });

  it("centres display maths and leaves inline maths in the sentence", async () => {
    const { document } = await run("Before $x$ after.\n\n$$\ny\n$$\n");
    const withMath = paragraphs(document).filter((p) => p.includes("<m:oMath"));

    expect(withMath).toHaveLength(2);
    // Inline: the surrounding prose is in the same paragraph.
    expect(textOf(withMath[0] ?? "")).toBe("Before  after.");
    expect(withMath[0]).not.toContain("<w:jc ");
    // Display: its own centred paragraph. mathml2omml never emits m:oMathPara,
    // so the alignment is the paragraph's.
    expect(withMath[1]).toContain('<w:jc w:val="center"/>');
  });

  it("reports how much it converted", async () => {
    const parsed = parse("$x$ and $y$ and $x$\n\n$$z$$");
    const result = await convertDocumentMath(parsed);

    expect(result.ommlCount).toBe(4);
    expect(result.unconvertedCount).toBe(0);
    expect(result.imageCount).toBe(0);
    expect(result.diagnostics).toEqual([]);
  });

  it("converts a broad corpus without a single equation being refused", async () => {
    // The guard against the converter's own drift. `checkOmml` allows the `m:`
    // namespace wholesale but nothing outside it, so a `mathml2omml` release
    // that starts emitting a new attribute or element shows up here as a
    // document full of literal TeX rather than as a silent regression.
    const corpus = [
      "E = mc^2",
      "\\frac{a}{b}",
      "\\sqrt[3]{x}",
      "\\sum_{i=1}^{n} i^2",
      "\\int_a^b f(x)\\,dx",
      "\\prod_{k=1}^n k",
      "\\lim_{x\\to\\infty}\\frac{1}{x}",
      "\\begin{matrix}1&2\\\\3&4\\end{matrix}",
      "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}",
      "\\begin{bmatrix}1\\\\2\\end{bmatrix}",
      "\\begin{vmatrix}a&b\\\\c&d\\end{vmatrix}",
      "\\begin{cases}1&x>0\\\\0&\\text{otherwise}\\end{cases}",
      "\\begin{aligned}a&=b\\\\c&=d\\end{aligned}",
      "\\begin{array}{cc}1&2\\\\3&4\\end{array}",
      "\\alpha\\beta\\gamma\\delta\\Gamma\\Omega",
      "x_{i}^{2}",
      "x^{y^{z}}",
      "\\mathbb{R}\\mathcal{L}\\mathfrak{g}\\mathbf{v}\\mathrm{d}",
      "\\text{hello world}",
      "\\hat{x}\\bar{y}\\vec{z}\\dot{a}\\ddot{b}\\tilde{c}",
      "\\widehat{xy}\\overrightarrow{AB}",
      "\\overline{AB}\\underline{CD}",
      "\\binom{n}{k}",
      "\\frac{\\partial f}{\\partial x}",
      "\\nabla\\cdot\\vec{F}",
      "a\\le b\\ge c\\ne d\\approx e",
      "\\left(\\frac{a}{b}\\right)",
      "\\{x\\mid x>0\\}",
      "\\sin x+\\cos y",
      "\\log_2 n",
      "\\forall x\\exists y",
      "\\xrightarrow{f}",
      "\\overbrace{a+b}^{n}",
      "\\underbrace{x}_{y}",
      "\\substack{a\\\\b}",
      "\\boxed{x}",
      "\\stackrel{?}{=}",
      "f''(x)",
      "10\\%\\ \\#\\ \\_",
      "\\operatorname{argmax}_x f(x)",
    ];

    const source = corpus.map((tex) => `$$\n${tex}\n$$`).join("\n\n");
    const { document, warnings } = await run(source);

    expect(warnings.filter((warning) => warning.code === "omml-rejected")).toEqual([]);
    expect(warnings.filter((warning) => warning.code === "tex-invalid")).toEqual([]);
    expect(ommlElements(document)).toHaveLength(corpus.length);
  });

  it("puts an n-ary operator's operand inside it, so the slot is not an empty box", async () => {
    // MathML makes the integrand a SIBLING of the operator; OMML requires it
    // inside <m:e>. mathml2omml maps the operator faithfully and leaves
    // <m:e/> empty, which LibreOffice draws as a placeholder box and Word
    // shows as an empty slot. Caught by looking at a rendered PDF, not by any
    // structural assertion - the XML was well-formed and full of m:oMath the
    // whole time.
    expect(repairOmml("<m:nary><m:e/></m:nary><m:r><m:t>x</m:t></m:r>")).toBe(
      "<m:nary><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>",
    );

    // Nothing to hoist: left alone rather than corrupted.
    expect(repairOmml("<m:nary><m:e/></m:nary>")).toBe("<m:nary><m:e/></m:nary>");

    for (const tex of [
      "\\int_0^1 x^2 dx = \\frac{1}{3}",
      "\\sum_{k=1}^{n} k^2 = 6",
      "\\int_a^b f(x)dx",
    ]) {
      const { document, warnings } = await run(`$$${tex}$$`);
      expect(warnings).toEqual([]);
      expect(document).toContain("m:nary");
      // The regression itself: not one empty operand anywhere.
      expect(document).not.toContain("<m:e/>");
    }
  });

  it("repairs the one converter bug that would make Word call the file corrupt", async () => {
    // mathml2omml has no mapping for mathvariant="normal" and emits
    // `<m:sty m:val="undefined"/>`, which is outside OMML's ST_Style
    // enumeration. temml uses that variant for every upright letter, so this
    // fires on \Gamma, \Omega and \mathrm{} - i.e. constantly.
    expect(repairOmml('<m:sty m:val="undefined"/>')).toBe('<m:sty m:val="p"/>');
    expect(repairOmml("<m:r/>")).toBe("<m:r/>");

    const { document, warnings } = await run("$\\Gamma$ and $\\mathrm{d}x$ and $\\Omega$");

    expect(warnings).toEqual([]);
    expect(ommlElements(document)).toHaveLength(3);
    expect(document).not.toContain('m:val="undefined"');
    expect(document).toContain('<m:sty m:val="p"/>');
  });

  it("packs into a .docx a real Office implementation can open", async () => {
    const { bytes, document } = await run(
      [
        "# Equations",
        "",
        "The Gaussian integral is $\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}$.",
        "",
        "$$",
        "\\mathcal{L}(\\theta) = \\frac{1}{N} \\sum_{i=1}^{N} \\ell(f_\\theta(x_i), y_i)",
        "$$",
        "",
        "| Symbol | Meaning |",
        "| ------ | ------- |",
        "| $\\eta$ | learning rate |",
        "",
        "$$",
        "A = \\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}",
        "$$",
      ].join("\n"),
    );

    // CI converts everything in this directory with LibreOffice.
    await writeFixture("golden-math.docx", bytes);
    expect(ommlElements(document)).toHaveLength(4);
  });

  it("produces a package whose every part is well-formed XML", async () => {
    // Worth asserting explicitly for maths and for nothing else in this suite:
    // it is the one feature that puts XML *generated by another library*
    // straight into the package, and a single stray "<" would make Word refuse
    // the whole file rather than one equation.
    const { bytes } = await run(
      "$\\frac{a}{b}$ and $\\text{a & b}$\n\n$$\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}$$",
    );
    const zip = await JSZip.loadAsync(bytes);

    const parts = Object.keys(zip.files).filter(
      (name) => name.endsWith(".xml") || name.endsWith(".rels"),
    );
    expect(parts.length).toBeGreaterThan(5);
    for (const name of parts) {
      const xml = await (zip.file(name)?.async("string") ?? Promise.resolve(""));
      expect(() => ImportedXmlComponent.fromXmlString(xml), `${name} must parse`).not.toThrow();
    }
  });
});

/* -------------------------------------------------------------------------- */
/* Malformed TeX                                                               */
/* -------------------------------------------------------------------------- */

describe("malformed TeX never throws", () => {
  const nasty: ReadonlyArray<readonly [string, string]> = [
    ["unbalanced open brace", "$\\frac{a$"],
    ["unbalanced close brace", "$a}$"],
    ["an undefined macro", "$\\undefinedmacro$"],
    ["an empty display block", "$$\n$$"],
    ["a whitespace-only display block", "$$\n   \n$$"],
    ["a lone backslash", "$\\$"],
    ["an unterminated inline delimiter", "an unterminated $x here"],
    ["an environment that does not exist", "$$\\begin{nope}x\\end{nope}$$"],
    ["mismatched environments", "$$\\begin{matrix}x\\end{pmatrix}$$"],
    ["an expansion bomb", "$\\def\\a{\\a}\\a$"],
    ["a \\newcommand bomb", "$\\newcommand{\\x}{\\x\\x}\\x$"],
    ["an untrusted command", "$\\includegraphics{http://example.com/x.png}$"],
    ["a control character", "$a\u0000b$"],
    ["a lone surrogate", "$a\ud800b$"],
    ["raw XML", "$<script>alert(1)</script>$"],
    ["display maths in inline position", "$\\begin{align}a &= b\\end{align}$"],
  ];

  for (const [what, source] of nasty) {
    it(`survives ${what}`, async () => {
      const { document } = await run(source);
      // A document, every time, with the mandatory parts intact.
      expect(document).toContain("<w:body>");
      expect(document).not.toContain("<undefined");
    });
  }

  it("survives a \\frac nested 50 deep, and converts it", async () => {
    let tex = "x";
    for (let depth = 0; depth < 50; depth += 1) tex = `\\frac{1}{${tex}}`;

    const { document, warnings } = await run(`$${tex}$`);
    expect([...document.matchAll(/<m:f>/g)]).toHaveLength(50);
    expect(warnings).toEqual([]);
  });

  it("survives nesting deep enough to exhaust the call stack", async () => {
    let tex = "x";
    for (let depth = 0; depth < 5000; depth += 1) tex = `\\frac{1}{${tex}}`;

    const { document, warnings } = await run(`$${tex}$`);
    // A RangeError from the engine is a failed equation, not a failed document.
    expect(document).toContain("<w:body>");
    expect(warnings.map((warning) => warning.code)).toContain("tex-invalid");
  });

  it("converts a 100 kB equation", async () => {
    const tex = `${"x+".repeat(50_000)}1`;
    expect(tex.length).toBeGreaterThan(100_000);

    const { document, warnings } = await run(`$${tex}$`);
    expect(ommlElements(document)).toHaveLength(1);
    expect(warnings).toEqual([]);
  });

  it("refuses an equation past maxLength without parsing it", async () => {
    const tex = `${"x+".repeat(1000)}1`;
    const { document, warnings } = await run(`$${tex}$`, { math: { maxLength: 100 } });

    expect(ommlElements(document)).toEqual([]);
    expect(warnings.map((warning) => warning.code)).toEqual(["tex-too-large"]);
    expect(warnings[0]?.severity).toBe("error");
  });

  it("leaves a broken equation in the document as its own source text", async () => {
    const { document, warnings } = await run("before $\\frac{a$ after");

    expect(ommlElements(document)).toEqual([]);
    expect(textOf(document)).toContain("\\frac{a");
    expect(warnings.map((warning) => warning.code)).toEqual(["tex-invalid"]);
    expect(warnings[0]?.message).toContain("left as literal text");
  });

  it("warns once per distinct equation, however often it is repeated", async () => {
    const { warnings } = await run("$\\bad{$ and $\\bad{$ and $\\bad{$");
    expect(warnings).toHaveLength(1);
  });

  it("notices a construct OMML has no equivalent for, without failing it", async () => {
    const { document, warnings } = await run("$\\phantom{x}y$");

    expect(ommlElements(document)).toHaveLength(1);
    expect(warnings.map((warning) => warning.code)).toEqual(["tex-unsupported"]);
    expect(warnings[0]?.severity).toBe("notice");
    expect(warnings[0]?.message).toContain("mphantom");
  });

  it("keeps mathml2omml's console.warn off the host's console", async () => {
    const original = console.warn;
    const seen: unknown[] = [];
    console.warn = (...args: unknown[]) => seen.push(args);
    try {
      await run("$\\phantom{x}y$");
    } finally {
      console.warn = original;
    }
    expect(seen).toEqual([]);
    // …and the swap is undone even so.
    expect(console.warn).toBe(original);
  });
});

/* -------------------------------------------------------------------------- */
/* XML injection                                                               */
/* -------------------------------------------------------------------------- */

describe("XML injection cannot reach document.xml", () => {
  const payloads: ReadonlyArray<readonly [string, string]> = [
    ["a script tag", "<script>alert(1)</script>"],
    ["a WordprocessingML paragraph", "<w:p><w:r><w:t>PWNED</w:t></w:r></w:p>"],
    ["a bare element", "<b>x</b>"],
    ["a self-closing break", "\\text{<w:br/>}"],
    ["an external drawing", '\\text{<w:drawing><a:blip r:embed="rId9"/></w:drawing>}'],
    ["an entity", "\\text{&amp;lt;w:p&amp;gt;}"],
  ];

  for (const [what, tex] of payloads) {
    it(`escapes ${what}`, async () => {
      const { document } = await run(`$${tex}$\n\n$$${tex}$$`);

      // Nothing the payload named became an element.
      expect(document).not.toContain("<script");
      expect(document).not.toContain("<b>");
      expect(document).not.toContain("<w:br/>");
      expect(document).not.toContain("<w:drawing>");
      expect(document).not.toContain("PWNED</w:t>");
      // Every paragraph in the body is one we put there: a heading-less two
      // paragraph document stays two paragraphs.
      expect(paragraphs(document).length).toBeLessThanOrEqual(2);
    });
  }

  it("keeps the escaping in the OMML itself", async () => {
    const { document } = await run("$\\text{<w:p>}$");
    const [omml] = ommlElements(document);

    expect(omml).toContain("&lt;w:p&gt;");
    expect(omml).not.toContain("<w:p>");
  });

  it("survives an ampersand, a less-than and a quote in the maths", async () => {
    const { document } = await run('$a \\& b$ and $\\text{\\{x\\}}$ and $\\text{"q"}$');

    expect(ommlElements(document)).toHaveLength(3);
    expect(document).toContain("a&amp;b");
  });

  describe("checkOmml", () => {
    const wrap = (body: string): string =>
      `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${body}</m:oMath>`;

    it("accepts what the converter really produces", () => {
      expect(checkOmml(wrap('<m:r><m:t xml:space="preserve">x</m:t></m:r>'))).toEqual({
        ok: true,
        empty: false,
      });
    });

    it("reports an equation with no content", () => {
      expect(
        checkOmml(
          '<m:oMath xmlns:m="' +
            "http://schemas.openxmlformats.org/officeDocument/2006/math" +
            '"/>',
        ),
      ).toEqual({
        ok: true,
        empty: true,
      });
    });

    const refusals: ReadonlyArray<readonly [string, string, string]> = [
      ["a foreign element", wrap("<script>alert(1)</script>"), "not OMML"],
      ["a WordprocessingML paragraph", wrap("<w:p><w:t>x</w:t></w:p>"), "not OMML"],
      ["an unescaped ampersand", wrap("<m:r><m:t>a&b</m:t></m:r>"), 'unescaped "&"'],
      ["an unescaped less-than", wrap("<m:r><m:t>a<b</m:t></m:r>"), "not OMML"],
      ["an unexpected attribute", wrap('<m:r style="x"/>'), "style attribute"],
      ["a forged namespace", '<m:oMath xmlns:m="http://evil.example/"/>', "not OMML's"],
      ["a comment", wrap("<!-- hi -->"), "comment"],
      ["a processing instruction", wrap("<?xml-stylesheet?>"), "comment, doctype"],
      ["an unclosed element", '<m:oMath xmlns:m="x"><m:r>', "not OMML's"],
      ["a mismatched close", wrap("<m:r></m:e>"), "does not close"],
      ["trailing content", `${wrap("<m:r/>")}<m:r/>`, "after the <m:oMath> root closed"],
      ["the literal string undefined", "undefined", "did not return an <m:oMath>"],
      ["nothing at all", "", "returned nothing"],
    ];

    for (const [what, omml, reason] of refusals) {
      it(`refuses ${what}`, () => {
        const result = checkOmml(omml);
        expect(result.ok).toBe(false);
        expect(result.ok ? "" : result.reason).toContain(reason);
      });
    }

    it("refuses an enumeration value the converter failed to map", () => {
      const result = checkOmml(wrap('<m:r><m:rPr><m:sty m:val="undefined"/></m:rPr></m:r>'));
      expect(result.ok).toBe(false);
      expect(result.ok ? "" : result.reason).toContain('m:val="undefined"');
    });

    it("is linear in the length of its input", () => {
      // A quadratic scanner turns a pathological equation into a hang; this is
      // the shape that would trigger it.
      const long = `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${"<m:r><m:t>x</m:t></m:r>".repeat(20_000)}</m:oMath>`;
      const started = Date.now();
      expect(checkOmml(long).ok).toBe(true);
      expect(Date.now() - started).toBeLessThan(1000);
    });
  });

  it("refuses OMML from an engine that stops escaping, and says so", async () => {
    const hostile: MathEngine = {
      renderToMathml: () => "<math/>",
      mathmlToOmml: () =>
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
        "<w:p><w:r><w:t>PWNED</w:t></w:r></w:p></m:oMath>",
    };

    const warnings: MathWarning[] = [];
    const result = await convertDocumentMath(parse("$x$"), {
      load: () => Promise.resolve(hostile),
      onWarning: (warning) => warnings.push(warning),
    });
    const bytes = await packDocument(renderDocument(result.document));
    const document = await readDocxPart(bytes, "word/document.xml");

    expect(warnings.map((warning) => warning.code)).toEqual(["omml-rejected"]);
    expect(document).not.toContain("PWNED</w:t>");
    expect(textOf(document)).toBe("x");
  });
});

/* -------------------------------------------------------------------------- */
/* Modes                                                                       */
/* -------------------------------------------------------------------------- */

describe("math: 'off'", () => {
  it("leaves the TeX as literal text and never loads an engine", async () => {
    let loaded = 0;
    const load: MathEngineLoader = () => {
      loaded += 1;
      return Promise.reject(new Error("should not be called"));
    };

    const { document } = await run("Inline $x^2$.\n\n$$\n\\frac{a}{b}\n$$\n", {
      math: { math: "off", load },
    });

    expect(loaded).toBe(0);
    expect(ommlElements(document)).toEqual([]);
    expect(textOf(document)).toContain("x^2");
    expect(textOf(document)).toContain("\\frac{a}{b}");
  });

  it("styles the literal TeX as code, which is the renderer's existing fallback", async () => {
    const { document } = await run("$$\nx^2\n$$", { math: { math: "off" } });
    const [block] = paragraphs(document);
    expect(pStyle(block ?? "")).toBe("CodeBlock");
  });
});

describe("math: 'image'", () => {
  const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);

  /** A rasteriser that returns a fixed raster and records what it was asked. */
  function fakeRasterizer(): ImageRasterizer & { readonly calls: string[] } {
    const calls: string[] = [];
    return {
      calls,
      rasterize: (request) => {
        calls.push(new TextDecoder().decode(request.data));
        const raster: ResolvedRasterImage = {
          format: "png",
          data: PNG,
          width: 240,
          height: 60,
        };
        return Promise.resolve(raster);
      },
    };
  }

  it("draws the equation and embeds the SVG with its mandatory raster twin", async () => {
    const rasterizer = fakeRasterizer();
    const { document, warnings } = await run("Inline $x^2$ here.", {
      math: { math: "image", rasterizer },
    });

    expect(warnings).toEqual([]);
    expect(document).toContain("<w:drawing>");
    expect(document).toContain("svgBlip");
    expect(ommlElements(document)).toEqual([]);
    // What the rasteriser was handed: MathML inside an SVG foreignObject.
    expect(rasterizer.calls[0]).toContain("<foreignObject");
    expect(rasterizer.calls[0]).toContain("<math");
  });

  it("puts a display equation in its own paragraph", async () => {
    const { document } = await run("$$\n\\frac{a}{b}\n$$", {
      math: { math: "image", rasterizer: fakeRasterizer() },
    });
    const drawn = paragraphs(document).filter((p) => p.includes("<w:drawing>"));
    expect(drawn).toHaveLength(1);
  });

  it("says so honestly when there is no rasterizer, and writes a real equation instead", async () => {
    const { document, warnings } = await run("$x^2$", { math: { math: "image" } });

    expect(warnings.map((warning) => warning.code)).toEqual(["image-no-rasterizer"]);
    expect(warnings[0]?.message).toContain("browser-only");
    expect(warnings[0]?.severity).toBe("notice");
    // The equation is still in the document, in the better format.
    expect(ommlElements(document)).toHaveLength(1);
  });

  it("falls back to OMML when the rasterizer declines or throws", async () => {
    const declines: ImageRasterizer = { rasterize: () => Promise.resolve(null) };
    const throws: ImageRasterizer = {
      rasterize: () => {
        throw new Error("no canvas here");
      },
    };

    for (const rasterizer of [declines, throws]) {
      const { document, warnings } = await run("$x^2$", { math: { math: "image", rasterizer } });
      expect(warnings.map((warning) => warning.code)).toEqual(["image-failed"]);
      expect(ommlElements(document)).toHaveLength(1);
    }
  });

  it("rejects a raster OOXML cannot hold", async () => {
    const rasterizer: ImageRasterizer = {
      rasterize: () =>
        Promise.resolve({
          format: "png",
          data: new Uint8Array(),
          width: 0,
          height: 0,
        } as ResolvedRasterImage),
    };

    const { document, warnings } = await run("$x$", { math: { math: "image", rasterizer } });
    expect(warnings.map((warning) => warning.code)).toEqual(["image-failed"]);
    expect(ommlElements(document)).toHaveLength(1);
  });

  it("takes its shape from the rasterizer and its scale from the font size", () => {
    const mathml = "<math><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";
    const box = estimateMathBox(mathml, 16);

    // One line plus one fraction: taller than a plain line, never zero.
    expect(box.height).toBeGreaterThan(16);
    expect(box.width).toBeGreaterThan(0);
    expect(wrapMathmlInSvg(mathml, box.width, box.height)).toContain(
      `width="${box.width}" height="${box.height}"`,
    );
  });
});

/* -------------------------------------------------------------------------- */
/* The engine                                                                  */
/* -------------------------------------------------------------------------- */

describe("engine failures", () => {
  it("degrades the whole document to literal TeX when the engines are missing", async () => {
    const warnings: MathWarning[] = [];
    const load: MathEngineLoader = () => Promise.reject(new Error("Cannot find package 'temml'"));

    const parsed = parse("$x^2$ and $y$");
    const result = await convertDocumentMath(parsed, {
      load,
      onWarning: (warning) => warnings.push(warning),
    });

    expect(warnings.map((warning) => warning.code)).toEqual(["engine-unavailable"]);
    expect(warnings[0]?.message).toContain("npm i temml mathml2omml");
    // Untouched, so the renderer's own fallback prints the TeX source.
    expect(result.document).toBe(parsed);
    expect(result.unconvertedCount).toBe(2);
  });

  it("rejects a loader that returns something that is not an engine", async () => {
    const warnings: MathWarning[] = [];
    const converter = await createMathConverter(
      resolveMathOptions({ load: () => Promise.resolve({} as unknown as MathEngine) }),
      createMathWarningSink((warning) => warnings.push(warning)),
    );

    expect(converter).toBeNull();
    expect(warnings[0]?.code).toBe("engine-unavailable");
  });

  it("does not import temml when the document has no maths", async () => {
    let loaded = 0;
    const result = await convertDocumentMath(parse("# Just prose"), {
      load: () => {
        loaded += 1;
        return Promise.reject(new Error("should not be called"));
      },
    });
    expect(loaded).toBe(0);
    expect(result.diagnostics).toEqual([]);
  });

  it("applies macros, freshly, to every equation", async () => {
    const { document } = await run("$\\RR$ and $\\RR$", {
      math: { macros: { "\\RR": "\\mathbb{R}" } },
    });
    expect(ommlElements(document)).toHaveLength(2);
    expect(document).toContain("ℝ");
  });

  it("does not let one equation's \\newcommand leak into the next", async () => {
    const { warnings } = await run("$\\newcommand{\\zz}{z}\\zz$ and $\\zz$");
    // The second equation must not see the first one's definition.
    expect(warnings.map((warning) => warning.code)).toEqual(["tex-invalid"]);
    expect(warnings[0]?.message).toContain("\\zz");
  });
});

/* -------------------------------------------------------------------------- */
/* Options                                                                     */
/* -------------------------------------------------------------------------- */

describe("options", () => {
  it("defaults to omml", () => {
    expect(resolveMathOptions().math).toBe("omml");
    expect(resolveMathOptions({}).delimiters).toBe("dollars");
  });

  const bad: ReadonlyArray<readonly [string, Record<string, unknown>]> = [
    ["an unknown mode", { math: "svg" }],
    ["a negative maxLength", { maxLength: -1 }],
    ["a fractional maxLength", { maxLength: 1.5 }],
    ["a font size out of range", { fontSize: 0 }],
    ["macros that are not an object", { macros: "\\RR" }],
    ["a macro body that is not a string", { macros: { "\\RR": 3 } }],
    ["a rasterizer with no rasterize()", { rasterizer: {} }],
    ["a loader that is not a function", { load: "temml" }],
    ["a handler that is not a function", { onWarning: true }],
  ];

  for (const [what, options] of bad) {
    it(`rejects ${what}`, () => {
      expect(() => resolveMathOptions(options)).toThrow(/invalid|must be/i);
    });
  }

  it("validates when the plugin is built, not at the first equation", () => {
    expect(() => mathPlugin({ math: "nope" as "omml" })).toThrow(/must be one of/);
  });

  it("has a severity for every warning code", () => {
    const codes: readonly MathWarningCode[] = MATH_WARNING_CODES;
    expect(new Set(codes).size).toBe(codes.length);
    for (const code of codes) {
      expect(["error", "notice"]).toContain(mathWarningSeverity(code));
    }
  });

  it("swallows a warning handler that throws", () => {
    const sink = createMathWarningSink(() => {
      throw new Error("bad logger");
    });
    expect(() => sink.report("tex-invalid", "x", false, "boom")).not.toThrow();
  });
});

/* -------------------------------------------------------------------------- */
/* The document rewrite                                                        */
/* -------------------------------------------------------------------------- */

describe("convertDocumentMath: the rewrite", () => {
  it("returns the very same document when there is no maths in it", async () => {
    const source = doc([paragraph([text("no maths here")])]);
    const result = await convertDocumentMath(source);
    expect(result.document).toBe(source);
  });

  it("rebuilds only the path to each equation", async () => {
    const untouched = paragraph([text("prose")]);
    const source = doc([untouched, paragraph([mathInline("x")]), mathBlock("y")]);

    const { document } = await convertDocumentMath(source);

    expect(document).not.toBe(source);
    // Identity matters: prepareHighlights and resolveDocumentImages key their
    // maps by node.
    expect(document.children[0]).toBe(untouched);
    expect(document.metadata).toBe(source.metadata);
  });

  it("fills equations nested in lists, quotes, tables and footnotes", async () => {
    const parsed = parse(
      [
        "- item with $a$",
        "",
        "> quoted $b$",
        "",
        "| $c$ |",
        "| --- |",
        "| $d$ |",
        "",
        "note[^1]",
        "",
        "[^1]: footnote with $e$",
      ].join("\n"),
    );

    const { document, ommlCount } = await convertDocumentMath(parsed);
    expect(ommlCount).toBe(5);
    expect(nodesOfType(document, "mathInline").every((node) => node.omml !== null)).toBe(true);
  });

  it("does not mutate the document it was given", async () => {
    const parsed = parse("$x$");
    const before = JSON.stringify(parsed);
    await convertDocumentMath(parsed);
    expect(JSON.stringify(parsed)).toBe(before);
  });
});

/* -------------------------------------------------------------------------- */
/* convertWithMath                                                             */
/* -------------------------------------------------------------------------- */

describe("convertWithMath", () => {
  it("produces exactly what convert() does when there is no maths", async () => {
    const source = [
      "# Title",
      "",
      "Some **prose**, a [link](https://example.com) and a list:",
      "",
      "- one",
      "- two",
      "",
      "```ts",
      "const x = 1;",
      "```",
    ].join("\n");

    // docx mints a random relationship id per hyperlink, so `convert()` is not
    // byte-reproducible against itself either; everything else must match.
    const stableIds = (xml: string): string => xml.replace(/rId[A-Za-z0-9_-]{8,}/g, "rIdN");
    const [plain, withMath] = await Promise.all([
      convert(source).then((bytes) => readDocxPart(bytes, "word/document.xml")),
      convertWithMath(source).then((bytes) => readDocxPart(bytes, "word/document.xml")),
    ]);

    // The proof that the extra stage is `convert()` plus maths, and not a fork
    // that has drifted.
    expect(stableIds(withMath)).toBe(stableIds(plain));
  });

  it("honours the ordinary convert options", async () => {
    const bytes = await convertWithMath("$x$", {
      metadata: { title: "Equations" },
      theme: "print",
    });
    expect(await readDocxPart(bytes, "docProps/core.xml")).toContain("Equations");
  });

  it("keeps the caller's markdown-it plugins", async () => {
    let used = 0;
    await convertWithMath("$x$", {
      plugins: [
        () => {
          used += 1;
        },
      ],
    });
    expect(used).toBe(1);
  });

  it("rejects a non-string source with a DownwordError, not a TypeError", async () => {
    await expect(convertToDocumentWithMath(undefined as unknown as string)).rejects.toMatchObject({
      code: "invalid-input",
    });
  });

  it("is unharmed when the caller also passes the math tokenizer itself", async () => {
    // An easy mistake: `convertWithMath` appends the rule, and the caller adds
    // it again through `plugins`. Two identically-named markdown-it rules are
    // tried in turn and the first one consumes the delimiter, so the result is
    // the same document rather than doubled or duplicated equations.
    const { document, warnings } = await run("$x^2$ and $\\frac{a}{b}$", {
      plugins: [mathPlugin().markdownIt],
    });

    expect(warnings).toEqual([]);
    expect(ommlElements(document)).toHaveLength(2);
    expect(textOf(document)).toBe(" and ");
  });

  it("exposes the plugin object", async () => {
    const math = mathPlugin();
    const parsed = parseMarkdown("$x^2$", { plugins: [math.markdownIt] });
    const { document } = await math.convertDocument(parsed);
    expect(nodesOfType(document, "mathInline")[0]?.omml).toContain("<m:sSup>");
  });
});

/* -------------------------------------------------------------------------- */
/* Bundle discipline                                                           */
/* -------------------------------------------------------------------------- */

const PACKAGE_DIR = fileURLToPath(new URL("..", import.meta.url));

async function bundle(entry: string, options: BuildOptions = {}): Promise<Metafile> {
  const result = await build({
    absWorkingDir: PACKAGE_DIR,
    entryPoints: [entry],
    bundle: true,
    format: "esm",
    target: "es2022",
    platform: "browser",
    write: false,
    metafile: true,
    logLevel: "silent",
    define: { __DOWNWORD_VERSION__: '"0.0.0-test"' },
    ...options,
  });
  return result.metafile as Metafile;
}

describe("@ksprtech/downword/plugins/math: the engines load lazily or not at all", () => {
  it("keeps temml and mathml2omml out of the main entry", async () => {
    const metafile = await bundle("src/index.ts");
    const inputs = Object.keys(metafile.inputs).join("\n");

    expect(inputs).not.toContain("temml");
    expect(inputs).not.toContain("mathml2omml");
    expect(inputs).not.toContain("src/math/");
    expect(inputs).not.toContain("src/plugins/");
  });

  it("imports them dynamically, and only from engine.ts", async () => {
    const result = await build({
      absWorkingDir: PACKAGE_DIR,
      entryPoints: ["src/plugins/math.ts"],
      bundle: true,
      format: "esm",
      target: "es2022",
      platform: "browser",
      packages: "external",
      write: false,
      metafile: true,
      logLevel: "silent",
      define: { __DOWNWORD_VERSION__: '"0.0.0-test"' },
    });
    const text = result.outputFiles?.[0]?.text ?? "";

    const staticImports = [...text.matchAll(/(?:^|\n)import\s[^;]*?from\s*["']([^"']+)["']/g)].map(
      (match) => match[1],
    );
    const dynamicImports = [...text.matchAll(/import\(\s*["']([^"']+)["']\s*\)/g)].map(
      (match) => match[1],
    );

    expect(staticImports).not.toContain("temml");
    expect(staticImports).not.toContain("mathml2omml");
    expect(dynamicImports).toContain("temml");
    expect(dynamicImports).toContain("mathml2omml");
  });

  it("is browser-safe: no Node builtins anywhere in the entry", async () => {
    const result = await build({
      absWorkingDir: PACKAGE_DIR,
      entryPoints: ["src/plugins/math.ts"],
      bundle: true,
      format: "esm",
      target: "es2022",
      platform: "browser",
      packages: "external",
      write: false,
      metafile: true,
      logLevel: "silent",
      define: { __DOWNWORD_VERSION__: '"0.0.0-test"' },
    });
    const output = Object.values(result.metafile.outputs)[0];
    const externals = (output?.imports ?? [])
      .filter((entry) => entry.external)
      .map((entry) => entry.path);

    expect(externals.filter((path) => path.startsWith("node:"))).toEqual([]);
    expect([...new Set(externals)].sort()).toEqual([
      "docx",
      "markdown-it",
      "markdown-it-footnote",
      "mathml2omml",
      "temml",
    ]);
  });
});
