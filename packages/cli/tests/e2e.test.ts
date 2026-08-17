/**
 * End to end: spawn the built binary, then open what it produced.
 *
 * Every assertion here is against the `.docx` — its parts, its
 * `word/document.xml`, its `docProps` — rather than against an exit code, which
 * only proves the process did not crash. A CLI that exits 0 having written a
 * zip with no `word/document.xml` in it is broken in exactly the way a status
 * check cannot see.
 *
 * Outputs land in `tests/__fixtures__/out/`, which CI's LibreOffice gate
 * (`scripts/docx-validity.mjs`) converts to PDF — so everything written here is
 * also checked for real-world openability, by a real word processor, in CI.
 */

import { spawnSync } from "node:child_process";
import { chmodSync, existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterAll, describe, expect, it } from "vitest";

import { FIXTURES, openDocx, outPath, readDocx, runCli } from "./helpers/cli.js";

/** Scratch directories made by these tests, removed at the end. */
const scratch: string[] = [];

function tempDir(prefix: string): string {
  const dir = mkdtempSync(join(tmpdir(), `downword-cli-${prefix}-`));
  scratch.push(dir);
  return dir;
}

afterAll(() => {
  for (const dir of scratch) rmSync(dir, { recursive: true, force: true });
});

/** `file(1)`'s verdict, or `null` where there is no `file`. */
function fileType(path: string): string | null {
  const result = spawnSync("file", ["--brief", path], { encoding: "utf8" });
  return result.status === 0 ? result.stdout.trim() : null;
}

describe("converting a file", () => {
  it("writes <input>.docx next to the input by default", async () => {
    const dir = tempDir("default-target");
    const input = join(dir, "notes.md");
    writeFileSync(input, "# Hello\n\nFrom **downword**.\n");

    const { status, stderr, stdout } = runCli([input]);

    expect(stderr).toBe("");
    expect(status).toBe(0);
    // Nothing on stdout: the document went to a file.
    expect(stdout.length).toBe(0);

    const docx = await readDocx(join(dir, "notes.docx"));
    expect(docx.parts).toContain("word/document.xml");
    expect(docx.parts).toContain("word/styles.xml");
    expect(docx.text).toContain("Hello");
    expect(docx.text).toContain("From downword.");
  });

  it("renders headings, lists, tables, code and footnotes as Word constructs", async () => {
    const target = outPath("cli-kitchen-sink.docx");
    const { status } = runCli([join(FIXTURES, "kitchen-sink.md"), "-o", target]);
    expect(status).toBe(0);

    const docx = await readDocx(target);

    // A real heading style, not bold text.
    expect(docx.document).toContain('w:val="Heading1"');
    // A real table, not tab stops.
    expect(docx.document).toContain("<w:tbl>");
    // A real numbering definition for the ordered list.
    expect(docx.parts).toContain("word/numbering.xml");
    // Real footnotes, in their own part, with the note text out of the body.
    expect(docx.parts).toContain("word/footnotes.xml");
    expect(docx.part("word/footnotes.xml")).toContain("Footnotes become real Word footnotes.");
    expect(docx.text).not.toContain("Footnotes become real Word footnotes.");
    // Task list glyphs and the code fence survived.
    expect(docx.text).toContain("shipped");
    expect(docx.text).toContain("export const answer: number = 42;");
  });

  it("honours -o and reports nothing when there is nothing to report", () => {
    const target = outPath("cli-explicit-output.docx");
    const { status, stderr } = runCli([join(FIXTURES, "kitchen-sink.md"), "-o", target]);

    expect(stderr).toBe("");
    expect(status).toBe(0);
    expect(existsSync(target)).toBe(true);
  });
});

describe("document options", () => {
  it("--theme academic swaps the styles, not the text", async () => {
    const plain = outPath("cli-theme-default.docx");
    const academic = outPath("cli-theme-academic.docx");
    runCli([join(FIXTURES, "kitchen-sink.md"), "-o", plain]);
    runCli([join(FIXTURES, "kitchen-sink.md"), "-o", academic, "--theme", "academic"]);

    const before = await readDocx(plain);
    const after = await readDocx(academic);

    expect(after.part("word/styles.xml")).toContain("Times New Roman");
    expect(before.part("word/styles.xml")).not.toContain("Times New Roman");
    expect(after.text).toBe(before.text);
  });

  it("--page-size Letter changes the section's paper size", async () => {
    const target = outPath("cli-letter.docx");
    runCli([join(FIXTURES, "kitchen-sink.md"), "-o", target, "--page-size", "Letter"]);

    // Letter is 12240 x 15840 twips; A4 is 11906 x 16838.
    expect((await readDocx(target)).document).toContain('w:w="12240"');
  });

  it("--title and --author land in docProps", async () => {
    const target = outPath("cli-metadata.docx");
    runCli([
      join(FIXTURES, "kitchen-sink.md"),
      "-o",
      target,
      "--title",
      "Q3 report",
      "--author",
      "A. Writer",
    ]);

    const core = (await readDocx(target)).part("docProps/core.xml");
    expect(core).toContain("Q3 report");
    expect(core).toContain("A. Writer");
  });

  it("--toc adds a real TOC field and marks it for updating", async () => {
    const target = outPath("cli-toc.docx");
    const { status, stderr } = runCli([join(FIXTURES, "kitchen-sink.md"), "-o", target, "--toc"]);

    expect(status).toBe(0);
    const docx = await readDocx(target);
    expect(docx.document).toContain("TOC ");
    expect(docx.part("word/settings.xml")).toContain("<w:updateFields");
    // The field is inert until the reader updates it. That is a notice, so the
    // default mode counts it and --verbose spells it out.
    expect(stderr).toContain("1 notice (--verbose to see them)");
  });

  it("--page-numbers adds a footer", async () => {
    const target = outPath("cli-page-numbers.docx");
    runCli([join(FIXTURES, "kitchen-sink.md"), "-o", target, "--page-numbers"]);

    const docx = await readDocx(target);
    expect(docx.parts.some((part) => /^word\/footer\d+\.xml$/.test(part))).toBe(true);
  });

  it("--highlight colours the code, and nothing colours it without", async () => {
    const plain = outPath("cli-plain-code.docx");
    const coloured = outPath("cli-highlighted-code.docx");
    runCli([join(FIXTURES, "kitchen-sink.md"), "-o", plain]);
    const lit = runCli([join(FIXTURES, "kitchen-sink.md"), "-o", coloured, "--highlight"]);

    expect(lit.status).toBe(0);
    // `export` is a keyword: bold, in the theme's keyword ink.
    expect((await readDocx(coloured)).document).toMatch(
      /<w:rPr><w:b\/><w:bCs\/><w:color w:val="[0-9A-F]{6}"\/><\/w:rPr><w:t[^>]*>export</,
    );
    // Without the flag the whole fence is one plain run.
    expect((await readDocx(plain)).document).toContain(
      '<w:t xml:space="preserve">export const answer: number = 42;</w:t>',
    );
  });

  it("--no-footnotes leaves the syntax as text", async () => {
    const target = outPath("cli-no-footnotes.docx");
    runCli([join(FIXTURES, "kitchen-sink.md"), "-o", target, "--no-footnotes"]);

    const docx = await readDocx(target);
    expect(docx.text).toContain("[^note]");
    // `docx` always writes a footnotes part - it holds the separator and
    // continuation notes every document has - so the assertion that matters is
    // that *this document's* note is not in it, and is still in the body.
    expect(docx.part("word/footnotes.xml")).not.toContain("Footnotes become real Word footnotes.");
    expect(docx.text).toContain("Footnotes become real Word footnotes.");
  });
});

describe("maths", () => {
  it("leaves dollars alone by default", async () => {
    const target = outPath("cli-math-off.docx");
    const { status, stderr } = runCli([join(FIXTURES, "math.md"), "-o", target]);

    expect(status).toBe(0);
    expect(stderr).toBe("");
    const docx = await readDocx(target);
    expect(docx.document).not.toContain("<m:oMath");
    expect(docx.text).toContain("$E = mc^2$");
    expect(docx.text).toContain("it costs $5 and $10");
  });

  it("--math omml produces native Word equations", async () => {
    const target = outPath("cli-math-omml.docx");
    const { status } = runCli([join(FIXTURES, "math.md"), "-o", target, "--math", "omml"]);

    expect(status).toBe(0);
    const docx = await readDocx(target);
    expect(docx.document).toContain("<m:oMath");
    // A dollar amount is still a dollar amount.
    expect(docx.text).toContain("it costs $5 and $10");
    // And the equation is an equation, not a picture of one.
    expect(docx.parts.some((part) => part.startsWith("word/media/"))).toBe(false);
  });

  it("--math image says it has no rasteriser instead of pretending", () => {
    const target = outPath("cli-math-image.docx");
    const { status, stderr } = runCli([
      join(FIXTURES, "math.md"),
      "-o",
      target,
      "--math",
      "image",
      "--verbose",
    ]);

    expect(status).toBe(0);
    expect(stderr).toContain("math/image-no-rasterizer");
  });
});

describe("mermaid", () => {
  it("names every fence it skipped, and still produces the document", async () => {
    const target = outPath("cli-mermaid.docx");
    const { status, stderr } = runCli([join(FIXTURES, "mermaid.md"), "-o", target]);

    expect(status).toBe(0);

    // Both fences, by name, with their line numbers, in the default mode.
    expect(stderr).toContain("mermaid.md:3: mermaid/no-dom");
    expect(stderr).toContain('"The request pipeline"');
    expect(stderr).toContain("mermaid.md:10: mermaid/no-dom");
    expect(stderr).toContain('"sequenceDiagram"');
    expect(stderr).toContain("needs a browser DOM");
    expect(stderr).toContain("![](diagram.png)");

    // The diagram source is still in the document, as a code block.
    const docx = await readDocx(target);
    expect(docx.text).toContain("flowchart LR");
    expect(docx.text).toContain("Prose after the diagram.");
  });

  it("says nothing under --quiet", () => {
    const { status, stderr } = runCli([
      join(FIXTURES, "mermaid.md"),
      "-o",
      outPath("cli-mermaid-quiet.docx"),
      "--quiet",
    ]);

    expect(status).toBe(0);
    expect(stderr).toBe("");
  });
});

describe("images", () => {
  it("reads a picture next to the markdown, and refuses one above it", async () => {
    const target = outPath("cli-images.docx");
    const { status, stderr } = runCli([join(FIXTURES, "images.md"), "-o", target]);

    expect(status).toBe(0);
    expect(stderr).toContain("image/path-not-allowed");
    expect(stderr).toContain("outside-pixel.png");
    // The library's advice ("pass { allowOutsideBaseDir: true }") is restated
    // as the flag a person at a prompt can actually type.
    expect(stderr).toContain("--allow-outside-images");

    const docx = await readDocx(target);
    expect(docx.parts.some((part) => part.startsWith("word/media/"))).toBe(true);
  });

  it("--allow-outside-images lets the second one through", async () => {
    const target = outPath("cli-images-outside.docx");
    const { status, stderr } = runCli([
      join(FIXTURES, "images.md"),
      "-o",
      target,
      "--allow-outside-images",
    ]);

    expect(status).toBe(0);
    expect(stderr).toBe("");
    const docx = await readDocx(target);
    expect(docx.parts.filter((part) => part.startsWith("word/media/")).length).toBe(2);
  });

  it("refuses a remote image by default, and says which flag would allow it", () => {
    const dir = tempDir("remote");
    const input = join(dir, "remote.md");
    writeFileSync(input, "# Remote\n\n![logo](https://example.invalid/logo.png)\n");

    const { status, stderr } = runCli([input]);

    expect(status).toBe(0);
    expect(stderr).toContain("image/remote-blocked");
    expect(stderr).toContain("--remote-images");
  });
});

describe("stdin to stdout", () => {
  it("round-trips a document a word processor recognises", async () => {
    const { status, stdout, stderr } = runCli([], {
      input: "# Piped\n\nStraight through the process.\n",
    });

    expect(stderr).toBe("");
    expect(status).toBe(0);

    const target = outPath("cli-stdin-stdout.docx");
    writeFileSync(target, stdout);

    const type = fileType(target);
    if (type !== null) expect(type).toContain("Microsoft Word 2007+");
    // …and independently of `file(1)`, it really is the zip Word expects.
    expect(stdout.subarray(0, 2).toString("latin1")).toBe("PK");
    const docx = await openDocx(stdout);
    expect(docx.parts).toContain("[Content_Types].xml");
    expect(docx.text).toContain("Straight through the process.");
  });

  it("honours -o - as an explicit request for stdout", async () => {
    const { status, stdout } = runCli(["-o", "-", join(FIXTURES, "kitchen-sink.md")]);

    expect(status).toBe(0);
    expect((await openDocx(stdout)).text).toContain("Quarterly report");
  });
});

describe("batch mode", () => {
  it("writes one .docx per input into --outdir", async () => {
    const dir = tempDir("batch");
    const { status, stderr } = runCli(["batch/**/*.md", "--outdir", dir], { cwd: FIXTURES });

    expect(stderr).toBe("");
    expect(status).toBe(0);

    for (const [name, heading] of [
      ["one.docx", "One"],
      ["two.docx", "Two"],
      ["three.docx", "Three"],
    ] as const) {
      expect(existsSync(join(dir, name))).toBe(true);
      expect((await readDocx(join(dir, name))).text).toContain(heading);
    }
    // notes.txt did not match, and nothing else was invented.
    expect(existsSync(join(dir, "notes.docx"))).toBe(false);
  });

  it("creates --outdir when it does not exist", () => {
    const dir = join(tempDir("batch-mkdir"), "deep", "build");
    const { status } = runCli(["batch/*.md", "--outdir", dir], { cwd: FIXTURES });

    expect(status).toBe(0);
    expect(existsSync(join(dir, "one.docx"))).toBe(true);
  });

  it("converts the rest when one document fails, and exits non-zero", () => {
    const dir = tempDir("batch-partial");
    const out = tempDir("batch-partial-out");
    writeFileSync(join(dir, "good.md"), "# Good\n");
    const bad = join(dir, "bad.md");
    writeFileSync(bad, "# Bad\n");
    chmodSync(bad, 0o000);

    try {
      const { status, stderr } = runCli(["*.md", "--outdir", out], { cwd: dir });
      // Root can read anything; there is nothing to fail with.
      if (status === 0) return;

      expect(status).toBe(66);
      expect(stderr).toContain("bad.md");
      // The other document was still converted.
      expect(existsSync(join(out, "good.docx"))).toBe(true);
      expect(existsSync(join(out, "bad.docx"))).toBe(false);
    } finally {
      chmodSync(bad, 0o644);
    }
  });

  it("refuses a batch where two inputs would overwrite each other", () => {
    const dir = tempDir("batch-collision");
    const out = tempDir("batch-collision-out");
    mkdirSync(join(dir, "docs"));
    mkdirSync(join(dir, "guides"));
    writeFileSync(join(dir, "docs", "api.md"), "# Docs API\n");
    writeFileSync(join(dir, "guides", "api.md"), "# Guides API\n");

    const { status, stderr } = runCli(["**/*.md", "--outdir", out], { cwd: dir });

    expect(status).toBe(64);
    expect(stderr).toContain("would both be written to");
    // Nothing was written: the batch stopped before it started.
    expect(existsSync(join(out, "api.docx"))).toBe(false);
  });

  it("counts what it did under --verbose", () => {
    const dir = tempDir("batch-verbose");
    const { status, stderr } = runCli(["batch/*.md", "--outdir", dir, "--verbose"], {
      cwd: FIXTURES,
    });

    expect(status).toBe(0);
    expect(stderr).toContain("one.md ->");
    expect(stderr).toContain("converted 2 of 2 files");
  });
});

describe("the diagnostics channel", () => {
  it("prints what the document lost, with the line it lost it on", () => {
    const dir = tempDir("warnings");
    const input = join(dir, "warnings.md");
    writeFileSync(input, "# Warnings\n\n![gone](./missing.png)\n");

    const { status, stderr } = runCli([input]);

    // A broken image is a warning, not a failure: the document is still made.
    expect(status).toBe(0);
    expect(stderr).toContain("image/not-found");
    expect(existsSync(join(dir, "warnings.docx"))).toBe(true);
  });

  it("counts notices instead of listing them, and lists them under --verbose", () => {
    const target = outPath("cli-notices.docx");
    const args = [join(FIXTURES, "kitchen-sink.md"), "-o", target, "--toc"];

    const normal = runCli(args).stderr;
    expect(normal).toContain("notice (--verbose to see them)");
    expect(normal).not.toContain("toc-needs-update");

    expect(runCli([...args, "--verbose"]).stderr).toContain("render/toc-needs-update");
  });

  it("--quiet prints nothing at all for a document that converted", () => {
    const dir = tempDir("quiet");
    const input = join(dir, "quiet.md");
    writeFileSync(input, "# Quiet\n\n![gone](./missing.png)\n");

    const { status, stderr } = runCli([input, "--quiet"]);

    expect(status).toBe(0);
    expect(stderr).toBe("");
  });

  it("--quiet still prints a fatal error", () => {
    const { status, stderr } = runCli(["no-such-file.md", "--quiet"]);

    expect(status).toBe(66);
    expect(stderr).toContain("no such file");
  });
});
