# downword

**Your LLM answers in markdown. The person who asked for it wants a Word document.**

> markdown → Word (`.docx`), entirely in your browser.
> Paste from ChatGPT/Claude, get a real Word document.

[![CI](https://github.com/kspr-technologies/downword/actions/workflows/ci.yml/badge.svg)](https://github.com/kspr-technologies/downword/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/downword?color=cb3837&logo=npm&label=npm)](https://www.npmjs.com/package/downword)
[![core bundle](https://img.shields.io/badge/core-74.2%20kB%20min%2Bgzip-1f6feb)](packages/core/.size-limit.json)
[![license](https://img.shields.io/badge/license-MIT-1f6feb)](LICENSE)

<!-- ───────────────────────────────────────────────────────────────────────────
     TODO(hero): record the demo GIF. This needs a human — it is a screen
     recording, and no agent can make one.

     What to record (~12 s, no cuts, no cursor jumps):
       1. Open https://ksprtech.com/tools/markdown-to-word in a clean window,
          1280x800, light theme, browser chrome cropped out.
       2. Paste a ChatGPT/Claude answer that contains, in this order: an H1, a
          short paragraph, a bullet list, a pipe table, a fenced ts block, and
          one $…$ equation. The point of the shot is that all six survive.
       3. Pause ~1 s so the rendered preview is legible.
       4. Click Download. Let the browser's download shelf appear.
       5. Cut to Word (or LibreOffice) with the file already open, scrolled to
          show the table and the code block. Open the Styles pane so
          "Heading 1" is visibly selected — that is the whole differentiator in
          one frame.

     How to produce it:
       - Record at 2x DPI, then:
           ffmpeg -i in.mov -vf "fps=15,scale=1200:-1:flags=lanczos,split[a][b];[a]palettegen=stats_mode=diff[p];[b][p]paletteuse=dither=bayer" -loop 0 docs/media/hero.gif
       - Budget: under 4 MB. If it is larger, cut the pause in step 3 first.
       - Also export an MP4 from the same source — GitHub renders <video> in
         READMEs and it is a quarter of the size.
       - Commit to docs/media/, then replace this comment with:
           ![downword converting a ChatGPT answer to a .docx](docs/media/hero.gif)
       - Alt text is required. Describe the actions, not the aesthetics.
     ─────────────────────────────────────────────────────────────────────── -->

> [!WARNING]
> **Pre-release.** The library, the CLI and the test suite are real and green —
> but nothing is published to npm yet, so the version badge above will stay
> empty and `npm install downword` will not work until the first release. The
> hosted web tool ships with it. Until then, build from source; see
> [Development](#development).

---

Every other "markdown to Word" tool either uploads your document to somebody's
server, or hands you an HTML file with a `.docx` extension that Word offers to
repair. downword builds a real OOXML package — `styles.xml`, `numbering.xml`,
`footnotes.xml`, the lot — in the tab you are already in. Nothing leaves it.

## Quick start

### 1. The hosted web tool

Nothing to install. Paste, download, done — the conversion runs in your browser
and the markdown never reaches a server.

**[ksprtech.com/tools/markdown-to-word](https://ksprtech.com/tools/markdown-to-word?utm_source=github&utm_medium=readme)**

### 2. The command line

```sh
npx downword notes.md                        # -> notes.docx
npx downword "docs/**/*.md" --outdir build/  # one .docx per file
cat notes.md | npx downword > notes.docx     # stdin -> stdout
```

```sh
npx downword report.md --theme academic --toc --page-numbers --math omml
```

Every flag, default and exit code: **[packages/cli/README.md](packages/cli/README.md)**.

### 3. The library

```sh
npm install downword
```

```ts
import { convert } from "downword";

const bytes = await convert("# Hello\n\nFrom **downword**.");
// bytes is a Uint8Array holding a real .docx
```

In a browser, straight to a download:

```ts
import { convertToBlob } from "downword";

const markdown =
  "# Quarterly review\n\n| Region | Revenue |\n| ------ | ------: |\n| EMEA   |    1.2M |\n";

const link = document.createElement("a");
link.href = URL.createObjectURL(await convertToBlob(markdown));
link.download = "review.docx";
link.click();
```

With the options that make it look like a document somebody meant to write:

```ts
import { convert, inchesToTwips } from "downword";

const markdown = "# Q3 Report\n\n## Revenue\n\nUp and to the right.\n";

const bytes = await convert(markdown, {
  theme: "academic",
  pageSize: "Letter",
  margins: inchesToTwips(1),
  toc: true,
  pageNumbers: true,
  metadata: { title: "Q3 Report", author: "A. Writer" },
  onWarning: (warning) => console.warn(`${warning.code}: ${warning.message}`),
});
```

Full API, every option, and the warning channel:
**[packages/core/README.md](packages/core/README.md)**.

## Features

| Feature                        | What you get                                                                                                      |
| ------------------------------ | ----------------------------------------------------------------------------------------------------------------- |
| **CommonMark + GFM**           | Headings, emphasis, links, images, tables with alignment, task lists, strikethrough, footnotes                    |
| **Real Word styles**           | `Heading1`–`Heading6`, `Title`, `Quote`, `ListParagraph` — restylable, not stamped-on formatting                  |
| **List numbering that counts** | Each list gets its own numbering instance, so sibling lists restart instead of continuing each other              |
| **Tables**                     | A real `<w:tbl>` with a header row that repeats across page breaks                                                |
| **Native equations**           | `$x^2$` and `$$…$$` become `m:oMath` — editable in Word's equation editor, not a picture                          |
| **Syntax highlighting**        | Fenced code coloured from a WCAG-AA palette; highlight.js loaded lazily, only when a fence needs it               |
| **Images**                     | `data:`, `blob:`, local files (Node) or remote (opt-in), sized to the text column, alt text kept                  |
| **Mermaid diagrams**           | Fences drawn to SVG + raster twin, with the info string as a caption. Browser only                                |
| **Footnotes**                  | Real `word/footnotes.xml` — Word numbers, positions and renumbers them itself                                     |
| **Table of contents**          | A native `TOC` field, so entries stay correct as the document is edited                                           |
| **Page numbers**               | `PAGE` / `NUMPAGES` fields in the footer, computed by the reader during layout                                    |
| **Bookmarks and links**        | Every heading gets a bookmark; `[text](#anchor)` becomes a working intra-document link                            |
| **Five themes**                | `default`, `github`, `academic`, `academic-double`, `print` — or your own token object                            |
| **RTL**                        | `direction: "rtl"` sets the paragraph base direction, so punctuation and list markers land correctly              |
| **No network by default**      | The default image resolver reads `data:` and `blob:` only. Egress is one explicit opt-in, proved by a `fetch` spy |
| **Warnings, not silence**      | Every stage reports what it could not represent, with a severity and the source line or `src`                     |
| **Small**                      | Core entry 74.2 kB min+gzip excluding `docx`, against a 150 kB budget enforced in CI                              |

## How it compares

Be suspicious of a comparison table written by the thing being compared. Here is
an honest one.

|                                | **downword**                                                                 | **pandoc**                                                                | **Cloud converters**                                                 | **Copy-paste into Word**                                       |
| ------------------------------ | ---------------------------------------------------------------------------- | ------------------------------------------------------------------------- | -------------------------------------------------------------------- | -------------------------------------------------------------- |
| **Install needed**             | None for the web tool; `npx` or `npm i` otherwise                            | **Yes** — a binary to install and keep updated                            | None                                                                 | Word itself                                                    |
| **Files uploaded to a server** | **Never** — the browser build cannot reach the network unless you turn it on | **Never** — it runs locally                                               | **Yes.** That is how they work; your document sits on someone's disk | **Never**                                                      |
| **Tables survive**             | Yes — real `<w:tbl>`, repeating header row                                   | Yes                                                                       | Usually                                                              | Only if you paste the _rendered_ answer [¹](#comparison-notes) |
| **Native editable equations**  | Yes — `m:oMath`, opt-in ([one caveat](#limitations))                         | Yes — `m:oMath`, on by default, and more complete                         | Rarely — usually a picture, or literal `$…$`                         | No                                                             |
| **Code highlighting**          | Yes — opt-in, WCAG-AA palette                                                | Yes — on by default, more languages and themes                            | Rarely                                                               | No                                                             |
| **Restylable Word styles**     | Yes — real built-in style IDs                                                | **Yes, and better** — `--reference-doc` restyles from _your_ own template | Rarely — mostly direct formatting                                    | Partly [¹](#comparison-notes)                                  |
| **Batch**                      | Yes — globs and `--outdir`, one warm process                                 | Yes — a shell loop, or many inputs concatenated into one document         | Often, behind an account or a quota                                  | No                                                             |

### Where downword loses

**If you can install pandoc, install pandoc.** It is twenty years old, it is
excellent, and it is more capable than this library in almost every dimension
that is not "runs in a browser tab":

- **Formats.** pandoc converts between dozens of formats in both directions.
  downword does exactly one conversion: markdown → `.docx`. That is the entire
  product.
- **Your own template.** `pandoc --reference-doc=corporate.docx` picks up your
  organisation's real styles, fonts and page furniture. downword has five themes
  and a token object, and cannot read a `.docx` at all.
- **Citations and cross-references.** `--citeproc`, CSL styles, bibliographies,
  numbered figures and tables, cross-references. downword has none of these.
- **Extensibility.** pandoc has Lua filters over a stable document AST. downword
  exposes its parse and render stages plus markdown-it plugins — useful, but
  narrower and less documented.
- **Maturity.** pandoc's `.docx` writer has been shaken out against real Word by
  an enormous number of people over many years. downword is pre-1.0 and, as the
  [fidelity matrix](docs/fidelity-matrix.md) says plainly, **has never been
  opened in Microsoft Word.** Our automated gate is headless LibreOffice, which
  proves the file parses — not that it looks right.
- **Headless diagrams.** pandoc renders mermaid through a filter on any machine
  with Node. downword needs a browser DOM, so in Node the fence stays a code
  block unless you inject a renderer and a rasterizer.

Cloud converters beat downword on nothing except, occasionally, format count —
and they beat it on that by holding your document. Copy-paste beats it on
nothing at all, but it is free and instant, and for a two-paragraph answer it is
genuinely the right tool.

**So who is downword for?** Someone who cannot install a binary (a locked-down
work laptop, a phone, a Chromebook), someone who will not put the document on a
third party's server, or someone embedding markdown → `.docx` inside their own
web application — the case pandoc cannot serve at all.

### Comparison notes

¹ **Copy-paste depends on what you copy.** Pasting the _rendered_ answer out of
a chat UI is an HTML paste: Word maps `<h1>`–`<h6>` onto its heading styles and
keeps tables, which is better than people expect — but it also stamps the web
page's fonts, sizes and colours directly onto the runs, so a Style Set change
afterwards has nothing to grip, and code blocks lose their monospace and
shading. Pasting the markdown _source_ (the copy-code button, or a `.md` file)
gives you literal `## Heading` and literal `| pipes |`.

## Why the output is different

Two things, and they are the reason this library exists.

### 1. Real Word style IDs

A converter can make a paragraph _look_ like a heading by stamping 18 pt bold
blue onto it. Word will render that faithfully and treat it as body text: no
navigation-pane entry, no TOC entry, and Design → Style Set cannot touch it,
because a Style Set rewrites _style definitions_ and there is no style there to
rewrite.

downword references Word's own built-in style IDs, and defines them in a real
`styles.xml`:

```xml
<w:p>
  <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
  <w:r><w:t>Revenue</w:t></w:r>
</w:p>
```

| Word's own IDs, used as Word intends                                                                                                                                 | Ours, named to read sensibly in the Styles pane                                                                                                                       |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Normal`, `Title`, `Heading1`–`Heading6`, `Quote`, `ListParagraph`, `Hyperlink`, `FootnoteReference`, `FootnoteText`, `TOCHeading`, `Footer`, `DefaultParagraphFont` | `CodeChar`, `CodeBlock`, `HtmlBlock`, `HtmlChar`, `HorizontalRule`, `TableText`, `TableHeading`, `TableSpacing`, `Figure`, `ImagePlaceholder`, `ImagePlaceholderChar` |

Everything is based on `Normal`, so restyling reaches body text too. The
practical consequence: edit `Heading 2` once and every H2 follows; apply a
different Style Set and the whole document changes; the navigation pane and the
TOC field both work, because they read the outline the styles declare.

This is also what makes the themes cheap. The three that are purely _looks_
produce a **byte-identical `document.xml`** and differ only in the generated
`styles.xml` — proved by `tests/themes.test.ts`, which renders one fixture under
each and diffs the parts.

### 2. Equations that stay equations

`$…$` and `$$…$$` become native Office Math — `<m:oMath>` in the document body,
with no drawing and no media part:

```xml
<m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>
```

Click it in Word and it opens in the equation editor, ready to edit. Most
converters that handle math at all hand you a PNG, which is a picture of an
equation: it cannot be corrected, searched, restyled or read aloud.

```ts
import { convertWithMath } from "downword/plugins/math";

const bytes = await convertWithMath("Euler: $e^{i\\pi} + 1 = 0$\n", {
  math: { math: "omml", onWarning: (warning) => console.warn(warning.message) },
});
```

It lives behind `downword/plugins/math` and needs two optional peers (`temml`,
`mathml2omml`), so a project with no equations never pays for it.

## Limitations

The honest list. Each of these is measured or reproduced, not guessed.

- **Remote images fail on many hosts, by design of the web.** Egress is off by
  default and `allowRemoteImages: true` turns it on — but in a browser that is a
  cross-origin `fetch`, and a host that does not send
  `Access-Control-Allow-Origin` refuses it. Plenty of image hosts and CDNs do
  not send it. `fetch` rejects with a deliberately vague `TypeError` in that
  case, so downword cannot even tell you it was CORS specifically. The image
  degrades to a visible placeholder paragraph and one `error` warning; the
  document is never lost. Node has no CORS, so the CLI and server-side use are
  unaffected (subject to the SSRF guard on private and loopback addresses).

  ```ts
  import { convert } from "downword";

  const bytes = await convert("![logo](https://example.com/logo.png)", {
    allowRemoteImages: true,
    onWarning: (warning) => {
      if (warning.stage === "image") console.warn(`${warning.src}: ${warning.message}`);
    },
  });
  ```

- **`convert()` is not byte-reproducible.** The same markdown twice gives you two
  different files. `docx` mints a nanoid relationship id per external hyperlink
  and stamps `docProps/core.xml` with the wall clock, and hands every part to
  JSZip with no date, so the archive records the clock in each local file header
  too — two conversions more than a second apart differ across most of their
  bytes. The reproducible unit is the **part**, not the package, which is what
  the golden tests compare after normalising those ambient values. A hash-stable
  artefact means rewriting the archive yourself.

- **Packing is super-linear; keep documents to a couple of MB.** Measured on Node
  22.19 (Apple silicon) over a mixed prose/list/table corpus, through
  `convert()`:

  | Markdown | parse  | render | **pack** | total   |
  | -------- | ------ | ------ | -------- | ------- |
  | 1 MB     | 0.28 s | 0.41 s | 2.1 s    | 2.6 s   |
  | 2 MB     | 0.50 s | 0.62 s | 6.5 s    | 7.3 s   |
  | 4 MB     | 0.96 s | 1.13 s | 24.9 s   | 26.1 s  |
  | 8 MB     | —      | —      | —        | **OOM** |

  Parsing and rendering are linear and cheap. `Packer.toArrayBuffer` — `docx`'s
  own XML serialisation and zipping — dominates and grows roughly as n^1.8, and
  at 8 MB it exhausts Node's default heap (~4.3 GB on that machine) and crashes.
  So: ~2 MB is comfortable, 4 MB is a coffee break, 8 MB needs
  `--max-old-space-size` and patience. Absolute times are machine-dependent; the
  shape of the curve is not. In a browser tab, do this in a worker.

- **`math: "image"` needs a rasterizer you supply.** Turning MathML into pixels
  needs a layout engine and this package does not contain one. With no
  `rasterizer` passed, `"image"` reports `image-no-rasterizer` and falls back to
  OMML rather than dropping the equation — it does not pretend to have drawn a
  picture.

- **N-ary operators (∫, ∑) draw an empty placeholder box.** `$\int_0^1 x^2 dx$`
  produces an `<m:nary>` whose body `<m:e/>` is empty, with the integrand
  following as a sibling, so Word and LibreOffice show the dotted "empty slot"
  box after the integral sign. The equation is complete and fully editable; the
  box is cosmetic. The cause is upstream — `temml` emits `<msubsup>` plus
  siblings and `mathml2omml` has nothing to put in `<m:e>` — and fixing it needs
  a semantic OMML rewrite in downword.

- **Mermaid needs a browser.** mermaid measures text by laying it out in a DOM,
  so there is no headless path that is not a headless browser. In Node the pass
  leaves each fence exactly as it found it — the diagram source is still in the
  document, as a code block — and raises one `no-dom` notice. Nothing throws and
  nothing disappears. To render server-side, inject both seams: a `renderer`
  (mermaid through jsdom, or a real headless browser) and a `rasterizer`.

- **The TOC field arrives empty.** OOXML stores the _instruction_, not the
  entries; the reader computes them. downword marks the field dirty and sets
  `<w:updateFields/>`, which is everything the format allows a generator to do.
  In Word, answer yes to the "update fields?" prompt on open, or select the field
  and press **F9**. In LibreOffice, Tools → Update → Indexes and Tables. Google
  Docs, Pages and Quick Look do not run fields at all and will show nothing
  there. Every document with a TOC raises a `toc-needs-update` notice so your UI
  can say so in its own words.

- **Raw HTML has no lossless answer.** LLM output is full of `<br>`, `<sub>` and
  `<div align="center">`, and OOXML has nowhere to put them. You get three honest
  options — `"escape"` (the default: literal text), `"keep"` (verbatim in a
  monospace run, visibly unconverted) and `"drop"` — and there is no fourth one
  that works.

- **Never verified in Microsoft Word.** See below.

## Fidelity

CI opens every generated `.docx` with headless LibreOffice and converts it to
PDF, with a positive control on the import filter and a negative control on a
deliberately corrupt file. That proves the container is well-formed and that one
real implementation parses it end to end. It does **not** prove that a heading
looks like a heading, that a header row repeats, or that an equation typesets.

**[docs/fidelity-matrix.md](docs/fidelity-matrix.md)** tracks how the output
actually renders across Word (Windows, Mac, Online), Google Docs, LibreOffice
Writer and Apple Pages. Every cell is currently marked "awaiting manual
verification", because it is — that file is a protocol and a scoreboard, not a
set of claims. Filling it in is the single most valuable contribution anyone with
a copy of Word can make.

## Packages

| Package                          | npm            | What it is                          |
| -------------------------------- | -------------- | ----------------------------------- |
| [`packages/core`](packages/core) | `downword`     | The library. Browser and Node.      |
| [`packages/cli`](packages/cli)   | `downword-cli` | The `downword` command line binary. |

Subpath entry points — everything heavy lives behind one, and
`tests/bundle.test.ts` fails if any of it reappears in the main entry's import
graph:

| Subpath                    | Contents                                       |
| -------------------------- | ---------------------------------------------- |
| `downword`                 | `convert`, the model, the themes, the warnings |
| `downword/highlight`       | the highlight.js adapter (optional peer)       |
| `downword/images/node`     | the filesystem + SSRF-guarded image resolver   |
| `downword/plugins/math`    | `$…$` → native Word equations (optional peers) |
| `downword/plugins/mermaid` | mermaid diagrams (optional peer)               |

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md). Short version:

```sh
pnpm install
pnpm format && pnpm lint && pnpm typecheck && pnpm test && pnpm build && pnpm size
```

Every code sample in this README and in
[packages/core/README.md](packages/core/README.md) is extracted and type-checked
against the **built `.d.ts` files** by `packages/core/tests/readme.test.ts`, so a
sample that drifts from the API stops compiling and CI goes red.

## License

[MIT](LICENSE) © 2026 KSPR Technologies
