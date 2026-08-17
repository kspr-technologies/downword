# downword

**Your LLM answers in markdown. The person who asked for it wants a Word document.**

> markdown → Word (`.docx`), entirely in your browser.
> Paste from ChatGPT/Claude, get a real Word document.

[![CI](https://github.com/kspr-technologies/downword/actions/workflows/ci.yml/badge.svg)](https://github.com/kspr-technologies/downword/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/downword?color=cb3837&logo=npm&label=npm)](https://www.npmjs.com/package/downword)
[![core bundle](https://img.shields.io/badge/core-74.2%20kB%20min%2Bgzip-1f6feb)](.size-limit.json)
[![license](https://img.shields.io/badge/license-MIT-1f6feb)](../../LICENSE)

Nothing to install, if you only want the output:
**[the hosted web tool](https://ksprtech.com/tools/markdown-to-word?utm_source=github&utm_medium=readme)**
is this library running in your browser tab.

Every "markdown to Word" tool either uploads your document to somebody's server
or hands you an HTML file with a `.doc` extension that Word complains about.
downword builds a real OOXML package in memory — headings, lists that restart
their numbering, tables with repeating header rows, footnotes, embedded images,
styled code — and nothing leaves the process unless you ask it to.

```sh
npm install downword
```

## Quick start

```ts
import { convert } from "downword";

const bytes = await convert("# Hello\n\nFrom **downword**.");
```

In a browser, `convertToBlob` gives you something you can hand straight to a
download link:

```ts
import { convertToBlob } from "downword";

const blob = await convertToBlob("# Hello");
const link = document.createElement("a");
link.href = URL.createObjectURL(blob);
link.download = "hello.docx";
link.click();
```

In Node:

```ts
import { writeFile } from "node:fs/promises";
import { convert } from "downword";

await writeFile("hello.docx", await convert("# Hello"));
```

## What it converts

CommonMark, plus the GitHub extensions people actually type.

| Markdown                             | Word                                                        |
| ------------------------------------ | ----------------------------------------------------------- |
| `#` … `######`                       | `Heading1`–`Heading6`, with outline levels and bookmarks    |
| `**bold**`, `*italic*`, `~~strike~~` | real character formatting, not a font hack                  |
| `` `code` ``, ` ```lang `            | `CodeChar` / `CodeBlock` styles, optional syntax colouring  |
| `-` / `1.` / `- [x]`                 | real list numbering that restarts per list and nests 9 deep |
| `> quote`                            | `Quote` style with a left rule, nested quotes indent        |
| tables, with alignment               | a real `<w:tbl>` with a repeating header row                |
| `[text](url)`, `[text](#anchor)`     | external hyperlinks and internal bookmark links             |
| `![alt](src)`                        | an embedded image, sized to the text column                 |
| `[^1]` footnotes                     | real Word footnotes in `word/footnotes.xml`                 |
| `---`                                | a horizontal rule                                           |

Everything is **style-driven**: the document carries a real `styles.xml`, so a
reader can restyle the whole thing by editing `Heading 2` once, exactly as they
would with a document Word produced.

## API

### `convert(markdown, options?): Promise<Uint8Array>`

The whole library in one call. Never performs I/O unless you opt in, and never
throws for bad markdown — markdown has no syntax errors, only constructs OOXML
cannot hold, and those are reported through `onWarning` while the conversion
carries on. Converting `""` gives you a valid, empty document.

### `convertToBlob(markdown, options?): Promise<Blob>`

The same, wrapped in a `Blob` tagged with `DOCX_MIME_TYPE`. Works in Node ≥18
too, where `Blob` is a global.

### `convertToDocument(markdown, options?): Promise<Document>`

`convert()` minus the final zip, for when you want `docx`'s own `Packer`:

```ts
import { Packer } from "docx";
import { convertToDocument } from "downword";

const file = await convertToDocument("# Hello");
const base64 = await Packer.toBase64String(file);
```

## Options

Every field is optional; the default is in the last column.

| Option              | Type                                     | Default      |
| ------------------- | ---------------------------------------- | ------------ |
| `theme`             | `ThemeName \| ThemeInit`                 | `"default"`  |
| `pageSize`          | `"A4" \| "Letter" \| { width, height }`  | `"A4"`       |
| `orientation`       | `"portrait" \| "landscape"`              | `"portrait"` |
| `margins`           | `number \| { top, right, bottom, left }` | `1440` (1in) |
| `direction`         | `"ltr" \| "rtl"`                         | `"ltr"`      |
| `tabSize`           | `number` (integer, 0–64)                 | `4`          |
| `titleBlock`        | `boolean`                                | `false`      |
| `toc`               | `boolean \| TocInit`                     | `false`      |
| `pageNumbers`       | `boolean \| PageNumbersInit`             | `false`      |
| `metadata`          | `ConvertMetadata`                        | none         |
| `footnotes`         | `boolean \| "inline"`                    | `true`       |
| `html`              | `"escape" \| "keep" \| "drop"`           | `"escape"`   |
| `linkify`           | `boolean`                                | `true`       |
| `typographer`       | `boolean`                                | `false`      |
| `lineBreaks`        | `"collapse" \| "preserve"`               | `"collapse"` |
| `plugins`           | `MarkdownItPlugin[]`                     | none         |
| `highlighter`       | `Highlighter`                            | none         |
| `imageResolver`     | `ImageResolver`                          | data:/blob:  |
| `allowRemoteImages` | `boolean`                                | `false`      |
| `onWarning`         | `(warning: ConvertWarning) => void`      | none         |

Lengths are in **twips** — twentieths of a point, 1440 to the inch, which is
what OOXML itself uses. `inchesToTwips` and `pointsToTwips` are exported:

```ts
import { convert, inchesToTwips } from "downword";

await convert("# Hello", {
  pageSize: "Letter",
  orientation: "landscape",
  margins: inchesToTwips(0.75),
  metadata: {
    title: "Q3 Report",
    author: "A. Writer",
    subject: "Revenue",
    keywords: ["finance", "q3"],
  },
});
```

### Table of contents

`toc` emits a **native Word `TOC` field** — not a list of headings frozen at
conversion time, but the instruction Word itself writes, so the entries stay
correct as the document is edited and each one links to its heading.

```ts
import { convert } from "downword";

await convert("# One\n\n## Two\n", {
  toc: { minLevel: 1, maxLevel: 3, title: "Contents" },
  pageNumbers: { format: "page-x-of-y" },
});
```

**The field is empty until the reader updates it.** OOXML stores the
instruction, not the entries, and the entries are computed by the word processor
from the document's heading outline. downword marks the field dirty and sets
`<w:updateFields/>`, which is everything the format allows a generator to do —
after that it is up to the reader:

- **Word** — answer yes to _"This document contains fields that may refer to
  other files. Update?"_ on open, or right-click the field → **Update Field**
  (or select it and press **F9**).
- **LibreOffice Writer** — Tools → Update → Indexes and Tables.
- **Google Docs, Pages, Quick Look, most converters** — they do not run fields
  at all, and will show nothing there.

Every document that gets a TOC raises one `toc-needs-update` notice so a host
can say this in its own words. Page numbers need no such step: `PAGE` and
`NUMPAGES` are computed during layout by every reader that paginates.

### Footnotes

`[^1]` becomes a real Word footnote — a superscript, clickable reference in the
body and the note in `word/footnotes.xml`, which Word numbers, positions and
renumbers itself. The body and the footnotes part are written as a matched pair:
no reference points at a note that is not there, and no note sits there with no
reference pointing at it.

`footnotes: "inline"` splices each note into the sentence that cited it, in
parentheses, for a reader with no footnote pane. `footnotes: false` is a
different answer again — the syntax is not parsed at all and `[^1]` stays
literal text.

Word cannot draw a footnote inside a footnote, so a `[^y]` written inside note
`x` is spliced into `x`'s text the same way rather than becoming a reference
into nothing. Two notes that cite each other stop after one expansion and say
so.

### Right-to-left documents

`direction: "rtl"` sets the document's **base** writing direction: prose
paragraphs get `<w:bidi/>`, their runs get `<w:rtl/>` — which is also what makes
Word use the complex-script face and size rather than the Latin ones — and
tables get `<w:bidiVisual/>` so the column order mirrors too. Code blocks, raw
HTML and display math stay left-to-right, because their content is not prose.

```ts
import { convert } from "downword";

await convert("# مرحبا بالعالم\n\n- عنصر\n- عنصر آخر\n", { direction: "rtl" });
```

It is not cosmetic. The Unicode Bidirectional Algorithm resolves trailing
punctuation and mixed Latin/Arabic segments against the _paragraph's_ base
direction, so leaving an Arabic or Hebrew document at `"ltr"` puts the full stop
on the wrong side of the line and the list marker on the wrong side of the item.
Doing that raises one `rtl-not-enabled` notice rather than passing in silence.

### Raw HTML

LLM output is full of stray `<br>`, `<sub>` and `<div align="center">`, and
OOXML has nowhere to put them, so there is no lossless answer — only three
honest ones. `"escape"` (the default) leaves the markup as literal text,
`"keep"` re-emits it verbatim in a monospace run so it is visibly unconverted,
and `"drop"` removes it. `"keep"` and `"drop"` warn once per construct;
`"escape"` does not parse the markup at all, so there is nothing to report and
nothing lost.

## Warnings

A warning never stops a conversion; it exists so you can tell a user _"three
images could not be loaded"_ instead of handing them a file with holes in it.

Every stage — the parser, the image pass, the highlighting pass and the
renderer — reports to `onWarning` in the same shape:

| field      | meaning                                                                                                 |
| ---------- | ------------------------------------------------------------------------------------------------------- |
| `stage`    | `"parse"`, `"image"` or `"render"`. Also the union's discriminant.                                      |
| `code`     | Machine-readable kind, from a small closed set per stage.                                               |
| `severity` | `"error"` if content is missing from the file, `"notice"` if it is all there but expressed differently. |
| `message`  | Human-readable, already naming the construct. Safe to show a user.                                      |

On top of that a stage adds what only it knows: a parse warning carries the
source `line`, an image diagnostic carries the `src` that failed. So the common
case — count them, group them, show the errors — needs no `switch` at all.

```ts
import { convert, type ConvertWarning } from "downword";

const problems: ConvertWarning[] = [];

await convert("![](https://example.com/logo.png)", {
  onWarning: (warning: ConvertWarning) => {
    if (warning.severity === "error") problems.push(warning);
  },
});

const summary = `${problems.length} problem(s): ${problems.map((p) => p.code).join(", ")}`;
```

`severity` is a property of the `code`, not of the moment — `renderWarningSeverity`
and `parseWarningSeverity` are the same lookups downword uses, and
`RENDER_WARNING_CODES` / `PARSE_WARNING_CODES` enumerate the sets, which is what
a UI needs to build a per-code label table.

The two async prepare passes report here too: `prepareHighlights` and
`prepareImages` both take `{ onWarning }` and raise `highlighter-failed` /
`image-resolver-failed` for each node whose enhancement an adapter lost.

One thing does **not** arrive on this channel: a highlighter's own diagnostics.
`convert()` only sees an adapter that _threw_, so a fence that came out grey
because highlight.js is not installed, or because the language is unknown, is
invisible from here. Those are things only the adapter knows, and
`createHighlighter` reports them to its own `onWarning` — wire that up as well
if you care why a block is uncoloured:

```ts
import { createHighlighter } from "downword/highlight";

const highlighter = createHighlighter({
  // "engine-unavailable" | "language-unknown" | "language-load-failed" | ...
  onWarning: (w) => console.warn(`${w.code} (${w.lang ?? "no language"}): ${w.message}`),
});
```

## Errors

Everything downword throws deliberately is a `DownwordError` carrying a
machine-readable `code` (`"invalid-input"`, `"invalid-options"`,
`"parse-failed"`, `"render-failed"`, `"pack-failed"`, `"not-implemented"`) and,
where there was one, the original error as `cause`. Use `isDownwordError`
rather than `instanceof`: it is a structural check, so it still works in a
project that ends up with both the ESM and the CJS build in its graph.

```ts
import { convert, isDownwordError } from "downword";

try {
  await convert("# Hello");
} catch (error) {
  if (isDownwordError(error) && error.code === "pack-failed") {
    console.error("could not build the zip", error.cause);
  } else {
    throw error;
  }
}
```

## Syntax highlighting

highlight.js is an **optional peer dependency** and lives behind its own
subpath, so a project that does not want it never pays for it. Install it, then:

````ts
import { convert } from "downword";
import { createHighlighter } from "downword/highlight";

const highlighter = createHighlighter();

const bytes = await convert("```ts\nconst x: number = 1;\n```", { highlighter });
````

Nothing is loaded eagerly: the engine arrives through `await
import("highlight.js/lib/core")` on the first fence that needs it, and each
grammar through its own dynamic import on the first fence written in that
language. A document with no code blocks — or only unlabelled ones — fetches
nothing at all. Colour comes from the **theme**, not from here — `"default"`
clears WCAG AA against the code block's own shading (floor 4.74:1) and
`"print"` clears AAA (floor 7.13:1). What the highlighter keeps for itself is
the part a palette cannot express: keywords and comments carry a weight/slant
signal too, so a greyscale printout still separates them. Pass `scopeStyles` to
take the colours back — `createHighlighter({ scopeStyles: PRINT_SCOPE_STYLES })`
pins the print inks whatever the theme is.

Create the highlighter once and reuse it; it caches the engine and every
grammar it has loaded.

## Images

The default resolver reads `data:` and `blob:` URLs and nothing else, so a
conversion **cannot make a network request** out of the box. That is a property
you can check, not a promise in a comment: the egress switch is tested with a
`fetch` spy that must never be called.

```ts
import { convert } from "downword";

// Fetches over the network, capped at 10 MiB and 10 s per image.
await convert("![logo](https://example.com/logo.png)", { allowRemoteImages: true });
```

In Node, `downword/images/node` adds local files, a base directory that paths
and symlinks cannot escape, and an SSRF guard on remote hosts:

```ts
import { convert } from "downword";
import { createNodeImageResolver } from "downword/images/node";

const bytes = await convert("![diagram](./diagram.png)", {
  imageResolver: createNodeImageResolver({ baseDir: "./docs" }),
});
```

An image that cannot be fetched, is too big, times out, or is in a format OOXML
cannot hold costs exactly one placeholder and one warning — never the document.
SVG needs a rasteriser, because OOXML stores an SVG _plus_ a raster twin; pass
one as `rasterizer`, or the image degrades to the placeholder.

## Diagrams (mermaid)

````ts
import { parseMarkdown, renderDocument } from "downword";
import { renderMermaid } from "downword/plugins/mermaid";

const parsed = parseMarkdown("# Design\n\n```mermaid\nflowchart LR\n  A --> B\n```\n");
const { document, warnings } = await renderMermaid(parsed);

const file = renderDocument(document);
````

Each ` ```mermaid ` fence becomes a centred figure: mermaid draws an SVG,
the SVG is rastered at **twice** its display size, and the PNG is embedded with
alt text. The info string is the caption — ` ```mermaid The request pipeline `
— and so is mermaid's own `title:` frontmatter; with one, an italic caption
paragraph follows the picture.

Three things are worth knowing before you reach for it.

**It needs a browser.** mermaid measures text by laying it out, so there is no
headless path that is not a headless browser. In Node the pass leaves every
fence exactly as it found it — the diagram's source is still in the document, as
a code block — and reports one `no-dom` warning saying so. Nothing throws and
nothing disappears. To render in Node, supply both seams: a `renderer` (mermaid
driven through jsdom, or a headless browser) and a `rasterizer` (`resvg`,
`sharp`, `@napi-rs/canvas`).

**mermaid is an optional peer dependency**, ~500 kB, loaded through
`await import("mermaid")` on the first diagram. A document with none fetches
nothing, and a project that never imports this subpath does not pay for it —
which `tests/bundle.test.ts` proves by bundling the main entry and failing if
mermaid appears in its import graph.

**What lands in the file is a picture.** OOXML cannot hold a bare SVG: it stores
a raster _plus_ an `asvg:svgBlip` extension, and `docx`'s `ImageRun` requires the
raster twin. `embed: "svg"` ships both halves, so Word 2016+ draws the vector and
everything else draws the same PNG.

```ts
import { parseMarkdown } from "downword";
import { renderMermaid } from "downword/plugins/mermaid";

const { document, diagrams, rendered, warnings } = await renderMermaid(parseMarkdown("# hi"), {
  embed: "svg",
  scale: 3,
  caption: false,
  config: { theme: "neutral" },
  onWarning: (warning) => console.warn(`${warning.code}: ${warning.message}`),
});
```

A warning is a `notice` when the environment could never have rendered the
diagram (`no-dom`, `engine-unavailable`) and an `error` when downword tried and
failed (`render-failed`, `rasterize-failed`, …). Either way the fence survives,
so a diagram is never silently missing.

## Themes

Every number, colour, font and glyph the renderer can emit is a theme token.
Pass a built-in by name, or an object that is merged one level deep over the
default:

```ts
import { convert, THEMES } from "downword";

await convert("# Hello", { theme: "print" });

await convert("# Hello", {
  theme: {
    fonts: { body: "Georgia", mono: "IBM Plex Mono" },
    colors: { heading: "1F3864" },
    sizes: { body: 24 },
  },
});

await convert("# Hello", {
  theme: { ...THEMES.print, spacing: { ...THEMES.print.spacing, paragraphAfter: 240 } },
});
```

Five built-ins, and the three that are _looks_ differ **only in the generated
`styles.xml`** — the same markdown produces the same `document.xml`, dressed
differently. That is what makes the output restylable in Word (Design → Style
Set rewrites style definitions, so anything stamped onto an element instead is
beyond its reach), and `tests/themes.test.ts` proves it by rendering one fixture
under each and diffing the parts.

| Theme               | What it is                                                                                      |
| ------------------- | ----------------------------------------------------------------------------------------------- |
| `"default"`         | Word's own look: Aptos with a Calibri fallback, Word's heading blues, 1.15 leading              |
| `"github"`          | a README as Word can render one: Segoe UI (fallback Arial), near-black headings, single leading |
| `"academic"`        | Times New Roman 12 pt (fallback Cambria), black headings, 1.5 leading, Courier New for code     |
| `"academic-double"` | the same, double-spaced                                                                         |
| `"print"`           | `"default"` with the AAA-contrast code palette                                                  |

OOXML has no font stack, so a theme names a preferred face and a fallback in the
`w:cs` slot rather than a list: `fonts: { body: { ascii: "Aptos", hAnsi: "Aptos", cs: "Calibri" } }`.
A plain string still works and is expanded into every script slot.

Colours are `RRGGBB` without a leading `#`; sizes are half-points (`24` = 12pt);
spacing is twips. Every group is merged one level deep except the three that are
designed sets rather than bags of tokens — `bulletLevels`, `orderedFormats` and
`codePalette` — which **replace** the default outright.

`codePalette` is what colours code, on both of the renderer's highlighting
paths: scope spans carried on the model, and a `highlighter` that reports its
scopes (which `createHighlighter()` does). So the theme governs a real fenced
block too — `"default"` clears WCAG AA against the code shading, `"print"`
clears AAA — while the highlighter keeps the part a palette cannot express,
bold keywords and italic comments. Pass `scopeStyles` to `createHighlighter` to
take the colours back.

## The two-stage pipeline

`convert()` is `parseMarkdown` → resolve images → highlight → `renderDocument` →
pack. Both synchronous stages are exported, so you can look at — or rewrite —
the document in between. `parseMarkdown` never throws, and `renderDocument` does
no I/O: everything downword writes into the package is a function of the model
alone. The reproducible unit is the part, not the package — `docx` stamps
`docProps/core.xml` with the wall clock, mints a random relationship id per
external hyperlink, and hands every part to JSZip with no date, so the archive
records the clock in each local file header as well. Two `.docx` files built
from one model more than two seconds apart therefore differ in most of their
bytes. A hash-stable artefact means rewriting the archive, not just those two
values.

```ts
import { Packer } from "docx";
import { nodesOfType, parseMarkdown, renderDocument } from "downword";

const document = parseMarkdown("# Title\n\n## Section\n\nBody.");

const outline = nodesOfType(document, "heading").map((heading) => ({
  level: heading.level,
  id: heading.id,
}));
console.log(outline); // [{ level: 1, id: "title" }, { level: 2, id: "section" }]

const bytes = new Uint8Array(await Packer.toArrayBuffer(renderDocument(document)));
```

## Limitations

Every item here is measured or reproduced, not guessed. Several are properties
of OOXML or of the web rather than bugs, but they will all surprise somebody, so
they are written down.

**Remote images fail on many hosts, by design of the web.** `allowRemoteImages:
true` turns egress on, but in a browser that is a cross-origin `fetch`, and a
host that does not send `Access-Control-Allow-Origin` refuses it — which plenty
of image hosts and CDNs do not send. `fetch` rejects with a deliberately vague
`TypeError` in that case, so downword cannot even report that it was CORS
specifically. The image degrades to a visible placeholder and one `error`
warning; the document is never lost. Node has no CORS, so `downword/images/node`
and the CLI are unaffected, subject to their SSRF guard.

**`convert()` is not byte-reproducible.** See
[The two-stage pipeline](#the-two-stage-pipeline) above: the reproducible unit is
the part, not the package. Same markdown, two runs, two different files.

**Packing is super-linear.** Measured on Node 22.19 (Apple silicon) over a mixed
prose/list/table corpus:

| Markdown | `parseMarkdown` | `renderDocument` | **pack** | total   |
| -------- | --------------- | ---------------- | -------- | ------- |
| 1 MB     | 0.28 s          | 0.41 s           | 2.1 s    | 2.6 s   |
| 2 MB     | 0.50 s          | 0.62 s           | 6.5 s    | 7.3 s   |
| 4 MB     | 0.96 s          | 1.13 s           | 24.9 s   | 26.1 s  |
| 8 MB     | —               | —                | —        | **OOM** |

Both exported stages are linear and cheap. `Packer.toArrayBuffer` — `docx`'s own
XML serialisation and zipping — dominates and grows roughly as n^1.8, and at
8 MB it exhausts Node's default heap (~4.3 GB on that machine). ~2 MB is
comfortable; past 4 MB you want `--max-old-space-size`. Absolute times are
machine-dependent, the shape of the curve is not. In a browser, convert in a
worker so the tab stays responsive.

**`math: "image"` needs a rasterizer you supply**, because turning MathML into
pixels needs a layout engine and this package does not contain one. Without one
it reports `image-no-rasterizer` and falls back to OMML rather than pretending.

**N-ary operators (∫, ∑) draw an empty placeholder box.** `$\int_0^1 x^2 dx$`
yields an `<m:nary>` whose body `<m:e/>` is empty, with the integrand following
as a sibling, so readers show the dotted "empty slot" box after the operator.
The equation is complete and fully editable; the box is cosmetic. The cause is
upstream — `temml` emits `<msubsup>` plus siblings, and `mathml2omml` has nothing
to hoist into `<m:e>`.

**Mermaid needs a browser.** In Node the fence survives as a code block and the
pass raises one `no-dom` notice. See [Diagrams (mermaid)](#diagrams-mermaid).

**The TOC field arrives empty** until the reader updates it. See
[Table of contents](#table-of-contents).

**Raw HTML has no lossless answer** — only the three honest ones under
[Raw HTML](#raw-html).

## Fidelity

CI opens every generated `.docx` with headless LibreOffice and converts it to
PDF, with a positive control on the import filter and a negative control on a
deliberately corrupt file. That proves the container is well-formed and that one
real implementation parses it end to end — **not** that a heading looks like a
heading or that an equation typesets.

How the output actually renders in Word (Windows, Mac, Online), Google Docs,
LibreOffice Writer and Apple Pages is tracked in
[docs/fidelity-matrix.md](https://github.com/kspr-technologies/downword/blob/main/docs/fidelity-matrix.md).
Every cell there is currently marked "awaiting manual verification", because it
is: the library has never been opened in Microsoft Word.

## Size and support

The whole public surface of the `.` entry is **74.2 kB min+gzip** excluding
`docx`, against a **150 kB** budget enforced in CI by `size-limit`. `docx`
itself adds ~110 kB gzip — the same entry measured _including_ it comes to
184.7 kB — and is budgeted separately, because it is a pre-bundled,
non-tree-shakeable blob that no amount of care here shrinks.

Nothing heavy is reachable from the main entry — highlight.js, and later
`temml` and `mermaid`, all live behind subpath exports. That is not a promise:
`tests/bundle.test.ts` bundles `src/index.ts` for a browser, reads esbuild's
metafile, and fails if any of them reappears in the import graph or if a Node
builtin becomes reachable.

- **Browsers**: any modern one. No Node builtins are reachable from the main
  entry, which is enforced by a test rather than asserted.
- **Node**: ≥20. ESM and CommonJS, both real builds.
- **TypeScript**: ≥5, `strict` and `exactOptionalPropertyTypes` clean.

Subpath entry points:

| Subpath                    | Contents                                            |
| -------------------------- | --------------------------------------------------- |
| `downword`                 | `convert` and everything above                      |
| `downword/highlight`       | the highlight.js adapter (optional peer dependency) |
| `downword/images/node`     | the filesystem + SSRF-guarded image resolver        |
| `downword/plugins/math`    | `$…$` → native Word equations (optional peers)      |
| `downword/plugins/mermaid` | mermaid diagrams (optional peer dependency)         |

---

- Source, issues and docs: <https://github.com/kspr-technologies/downword>
- CLI: [`downword-cli`](https://www.npmjs.com/package/downword-cli)

MIT © 2026 KSPR Technologies
