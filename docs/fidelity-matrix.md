# Fidelity matrix

How a downword `.docx` actually renders in the word processors people open it
with.

> [!IMPORTANT]
> **Microsoft Word has still never opened one of these files.** On 2026-08-18 a
> render pass filled in the two columns this project can actually drive on a
> Mac — **LibreOffice Writer** and **Apple Pages**. Everything else in the
> matrix is still 🧑.
>
> Be precise about what that pass was. It generated documents with the real
> CLI, had each application lay them out and export a PDF, rasterised the pages
> and **looked at the images**. That establishes what the page _looks like_. It
> is not a human sitting in the application, so every row that needs an
> interactive action — open the Styles gallery, edit `Heading 2` and watch the
> other H2s follow, press F9, delete a footnote — is still 🧑 even in the two
> filled columns. The [Verification log](#verification-log) records exactly
> what was run.
>
> Do not guess a cell. A matrix with invented ticks in it is worse than no
> matrix, because it converts an unknown into a promise.
>
> What CI proves independently is recorded in
> [What CI already proves](#what-ci-already-proves), and it remains a much
> smaller claim than this matrix makes: CI proves the files **open**, not that
> they **look right**.

## Legend

| Mark | Meaning                                                                               |
| ---- | ------------------------------------------------------------------------------------- |
| 🧑   | Awaiting verification. **The state of every Word, Word Online and Google Docs cell.** |
| ✅   | Verified: renders as intended.                                                        |
| ⚠️   | Verified: renders, with a caveat. **Must** carry a numbered note below.               |
| ❌   | Verified: broken, missing or wrong. **Must** carry a numbered note below.             |
| ➖   | Not applicable — the application has no such feature to get right.                    |

✅ / ⚠️ / ❌ mean "somebody looked". **Who looked, at what, and how** is recorded
in [Verification log](#verification-log) — the LibreOffice and Apple Pages
columns were established by rendered-output inspection, not by a human driving
the GUI, and that distinction decides which rows could be answered at all.

When you fill a cell in, record the exact application version and platform in
the log. "Word" is not a version.

## The matrix

Columns are the six readers a document produced by this library is most likely
to be opened in.

| Feature                                            | Word (Win) | Word (Mac) | Word Online | Google Docs | LibreOffice Writer | Apple Pages      |
| -------------------------------------------------- | ---------- | ---------- | ----------- | ----------- | ------------------ | ---------------- |
| **Structure**                                      |            |            |             |             |                    |                  |
| `Heading1`–`Heading6` appear in the Styles gallery | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n2](#notes)    | 🧑 [n2](#notes)  |
| Headings populate the navigation / outline pane    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n2](#notes)    | 🧑 [n2](#notes)  |
| Editing `Heading 2` once restyles every H2         | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n2](#notes)    | 🧑 [n2](#notes)  |
| Design → Style Set restyles the whole document     | 🧑         | 🧑         | 🧑          | ➖          | 🧑 [n2](#notes)    | ➖               |
| Heading levels are visually distinguishable        | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n3](#notes)    | ⚠️ [n3](#notes)  |
| `Title` style on the frontmatter title block       | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n4](#notes)    | 🧑 [n4](#notes)  |
| **Lists**                                          |            |            |             |             |                    |                  |
| Bullet lists, correct glyph per level              | 🧑         | 🧑         | 🧑          | 🧑          | ✅ [n5](#notes)    | ✅ [n5](#notes)  |
| Sibling ordered lists each restart at 1            | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Nested lists indent correctly to 9 levels          | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n6](#notes)    | ⚠️ [n6](#notes)  |
| Ordered list numbering format per level (1/a/i)    | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Task lists `- [x]` show checked/unchecked boxes    | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ⚠️ [n7](#notes)  |
| `ListParagraph` style is applied and editable      | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n2](#notes)    | 🧑 [n2](#notes)  |
| **Tables**                                         |            |            |             |             |                    |                  |
| Table renders with borders and correct cell text   | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Column alignment (`:--`, `:-:`, `--:`) honoured    | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Header row repeats across a page break             | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n8](#notes)    | 🧑 [n8](#notes)  |
| Two adjacent tables stay two tables                | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Column widths fit the text column, no overflow     | 🧑         | 🧑         | 🧑          | 🧑          | ✅ [n9](#notes)    | ✅ [n9](#notes)  |
| **Text and code**                                  |            |            |             |             |                    |                  |
| Bold / italic / strikethrough                      | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| `Quote` style, left rule, nested quotes indent     | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Inline `CodeChar` — monospace, shaded              | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Fenced `CodeBlock` — shading spans the block       | 🧑         | 🧑         | 🧑          | 🧑          | ✅ [n10](#notes)   | ✅ [n10](#notes) |
| Syntax highlighting colours survive                | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Horizontal rule (`---`) draws a rule               | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Tab size / leading match the theme                 | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ❌ [n11](#notes) |
| **Links and navigation**                           |            |            |             |             |                    |                  |
| External hyperlinks are clickable                  | 🧑         | 🧑         | 🧑          | 🧑          | ✅ [n12](#notes)   | 🧑 [n12](#notes) |
| `[text](#anchor)` jumps to the heading bookmark    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n12](#notes)   | 🧑 [n12](#notes) |
| TOC field populates after Update Field / F9        | 🧑         | 🧑         | 🧑          | ➖          | 🧑 [n13](#notes)   | ⚠️ [n14](#notes) |
| TOC entries link to their headings                 | 🧑         | 🧑         | 🧑          | ➖          | 🧑 [n13](#notes)   | 🧑 [n14](#notes) |
| Page numbers (`PAGE` / `NUMPAGES`) compute         | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| **Footnotes**                                      |            |            |             |             |                    |                  |
| Superscript marker in the body, note at page foot  | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Word renumbers notes after one is deleted          | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n2](#notes)    | 🧑 [n2](#notes)  |
| `CodeChar` and math inside a note body survive     | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n15](#notes)   | ⚠️ [n15](#notes) |
| **Images and diagrams**                            |            |            |             |             |                    |                  |
| Raster image, sized to the text column             | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n16](#notes)   | ⚠️ [n16](#notes) |
| Aspect ratio exact, no stretch                     | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| SVG + raster twin: vector drawn where supported    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑               |
| Alt text reaches the accessibility pane            | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n2](#notes)    | 🧑 [n2](#notes)  |
| Mermaid diagram picture + italic caption           | 🧑         | 🧑         | 🧑          | 🧑          | 🧑 [n17](#notes)   | 🧑 [n17](#notes) |
| Unresolvable image degrades to a visible note      | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑               |
| **Equations**                                      |            |            |             |             |                    |                  |
| Inline `$x^2$` typesets as a native equation       | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Display `$$…$$` typesets and centres               | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ❌ [n18](#notes) |
| Equation opens in the built-in equation editor     | 🧑         | 🧑         | 🧑          | ➖          | 🧑 [n2](#notes)    | ➖               |
| Fractions, radicals, matrices, sub/superscripts    | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n19](#notes)   | ⚠️ [n19](#notes) |
| Inline math leaves the line leading alone          | 🧑         | 🧑         | 🧑          | 🧑          | ❌ [n20](#notes)   | ❌ [n20](#notes) |
| N-ary operators (∫, ∑) — see [note 1](#notes)      | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n1](#notes)    | ✅ [n1](#notes)  |
| Display math tracks the theme body size            | 🧑         | 🧑         | 🧑          | 🧑          | ❌ [n21](#notes)   | 🧑 [n21](#notes) |
| **Document-level**                                 |            |            |             |             |                    |                  |
| Title/author show under File → Info (or equiv.)    | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | 🧑 [n22](#notes) |
| Page size and margins as configured                | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ✅               |
| Landscape orientation                              | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑               |
| RTL document: mirrored lists, tables, punctuation  | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑               |
| Theme fonts resolve, fallback face used if missing | 🧑         | 🧑         | 🧑          | 🧑          | ⚠️ [n23](#notes)   | ⚠️ [n23](#notes) |
| File opens with no repair prompt                   | 🧑         | 🧑         | 🧑          | 🧑          | ✅                 | ⚠️ [n24](#notes) |

## Notes

One note per ⚠️, ❌ or qualified 🧑 cell, naming the application and what
exactly was seen. Everything numbered here was observed on 2026-08-18 in the
run described in the [Verification log](#verification-log), except note 1,
which predates it.

1. **N-ary operators carry an empty placeholder box — confirmed in LibreOffice,
   absent in Pages.** `$\int_0^1 x^2 dx$` produces an `<m:nary>` whose `<m:e/>`
   body is **empty**, with the integrand following as a sibling.
   **LibreOffice 26.2.5 draws that empty slot** as a hollow grey rectangle
   between the operator and the integrand, for **both** `\int` and `\sum`. At
   11 pt it is conspicuous — a reader notices it immediately, though the
   equation stays complete and fully editable. **Apple Pages does not draw it
   at all**; the same OMML renders with no box — visible once the display-math
   sizing of note 18 was corrected, since at the size Pages ships it the
   equation is too small to read either way. The defect is upstream: `temml`
   emits `<msubsup>` plus siblings and `mathml2omml` has nothing to hoist into
   `<m:e>`. Fixing it needs a semantic OMML rewrite in downword. Still unknown
   in Word, which is the reader that matters most here.

2. **Needs a human in the GUI; cannot be answered by looking at rendered
   output.** The Styles gallery, the navigation pane, the accessibility pane,
   "edit `Heading 2` and watch every H2 follow", Design → Style Set, "delete a
   footnote and watch the rest renumber", and "click an equation and confirm
   the equation editor opens" are all interactions, not appearances. The
   2026-08-18 pass rendered pages; it could not press anything. These rows stay
   🧑 in every column.

3. **H4, H5 and H6 are not reliably distinguishable, and `academic` collapses
   them completely.** Measured from the rendered PDFs (glyph box heights, both
   readers agree):

   | Level | `default` | `academic` |
   | ----- | --------- | ---------- |
   | H1    | 16 pt     | 14 pt      |
   | H2    | 14 pt     | 13 pt      |
   | H3    | 12 pt     | 12 pt      |
   | H4–H6 | 11 pt     | 12 pt      |
   | body  | 11 pt     | 12 pt      |

   In `default` the three deepest levels are body size and separated only by
   weight, italics and colour (H4 bold italic, H5 regular, H6 italic) — thin,
   but Word's own convention, and legible. In `academic` H4, H5 and H6 are the
   same size, the same weight and all italic: on the page they are
   **indistinguishable from one another**, and H1 is only 2 pt larger than body
   text. A six-level document set in `academic` reads as a three-level
   document. Not wrong output — a weak preset.

4. **Not exercised.** The corpus used an H1, not a frontmatter title block, so
   nothing carried the `Title` style.

5. **Bullet glyphs cycle every three levels**, in both readers: level 1 `•`
   (Symbol), level 2 `○` (Courier New "o"), level 3 `▪` (Wingdings), level 4
   back to `•`. That is what `theme.bulletLevels` declares and what Word itself
   does, so the row is ✅ — recorded here so the repeat at level 4 is not
   mistaken for a bug.

6. **Verified to four levels, not nine.** The corpus nested bullets four deep
   and ordered lists four deep. Both readers drew a clean ~0.5 in indent ladder
   per level with the level-1 indent restored correctly after the nesting
   closed. Levels 5–9 were never rendered, so the row cannot be ✅ as worded.

7. **Apple Pages draws the task-list checkboxes noticeably smaller and fainter**
   than the surrounding text — legible, but visibly undersized against
   LibreOffice's, which match the text size. Cosmetic.

8. **Not exercised.** No table in the corpus crossed a page break, so the
   repeating header row was never put to the test. Needs a deliberately long
   table.

9. **Column widths fit, with a knock-on in `academic`.** No overflow in either
   reader, and the long cell wrapped inside its column as intended. Because
   `computeColumnWidths` reads `sizes.body`, `academic` (12 pt) narrows the
   third column enough that the header "Right aligned" wraps onto two lines
   where `default` (11 pt) keeps it on one. Documented behaviour, but it does
   change how the table looks between presets.

10. **The code-block panel has essentially no inset.** Shading spans the whole
    block and the full text column in both readers, and the syntax colours
    survive — but glyphs sit flush against the left edge of the shading
    (measured ~1.5 pt) and close to the top edge (~3 pt), so `export` and the
    closing `}` touch the panel. The panel also bleeds a hair into the left
    margin. Purely cosmetic, and the one thing in the code block a reader's eye
    catches. Pages additionally applies paragraph line spacing inside the
    block, so the listing is looser there than in LibreOffice.

11. **Apple Pages ignores `w:docDefaults`, so `academic` loses its 1.5
    leading.** The preset sets `<w:spacing w:line="360" w:lineRule="auto"/>` in
    `docDefaults` only — the `Normal` style itself carries no `pPr`. LibreOffice
    honours it and sets the body 1.5-spaced. Pages does not: body paragraphs
    come out single-spaced, and the most recognisable trait of the academic
    preset is simply gone. **Root cause confirmed by experiment**: copying that
    same `w:spacing` onto the `Normal` style definition and re-importing made
    Pages render the 1.5 leading correctly. See
    [Converter defects found in the same run](#converter-defects-found-in-the-same-run).

12. **Links render as links; click-through was not exercised.** Both readers
    style external and intra-document links in the hyperlink colour with an
    underline. LibreOffice's PDF export carries a real
    `/URI (https://example.com/)` annotation, which is why its external-link row
    is ✅. Pages' export contains link annotations but no recoverable `/URI`,
    and the extraction is not reliable enough on Pages' PDF structure to call
    that a failure — so it stays 🧑. Nobody clicked an `#anchor` in either
    application, so the bookmark-jump row stays 🧑 for both.

13. **LibreOffice's TOC could not be updated in this environment.** Updating it
    needs Tools → Update → Indexes and Tables, i.e. the GUI. Three headless
    routes were tried — a Basic macro via `vnd.sun.star.script:` and via
    `macro:///`, and a Python UNO client against `--accept=socket` — and all
    three failed under the sandbox this pass ran in (the UNO client was killed
    with SIGKILL). Un-updated, the field renders exactly as the CLI's notice
    warns: the "Contents" heading followed by empty space. Row stays 🧑.

14. **Apple Pages _does_ run the TOC field — the ➖ premise for this column was
    wrong — and the result is broken.** Pages populated all fifteen entries on
    import with correct page numbers and correct per-level indents, with no
    prompting. But every entry wraps: the heading text takes one line and its
    page number falls to the **next line, flush left**, because the
    right-aligned tab stop is not honoured. There is no dot leader, and the
    entries are set in a sans-serif face while the rest of the document is
    serif. Unusable as a table of contents, and worse than showing nothing.
    Whether the entries are clickable was not tested, so that row stays 🧑.

15. **`CodeChar` in a footnote body: verified. Math in a footnote body: not
    tested.** Both readers render inline code inside a footnote in a shaded
    monospace face, alongside bold, at the reduced footnote size. The corpus
    had no equation inside a note, so half of this row is unanswered.

16. **The column-fit cap was never exercised.** The test image is 320×160 px,
    narrower than the text column. Both readers placed it at its natural
    3.33 in width and centred it, with no upscaling — correct restraint, but it
    proves nothing about an image _wider_ than the column, which is what the
    row is really asking. Aspect ratio was exact in both (measured 1.98:1
    against a 2.00:1 source at 110 dpi; the square in the middle of the test
    card stayed square).

17. **Out of reach of the CLI.** Mermaid needs a browser DOM to measure text,
    so the CLI leaves a ```mermaid fence as a code block. Testing this row
    needs a fixture built by the browser path.

18. **Apple Pages renders display equations at roughly 2 pt, left-aligned —
    effectively unreadable.** The OMML itself is fine: blown up, the integral
    and the summation typeset correctly, with limits, fractions and
    superscripts all in the right places. They are simply drawn at about a
    fifth of body size, and the `w:jc="center"` on the paragraph is ignored.
    Inline math in the same document is correct and correctly sized, so this is
    specific to the display case. **Root cause confirmed by experiment**: the
    display-math paragraph is emitted as `<w:pPr><w:jc w:val="center"/></w:pPr>`
    with **no `<w:pStyle>`**, where the inline case carries
    `<w:pStyle w:val="Normal"/>`. Adding `<w:pStyle w:val="Normal"/>` to that
    `pPr` and re-importing fixed **both** symptoms in Pages — full size and
    properly centred — with no change at all to the LibreOffice rendering
    (glyph boxes and x-positions identical to the byte). See
    [Converter defects found in the same run](#converter-defects-found-in-the-same-run).

19. **Matrix delimiters do not grow to the height of the matrix.** In both
    readers a 2×2 `pmatrix` is wrapped in parentheses drawn at close to
    single-line size, so they look thin and undersized against the two rows
    they enclose. Everything else in this row is right: fractions stack with a
    proper rule, nested radicals draw a vinculum over the whole fraction,
    sub- and superscripts land correctly. In LibreOffice the `=` in
    `∑ □ k² = n(n+1)(2n+1)/6` also sits high, near the numerator rather than on
    the fraction's axis. Legible, but not typeset well. Whether this is the
    OMML or the reader's math layout was not determined.

20. **An inline fraction inflates the whole line's leading.** `$\frac{a}{b}$`
    inside a sentence is set as a full display-style stacked fraction rather
    than a text-style one, so the line it sits in grows and the paragraph
    develops a visible gap above and below. Both readers do it; `academic`
    (1.5 leading) makes it worse. A user writing `$\frac{a}{b}$` mid-sentence
    will notice their paragraph spacing break.

21. **Display equations do not track the theme body size.** Measured across the
    two presets in LibreOffice, the same equation renders at **identical size
    and identical x-positions** in `default` (11 pt body) and `academic` (12 pt
    body) — the equation glyphs are 12 pt in both. So an `academic` document
    happens to match, and a `default` document has equations slightly larger
    than its own body text. Adding `w:pStyle` (note 18) did **not** change this
    in LibreOffice, so this one looks like LibreOffice's own OMML sizing rather
    than something downword controls through `pPr`. Unmeasured in Pages, whose
    display-math sizing is broken for a different reason.

22. **Inconclusive for Pages.** LibreOffice carried `dc:title` and `dc:creator`
    from `docProps/core.xml` straight into its PDF export, which is good
    evidence the metadata arrived. Pages' PDF export names the file and writes
    no author — but Pages' exporter may simply not propagate those fields, so
    this says nothing about whether Pages imported them. Needs a look at
    Pages' own document inspector.

23. **The `default` theme renders as a serif on any machine without Office
    fonts.** `fonts.body` is `Aptos` with `Calibri` in the `w:cs` fallback slot.
    A stock macOS has neither, so **both** readers fell through to their own
    default: LibreOffice substituted Liberation Serif, Pages a Times-like
    serif. The preset is meant to be a sans, and it comes out a serif — which
    also makes `default` and `academic` look far more alike than intended.
    `academic` resolved exactly as designed (Times New Roman, present on
    macOS), as would `github` (Arial fallback). Worth considering whether the
    `default` fallback chain should name a face that exists outside Office.

24. **Pages opened and exported the file with no blocking dialog**, across five
    documents. That is not quite the same as "no repair prompt": screen capture
    was unavailable in this environment (`screencapture` failed with "could not
    create image from display" — macOS screen-recording permission), so a
    non-modal import-warnings notice cannot be ruled out. LibreOffice converted
    all five headlessly with no warning on stderr, hence ✅ there.

## Converter defects found in the same run

Two of these are reader-fidelity notes above, traced back to their cause in the
generated OOXML and confirmed by patching the XML by hand and re-importing.
They are recorded here because they are downword's to fix, not the reader's.

1. **Display-math paragraphs carry no `w:pStyle`.** Emitted as
   `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><m:oMath …>`, where the inline
   case emits `<w:pStyle w:val="Normal"/>`. Apple Pages then falls back to a
   default that renders the equation at roughly 2 pt and drops the centring
   (note 18). Adding `<w:pStyle w:val="Normal"/>` fixed both symptoms in Pages
   and changed nothing in LibreOffice. Cheap, safe, and it makes native
   equations usable in Pages.

2. **Theme leading lives only in `w:docDefaults`.** The `Normal` style is
   emitted as a bare `<w:style>` with no `<w:pPr>`, so a reader that ignores
   `docDefaults` — Pages does — loses the preset's line spacing entirely
   (note 11). Writing the spacing onto `Normal` as well restored it.

3. **With `--math off`, a lone `=` line inside `$$ … $$` is swallowed as a
   setext heading.** This markdown

   ```md
   $$
   \begin{pmatrix} a & b \\ c & d \end{pmatrix}
   \begin{pmatrix} x \\ y \end{pmatrix}
   =
   \begin{pmatrix} ax + by \\ cx + dy \end{pmatrix}
   $$
   ```

   converts, with math disabled, to a paragraph carrying **`w:pStyle="Heading1"`**
   — verified in `word/document.xml` — because CommonMark reads `text` followed
   by a line of `=` as a setext H1. The equation is then rendered large, bold
   and coloured, picks up a heading bookmark and lands in the TOC. Turning math
   off is documented as leaving the TeX alone ("a dollar sign is just a dollar
   sign"), but it is not inert: it can restructure the document. Worth either
   guarding `$$ … $$` spans when `math: "off"` or documenting the hazard.

## What CI already proves

Be precise about this, because it is easy to over-read.

**`docx-validity` job — every push and PR** (`scripts/docx-validity.mjs`, run
with `--require-soffice`):

1. **Structural check, no LibreOffice needed.** Every generated `.docx` starts
   with the zip magic `50 4b 03 04` and contains the two mandatory OOXML parts,
   `[Content_Types].xml` and `word/document.xml`.
2. **Headless LibreOffice opens each file and converts it to PDF**, and the PDF
   must be non-empty and start with `%PDF`. Covers every fixture the core and
   CLI test suites emit plus the `examples/` output — 89 files at the last run.
3. **Positive control** on the forced `MS Word 2007 XML` import filter, so a
   renamed or missing filter cannot quietly make step 4 vacuous.
4. **Negative control** against a deliberately corrupt fixture. If LibreOffice
   converts it successfully, the gate is declared broken and CI fails. A gate
   that never says "no" is worse than no gate.

**What that does _not_ prove.** It proves the container is well-formed and one
implementation can parse it end to end. It says nothing about whether a heading
looks like a heading, whether the header row repeats, whether numbering
restarts, or whether an equation typesets. "Converts to a PDF" and "renders
correctly" are different claims, and only the first is automated.

**One-off spot check, not CI.** During Phase 2 verification a combined
all-features fixture was rendered by LibreOffice 26.2.5 and read back with
`pdftotext`. The TOC, both tables, `∫₀¹x²dx=1/3`, `∑ₖ₌₁ⁿk²`, footnotes at the
page foot, task-list checkboxes and three ordered lists each restarting at 1
were all present in the extracted text. That is real evidence and it is text-
level only — it did not check layout, colour, style application or the styles
gallery, it was run once by hand, and it is not re-run by CI. It is recorded
here so it is neither lost nor mistaken for a passing matrix row.

**Never opened in Microsoft Word.** Not on Windows, not on macOS, not on the
web. Word is the primary target and the one implementation with zero coverage.
Until somebody works the protocol below, every Word column is an open question.

## Verification protocol

1. **Build the corpus.**

   ```sh
   pnpm install
   pnpm test        # emits fixtures into packages/*/tests/__fixtures__/out
   pnpm build
   pnpm examples
   ```

   Use the combined all-features fixtures — they are the ones designed to put
   every row of this matrix into one document. Add a long document (past one
   page) for the repeating-header-row and page-number rows.

2. **Open each file in each application from the column headings.** Open the
   original `.docx` directly. Do not round-trip through another converter, and
   do not import into Google Docs and then export — that measures the importer,
   not downword.

3. **Answer every row deliberately.** A row is ✅ only if you looked at the
   thing it names. If you did not test it, leave it 🧑 — an honest gap beats a
   guessed tick.

4. **For each ⚠️ or ❌, add a numbered note** saying what you saw, and open an
   issue with the fixture attached and a screenshot.

5. **Record the exact versions** in the log below, then commit the updated
   matrix.

### Rows that need a deliberate action, not just a look

- **TOC field.** It ships empty by design; OOXML stores the instruction, not
  the entries. Answer yes to Word's "update fields?" prompt on open, or select
  the field and press **F9**. In LibreOffice: Tools → Update → Indexes and
  Tables. Google Docs and Pages do not run fields at all — hence ➖.
- **Style restyling.** Change the `Heading 2` style definition once and confirm
  every H2 follows. Then apply a different Style Set from the Design tab. This
  is the row that distinguishes real styles from direct formatting, and it is
  the whole premise of the library.
- **Footnote renumbering.** Delete the first footnote reference and confirm the
  rest renumber.
- **Repeating header row.** Requires a table that crosses a page break.
- **Equation editing.** Click an equation and confirm it opens in the equation
  editor as an editable formula rather than selecting a picture.
- **Missing theme fonts.** Test at least one application that does not have
  Aptos installed, to exercise the `w:cs` fallback slot.

## Verification log

| Date       | Application and exact version                                   | Platform                        | Fixture(s)                                                                                                                       | By                                                    |
| ---------- | --------------------------------------------------------------- | ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------- |
| 2026-08-18 | LibreOffice 26.2.5.2 (cd7284b4cbbfeb507e630c1aac019f4157393acb) | macOS 15 (Darwin 25.5.0), arm64 | `all-features.md` → 5 `.docx` via `packages/cli` (`default` / `academic` × math off / `--math omml`, one `--toc --page-numbers`) | Automated render inspection (Claude, agent run)       |
| 2026-08-18 | Apple Pages 14 (macOS bundled)                                  | macOS 15 (Darwin 25.5.0), arm64 | the same 5 `.docx`, plus 2 hand-patched variants used to isolate notes 11 and 18                                                 | Automated render inspection (Claude, agent run)       |
| —          | _Microsoft Word — never opened, any platform_                   | —                               | —                                                                                                                                | **Open. Needs a human on a machine with Word.**       |
| —          | _Word Online — never opened_                                    | —                               | —                                                                                                                                | **Open. Needs a human with a Microsoft 365 account.** |
| —          | _Google Docs — never opened_                                    | —                               | —                                                                                                                                | **Open. Needs a human with a Google account.**        |

### Method used on 2026-08-18

Recorded so the two filled columns can be reproduced or challenged.

1. Built the CLI from source and generated a single all-features corpus with
   it — H1–H6, bold/italic/strike/inline code, a three-alignment table with
   rich cells plus two adjacent tables, ordered and bullet and task lists
   nested four deep, **two sibling ordered lists**, tagged and untagged fenced
   code, a nested blockquote, two footnotes, a `data:` image, a horizontal
   rule, intra-document links, and inline plus display math including `\int`,
   `\sum`, `pmatrix` and nested radicals.
2. **LibreOffice**: `soffice --headless --convert-to pdf` on each `.docx`.
3. **Apple Pages**: `open -a Pages <file>.docx`, then AppleScript
   `export document 1 … as PDF`. The original `.docx` was opened directly in
   Pages — the PDF is only how Pages' own layout was captured, not a converter
   in the path. `screencapture` was unavailable (no screen-recording
   permission), so no window screenshots exist.
4. Rasterised every page with `pdftoppm -png -r 110`, plus 200–600 dpi crops of
   the equations, the code block, the heading stack, the inline marks and the
   bullet ladder.
5. **Looked at every page image** and judged the rows above against them.
6. Measured, rather than eyeballed, wherever a number settled the question:
   glyph box heights and x-positions via `pdftotext -bbox` (note 3, note 21),
   substituted faces via `pdffonts` (note 23), link annotations by scanning the
   PDF objects (note 12), footer page numbers by cropped text extraction.
7. For notes 11 and 18, formed a root-cause hypothesis from
   `word/document.xml` and `word/styles.xml`, patched the XML by hand, re-zipped,
   re-imported into Pages, and re-rendered to confirm the fix and check
   LibreOffice for regression.

### What this pass did not establish

- **Anything at all about Microsoft Word**, on any platform. Word is not
  installed on the machine this ran on. It remains the primary target and the
  one implementation with zero coverage.
- **Anything about Word Online or Google Docs.**
- **Any row that needs an interactive action** — Styles gallery, navigation
  pane, accessibility pane, restyling `Heading 2`, Design → Style Set, F9,
  footnote renumbering, opening an equation in the equation editor, clicking a
  hyperlink or an `#anchor` (note 2, note 12).
- **LibreOffice's TOC after an update** (note 13).
- **Repeating header rows** — no table crossed a page break (note 8).
- **Oversize images, SVG twins, mermaid diagrams, unresolvable images, landscape
  orientation, RTL documents, the `Title` style, and math inside a footnote** —
  no fixture covered them (notes 4, 15, 16, 17).
