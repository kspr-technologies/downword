# Fidelity matrix

How a downword `.docx` actually renders in the word processors people open it
with.

> [!IMPORTANT]
> **Nothing in the matrix below has been verified by a human yet.** Every cell
> is 🧑 — "awaiting manual verification". This file is the protocol and the
> scoreboard, filled in with the rows and columns to test, and deliberately
> **not** filled in with results. Do not guess a cell. A matrix with invented
> ticks in it is worse than no matrix, because it converts an unknown into a
> promise.
>
> What _is_ already machine-verified is recorded in
> [What CI already proves](#what-ci-already-proves), and it is a much smaller
> claim than this matrix makes: CI proves the files **open**, not that they
> **look right**.

## Legend

| Mark | Meaning                                                                            |
| ---- | ---------------------------------------------------------------------------------- |
| 🧑   | Awaiting manual verification. **The current state of every cell.**                 |
| ✅   | Verified by a human: renders as intended.                                          |
| ⚠️   | Verified by a human: renders, with a caveat. **Must** carry a numbered note below. |
| ❌   | Verified by a human: broken, missing or wrong. **Must** carry a numbered note.     |
| ➖   | Not applicable — the application has no such feature to get right.                 |

When you fill a cell in, record the exact application version and platform in
[Verification log](#verification-log). "Word" is not a version.

## The matrix

Columns are the six readers a document produced by this library is most likely
to be opened in.

| Feature                                            | Word (Win) | Word (Mac) | Word Online | Google Docs | LibreOffice Writer | Apple Pages |
| -------------------------------------------------- | ---------- | ---------- | ----------- | ----------- | ------------------ | ----------- |
| **Structure**                                      |            |            |             |             |                    |             |
| `Heading1`–`Heading6` appear in the Styles gallery | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Headings populate the navigation / outline pane    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Editing `Heading 2` once restyles every H2         | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Design → Style Set restyles the whole document     | 🧑         | 🧑         | 🧑          | ➖          | 🧑                 | ➖          |
| `Title` style on the frontmatter title block       | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Lists**                                          |            |            |             |             |                    |             |
| Bullet lists, correct glyph per level              | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Sibling ordered lists each restart at 1            | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Nested lists indent correctly to 9 levels          | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Ordered list numbering format per level (1/a/i)    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Task lists `- [x]` show checked/unchecked boxes    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| `ListParagraph` style is applied and editable      | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Tables**                                         |            |            |             |             |                    |             |
| Table renders with borders and correct cell text   | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Column alignment (`:--`, `:-:`, `--:`) honoured    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Header row repeats across a page break             | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Two adjacent tables stay two tables                | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Column widths fit the text column, no overflow     | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Text and code**                                  |            |            |             |             |                    |             |
| Bold / italic / strikethrough                      | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| `Quote` style, left rule, nested quotes indent     | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Inline `CodeChar` — monospace, shaded              | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Fenced `CodeBlock` — shading spans the block       | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Syntax highlighting colours survive                | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Horizontal rule (`---`) draws a rule               | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Tab size / leading match the theme                 | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Links and navigation**                           |            |            |             |             |                    |             |
| External hyperlinks are clickable                  | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| `[text](#anchor)` jumps to the heading bookmark    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| TOC field populates after Update Field / F9        | 🧑         | 🧑         | 🧑          | ➖          | 🧑                 | ➖          |
| TOC entries link to their headings                 | 🧑         | 🧑         | 🧑          | ➖          | 🧑                 | ➖          |
| Page numbers (`PAGE` / `NUMPAGES`) compute         | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Footnotes**                                      |            |            |             |             |                    |             |
| Superscript marker in the body, note at page foot  | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Word renumbers notes after one is deleted          | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| `CodeChar` and math inside a note body survive     | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Images and diagrams**                            |            |            |             |             |                    |             |
| Raster image, sized to the text column             | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Aspect ratio exact, no stretch                     | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| SVG + raster twin: vector drawn where supported    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Alt text reaches the accessibility pane            | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Mermaid diagram picture + italic caption           | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Unresolvable image degrades to a visible note      | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Equations**                                      |            |            |             |             |                    |             |
| Inline `$x^2$` typesets as a native equation       | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Display `$$…$$` typesets and centres               | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Equation opens in the built-in equation editor     | 🧑         | 🧑         | 🧑          | ➖          | 🧑                 | ➖          |
| Fractions, radicals, matrices, sub/superscripts    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| N-ary operators (∫, ∑) — see [note 1](#notes)      | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| **Document-level**                                 |            |            |             |             |                    |             |
| Title/author show under File → Info (or equiv.)    | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Page size and margins as configured                | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Landscape orientation                              | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| RTL document: mirrored lists, tables, punctuation  | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| Theme fonts resolve, fallback face used if missing | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |
| File opens with no repair prompt                   | 🧑         | 🧑         | 🧑          | 🧑          | 🧑                 | 🧑          |

## Notes

Numbered notes belong here, one per ⚠️ or ❌ cell, naming the application and
what exactly went wrong.

1. **N-ary operators carry an empty placeholder box.** Not awaiting
   verification — this one is already known and reproducible today.
   `$\int_0^1 x^2 dx$` produces an `<m:nary>` whose `<m:e/>` body is **empty**,
   with the integrand following as a sibling, so a reader draws the dotted
   "empty slot" box after the integral sign. The equation is complete and fully
   editable; the box is cosmetic. The defect is upstream: `temml` emits
   `<msubsup>` plus siblings and `mathml2omml` has nothing to hoist into
   `<m:e>`. Fixing it needs a semantic OMML rewrite in downword. Verify how
   intrusive it looks per application; do not mark the row ❌ for this alone
   unless the equation is actually unreadable.

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

| Date | Application and exact version | Platform | Fixture(s) | By  |
| ---- | ----------------------------- | -------- | ---------- | --- |
| —    | _no verification runs yet_    | —        | —          | —   |
