# r/webdev

> **Draft. Do not post until every gate in [`../README.md`](../README.md) is
> closed** — this post links the repo and the npm package directly.

> **Timing overrides the schedule.** r/webdev restricts sharing your own project
> to **Showoff Saturday**, with the matching flair. So the Day-4 slot in
> [`../README.md`](../README.md) means "the first Saturday at least three days
> after the Show HN". If Day 0 is a Tuesday, Day 4 is that Saturday and it lines
> up. If Day 0 is a Wednesday, use the following Saturday (Day 10) and swap
> r/ObsidianMD earlier. **Check the current rule before posting** — it is
> enforced, and posting a project mid-week gets it removed.

**Voice here:** developers who have shipped things and have working sarcasm
detectors. The product pitch is not interesting to them; the engineering is.
Lead with the problem that was actually hard, be specific about the bugs, and
put the honest gap at the end as a request rather than a disclaimer.

---

## Title

```
I built a Markdown → .docx converter that runs entirely client-side. The hard part wasn't generating the file, it was proving it isn't lying.
```

Alternative, shorter:

```
Showoff Saturday: client-side Markdown → .docx, and how I test a file format I can't render
```

**Flair:** `Showoff Saturday`.

---

## Post body

---

Client-side Markdown → `.docx`. Paste, download, nothing uploaded — the whole
conversion is a Web Worker in your tab.

Tool: https://ksprtech.com/tools/markdown-to-word
Source (MIT): https://github.com/kspr-technologies/downword

The reason it exists is narrow: LLMs answer in markdown, the person who asked
for the output works in Word, and every converter I tried either wanted an
upload or handed me an HTML file with a `.docx` extension that Word offered to
repair. The thing I actually cared about, though, is a distinction most
converters get wrong: direct formatting versus a style. Eighteen-point bold on a
paragraph is not a heading — it renders identically and behaves completely
differently, because Word's outline, navigation pane, TOC field and Style Sets
all read the style definition, and there is not one. So this emits a genuine
`styles.xml` keyed to Word's own built-in style IDs.

markdown-it (CommonMark + GFM + footnotes) into a small document model, then the
`docx` package to build the OOXML package. That part was fine. Here is the part
that was not.

### How do you test a format you cannot render?

I can generate a `.docx` all day. I cannot look at one, because CI has no Word,
and neither do I. So the question is what a green build is actually allowed to
claim.

What I settled on: every generated file in the test suite — 89 at the last count
— gets opened by **headless LibreOffice and converted to PDF**, and the PDF has
to be non-empty and start with `%PDF`. Two controls around it:

- a **positive control** on the forced `MS Word 2007 XML` import filter, so a
  renamed or missing filter cannot quietly make the whole job vacuous;
- a **negative control** against a deliberately corrupt fixture. If LibreOffice
  successfully converts the corrupt one, the gate declares itself broken and
  fails the build. A gate that never says no is worse than no gate, and I only
  believe that after watching one pass for two weeks while silently doing
  nothing.

And then the discipline part, which I think is the actually transferable idea:
**that gate proves the container is well-formed and one implementation parses it
end to end. It does not prove a heading looks like a heading.** Those are
different claims and it is very easy to let a green tick quietly become the
bigger one. So the repo has a fidelity matrix — six word processors down the
columns, forty-odd features down the rows — where **every single cell currently
reads "awaiting manual verification"**, because nobody has done it. It is a
scoreboard of unknowns rather than a table of ticks. Filling in a guess would
convert an unknown into a promise, which is the failure mode I was trying to
avoid in the first place.

### Three bugs that were only findable by being paranoid

- **Every `wp:docPr id` was `1`.** The spec says that id is document-unique. The
  `docx` package builds its counter inside the constructor, so every drawing got
  the same one. Invisible until a document has two images — and no fixture had
  two images, which is exactly why nothing caught it until a cross-feature test
  put every feature into one document.
- **XML injection through the maths path.** The MathML → OMML converter decodes
  the XML entities in its input and writes the result back raw. So `$<w:p><w:r>
<w:t>x</w:t></w:r></w:p>$` — typed by a user, escaped correctly by the TeX
  renderer — got _un_-escaped on the way out and landed in `word/document.xml`
  as live WordprocessingML. Fixed at the source with `disableDecode`, and then
  again with an allowlist validator that re-reads every generated fragment and
  refuses anything that is not namespaced OMML with escaped text. Single
  left-to-right pass, no backtracking regex, so a hostile equation cannot turn
  the check into a DoS.
- **A missing lookup-table row, spelled `undefined` into the output.** The same
  converter maps MathML's `mathvariant` onto an OOXML enumeration and has no
  entry for `normal`, so `\Gamma` shipped `<m:sty m:val="undefined"/>` — a value
  outside the enumeration, which is precisely what makes Word declare a part
  corrupt and offer to repair the file. One capital Greek letter, one repair
  prompt, and nothing in the type system that could have caught it.

### Performance, honestly

Profiling put essentially all of it in one place, and it is not mine. Parsing
and rendering are linear and negligible; **the zip-and-serialise step inside
`docx` is the whole curve, at roughly n^1.8**. Node 22 on Apple silicon, total
wall clock: 2.6 s for 1 MB of markdown, 7.3 s for 2 MB, 26.1 s for 4 MB, and at
8 MB it eats the default heap and dies. Which is why the hosted tool refuses
anything over 2 MB and tells you the number instead of spinning — measured in
Chrome, a 2 MB document is 3.1 s end to end, and because the whole thing sits in
a Worker the worst main-thread gap over that run was 117 ms. Getting past this
means replacing the packer, and that is the biggest open item in the repo.

Bundle side, since somebody always asks: heavy dependencies each live behind
their own subpath entry, and a test bundles the main entry and **fails if
`highlight.js`, `katex`, `temml`, `mathml2omml` or `mermaid` show up anywhere in
its import graph**. size-limit holds the `.` entry to 150 kB min+gzip excluding
`docx`; it currently measures 74.2 kB. Nobody who never writes an equation
downloads a TeX parser.

### The gap, which is where you come in

**This has never been opened in Microsoft Word.** Not Windows, not Mac, not the
web. Word is the primary target and the one implementation with zero coverage.
If you have Word and ten minutes, opening one of the test fixtures and filling
in a column of `docs/fidelity-matrix.md` is worth more to this project than any
amount of code I could write this month. The protocol is in that file, including
the rows that need a deliberate action rather than a look — change the
`Heading 2` definition and confirm every H2 follows, delete a footnote and
confirm renumbering, check the header row repeats across a page break.

And the obligatory: **if you can install pandoc, install pandoc.** It is twenty
years old, it is better at this than I am, and `--reference-doc` restyles from
your own corporate template, which I cannot do at all. The gap I am standing in
is the browser tab on the locked-down laptop, and that is the only place I would
claim this wins.

Happy to go into any of it.

---

## Predictable questions here

**"Why not WASM pandoc?"** — Legitimate, and worth a real answer: size, and
shipping a general-purpose document engine for one conversion. Do not be
dismissive; pandoc-wasm is a reasonable choice for a different set of
constraints.

**"Why `docx` and not hand-rolled XML?"** — Because `styles.xml`,
`numbering.xml`, `footnotes.xml`, content types and relationships are a lot of
surface to get subtly wrong, and Word's repair prompt is unforgiving. The
package handles the package; the interesting decisions are all above it.

**"n^1.8, did you actually measure that or is it vibes?"** — Measured, on a
mixed prose/list/table corpus, and the table is in the README. Say the machine.

**"Show HN was three days ago"** — Yes, and say so if asked. It is a different
audience and a different post, and being cagey about it is worse than
volunteering it.

## What NOT to say in this sub

- Do not repeat the Show HN comment. Several people here will have read it.
  This post is about the testing strategy and the bugs; that one was about the
  product.
- Do not go deep on the OMML pipeline — that is the dev.to article three days
  later ([`../devto.md`](../devto.md)), and burning it here leaves that piece
  with nothing.
- No star-count or download-count claims, and do not ask for stars.
