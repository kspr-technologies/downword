# Show HN

> **Draft. Do not post until every gate in [`README.md`](README.md) is closed** —
> in particular the GitHub repo (gate 1) and the npm release (gate 2). This
> comment links to both.

**When:** Tuesday or Wednesday, 06:00–08:00 EST. Nothing else goes out that day.
Be at the keyboard for four hours afterwards.

**Submit at:** <https://news.ycombinator.com/submit>

---

## Title

```
Show HN: I made a Markdown → Word converter that runs entirely in your browser
```

78 characters, under HN's 80-character limit. If the arrow gets mangled on
submit, use this instead — 72 characters, same meaning:

```
Show HN: A Markdown to Word converter that runs entirely in your browser
```

## URL

```
https://ksprtech.com/tools/markdown-to-word
```

Bare, no UTM. See the note at the end of [`README.md`](README.md).

---

## The first comment

Post this yourself, immediately after submitting, from the same account.

---

I kept hitting the same twenty seconds of annoyance. I'd ask an LLM for a report
or a spec, get a good answer back in markdown, and then need to hand it to
somebody who works in Word. The copy button gives you the markdown source, so
Word shows you literal `## Heading` and literal `| pipes |`. Copying the
_rendered_ answer out of the chat UI is better than people expect — Word does
map `<h1>`–`<h6>` onto its heading styles — but it also stamps the web page's
fonts, sizes and colours directly onto every run, so a Style Set change
afterwards has nothing to grip, and code blocks lose their monospace. And the
converters that come up first all want you to upload the document.

So: downword. Paste markdown, get a `.docx`. The conversion happens in a Web
Worker in your tab.

**What it actually does differently.** Most converters make a paragraph _look_
like a heading — 18pt, bold, blue. Word renders that faithfully and treats it as
body text: no navigation pane entry, no TOC entry, and Design → Style Set can't
touch it, because a Style Set rewrites style _definitions_ and there's no style
there to rewrite. This writes a real `styles.xml` and references Word's own
built-in style IDs, so `Heading 2` is genuinely `Heading 2`. Edit it once,
every H2 follows. References → Table of Contents fills itself in.

The other one is equations. `$x^2$` becomes native `<m:oMath>` — click it in
Word and it opens in the equation editor. Not a PNG of an equation, which is
what most tools that handle math at all give you, and which can't be corrected,
searched, restyled or read aloud.

**On privacy, precisely.** The _converter_ makes zero network requests — I
enforce that with a `fetch` spy in the test suite, and remote images are an
explicit opt-in that's off by default. The _page_ is a different claim and I'm
not going to fudge it: it's on a normal marketing site and it loads Google
Analytics like every other page there. Clicking Download fires one event
containing the theme name, three booleans, two byte counts, a duration and a
count of warnings. Numbers and closed enums. Your markdown isn't in it and never
leaves the tab. If that's still one beacon too many, the library is MIT and runs
fine with the tab offline.

**Stack.** TypeScript. markdown-it (CommonMark + GFM + footnotes) into a small
document model of my own, then the `docx` package to build the OOXML. Math is
temml → MathML → mathml2omml → raw OMML, injected through
`ImportedXmlComponent`, because `docx`'s Math builders have no matrix at all —
that pipeline is most of the interesting engineering and it's the thing I'd
write up if anyone wants it. temml and mathml2omml sit behind a subpath entry
point and load lazily, so a document with no equations never downloads a TeX
parser; core is 74.2 kB min+gzip excluding `docx`, against a 150 kB budget
enforced in CI.

**Where it's weak, and you'll find all of this in about a minute:**

- **Packing is super-linear.** `docx`'s serialise-and-zip step dominates and
  grows roughly as n^1.8. On Node 22 / Apple silicon: 1 MB in 2.6 s, 2 MB in
  7.3 s, 4 MB in 26.1 s, and 8 MB exhausts the default heap. The hosted tool is
  capped at 2 MB of input and says so; in Chrome a 2 MB document took 3.1 s end
  to end with a worst main-thread gap of 117 ms. Fixing this properly means not
  going through that packer.
- **Remote images fail on plenty of hosts.** Turning egress on makes it a
  cross-origin `fetch`, and a host that doesn't send
  `Access-Control-Allow-Origin` refuses it. `fetch` rejects with a deliberately
  vague `TypeError`, so I can't even tell you it was CORS. It degrades to a
  visible `[image: alt]` placeholder rather than losing the document.
- **Mermaid is browser-only.** It measures text by laying it out in a DOM, so
  there's no headless path that isn't a headless browser. In Node the fence
  stays a code block.
- **∫ and ∑ carry an empty placeholder box.** The equation is complete and
  editable; there's a cosmetic dotted slot next to it. It's upstream — temml
  emits `<msubsup>` plus siblings and mathml2omml has nothing to hoist into
  `<m:e>` — and fixing it needs a semantic OMML rewrite I haven't done.
- **The hosted tool passes no syntax highlighter,** so fenced code arrives as
  monospace in one colour. The library has a highlight.js adapter; the website
  doesn't wire it up.
- **It has never been opened in Microsoft Word.** CI opens every generated file
  in headless LibreOffice and converts it to PDF. The import filter is itself
  under a positive control, and a deliberately corrupt fixture has to fail, or
  the gate declares itself broken — so I know the container is well-formed and
  one real implementation parses it end to end. That is a much smaller claim
  than "renders correctly", and I've kept them separate.
  `docs/fidelity-matrix.md` is a scoreboard with every cell still marked
  "awaiting manual verification", because it is.

**And to be clear about pandoc:** if you can install pandoc, install pandoc.
It's twenty years old, it's excellent, and it beats this on formats, on
citations, on Lua filters, and especially on `--reference-doc`, which restyles
from your organisation's actual template — something I can't do at all, since
downword can't read a `.docx`. What it can't do is run in a tab on a locked-down
work laptop, and that's the entire gap I'm standing in.

Repo: https://github.com/kspr-technologies/downword — MIT. There's a CLI
(`npx downword-cli notes.md`) and a library if you want markdown → `.docx` inside
your own app.

If anyone here has Word on Windows and ten minutes, filling in one column of
that fidelity matrix is the single most useful thing anyone could do to this
project.

---

## Prepared answers

**Not posted.** These are for you, so that at 06:40 EST you're editing a reply
instead of writing one. Every one of these has been asked of a comparable Show
HN. Answer in your own words; do not paste these verbatim.

**"Why not just use pandoc / pandoc-wasm?"**

> If you can install pandoc, you should. For the browser case: pandoc compiled
> to WASM is a large download and I'd still be shipping a general-purpose
> document engine to do one conversion. This entry is 74.2 kB min+gzip with
> `docx` excluded, and `docx` on top of that. But it's a legitimate design
> choice and pandoc-wasm is a real option
> — if what you need is fifty formats, that's the answer, not this.

**"Doesn't Word already open .md files?" / "Just paste it."**

> For a two-paragraph answer, pasting is genuinely the right tool and I say so in
> the README. The difference shows up on a long document: pasting the rendered
> answer stamps the web page's formatting onto every run, so restyling
> afterwards has nothing to grip. Pasting the source gives you literal `##`.

**"Why not HTML with a .docx extension, or altChunk?"**

> Because that's the thing I was annoyed by. An HTML file renamed `.docx` gets
> Word offering to repair it, and `altChunk` hands the import off to Word's HTML
> converter, so you inherit exactly the direct-formatting problem I'm trying to
> avoid — and it doesn't work in Word Online or LibreOffice. This builds a real
> OOXML package: `styles.xml`, `numbering.xml`, `footnotes.xml`.

**"How do I know it doesn't upload my document?"**

> Three answers, in increasing order of how much you should trust them: the
> network tab; the test suite stubs `fetch` to throw and asserts zero calls on
> the default path; the source is MIT and the library runs with the tab offline.
> I'd rather you check than believe me.

**"You load Google Analytics and you're pitching privacy."**

> Fair, and it's why I put it in the post rather than waiting to be caught. The
> claim I'm making is about the converter, not the page, and the difference is
> real: your markdown is never in a request. I'd rather state the boundary
> exactly than say "we don't track you" and have someone open devtools.

**"Is this AI slop?" / "Did an LLM write this?"**

> Answer honestly and specifically about the engineering decisions — the
> `disableDecode` XML-injection fix, the `wp:docPr` id collision, the OMML
> validator. Nobody who wrote those from scratch has trouble talking about them,
> and that is the actual test being applied.

**"LibreOffice tests prove nothing about Word."**

> Agreed, and that's exactly what the README says. It proves the container is
> well-formed and one implementation parses it. The Word columns of the fidelity
> matrix are empty because they're unverified, not because they passed.

**"Matrices? Chemical equations? Non-Latin scripts? RTL?"**

> Matrices work — that's `m:m`, and it's part of why I couldn't use `docx`'s Math
> builders. `mhchem` isn't wired up. RTL is a document-level option that sets
> paragraph base direction. Anything MathML can express and OMML can't gets a
> warning naming the element rather than silently vanishing.

**"What happens to my `<div>` / `<br>` / raw HTML?"**

> Escaped to literal text by default, and visible. There are three honest
> options — escape, keep verbatim in monospace, or drop — and there is no fourth
> one that works, because OOXML has nowhere to put a `<div>`. I made the default
> the one where you can see what happened.

**"2 MB is a low cap."**

> It is. It's where the packing curve stops being acceptable on a mid-range
> machine, not an arbitrary limit — the numbers are in the post. Raising it means
> replacing the packer, which is the biggest open piece of work here.

**"Will you keep maintaining it?"**

> Say something you'll still be comfortable with in a year. Do not promise a
> roadmap you don't have.
