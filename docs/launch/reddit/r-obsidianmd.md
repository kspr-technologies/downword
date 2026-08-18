# r/ObsidianMD

> **Draft. Do not post until every gate in [`../README.md`](../README.md) is
> closed.** Post on **Day 9**.

**Before you post, check:** r/ObsidianMD allows tool and plugin sharing but
expects disclosure and dislikes anything that smells like a funnel. Read the
current rules. Also re-verify the Obsidian-side facts below — plugin names and
core-plugin behaviour change between releases, and this was written on
2026-08-18.

**Voice here:** this sub knows markdown better than you do, runs local-first on
purpose, and will immediately ask about wikilinks, callouts, Dataview and
frontmatter. So lead with the honest compatibility table instead of waiting to
be asked. Every row below was checked against the actual parser, not guessed.

---

## Title

```
Vault note to .docx without installing pandoc — and an honest list of what does and doesn't survive
```

Alternative, if the sub is in a mood about tool posts:

```
I mapped exactly which Obsidian-flavoured Markdown survives a conversion to Word, and built the converter around it
```

**Flair:** "Resource", "Showcase" or whatever the sub currently uses for shared
tools.

---

## Post body

---

The reason this exists: I needed to hand a long note to someone in Word, on a
machine where I could not install anything. The Pandoc plugin is the right
answer to that problem when you can install pandoc — it is more capable than
what I built in essentially every direction, and if you have a terminal and
admin rights, stop reading and go use it. I did not, so I wrote a converter that
runs in a browser tab.

https://ksprtech.com/tools/markdown-to-word — paste, download, no upload, no
signup. The conversion happens in your tab; the note is never sent anywhere.

What I think is actually useful to this sub is not the tool, it is the
compatibility list. Obsidian-flavoured markdown is not CommonMark, and every
converter handles the difference differently and quietly. So here is exactly
what happens, verified against the parser rather than assumed.

### Survives, and survives properly

- **Headings** become real Word `Heading 1`–`Heading 6` styles, not text that
  has been made big and bold. This is the whole reason I bothered. It means the
  navigation pane populates, and References → Table of Contents builds a working
  TOC in one click.
- **Tables**, including column alignment, as real Word tables with a header row
  that repeats across page breaks.
- **Footnotes** (`[^1]`) become real Word footnotes in `footnotes.xml` — Word
  numbers them, positions them at the foot of the page, and renumbers when you
  delete one.
- **Math.** `$x^2$` and `$$…$$` become native Word equations you can click into
  and edit. Not images. This one took most of the work and it is the feature I
  would actually defend.
- **Mermaid** fences render to an image with the info string as a caption.
- **Task lists** `- [x]`, nesting to all nine levels OOXML allows, and sibling numbered
  lists that each restart at 1 instead of continuing each other.
- **Block-quotes, bold, italic, strikethrough, inline code, internal links to
  headings** — `[text](#some-heading)` becomes a real intra-document jump.

### Does not survive, and here is precisely what you get instead

I checked each of these rather than assuming, because "unsupported" covers a lot
of different outcomes and some of them are much worse than others.

- **YAML frontmatter.** The parser is CommonMark; it does not know that
  `---` at the top of a file means frontmatter. The opening `---` becomes a
  horizontal rule, and your property lines get swallowed into a Heading 2 by the
  closing `---` (that is setext heading syntax doing exactly what it is
  specified to do). **Delete the frontmatter block before converting.** This is
  the one that will actually bite you.
- **`[[Wikilinks]]`** come out as the literal text `[[Wikilinks]]`. Not a broken
  link, not a dropped link — visible literal brackets. Same for `![[embeds]]`.
- **Callouts.** `> [!note] Title` becomes an ordinary block-quote whose first
  line reads `[!note] Title`. The content is all there; the callout chrome is
  not, and the marker is visible.
- **Dataview / Templater / any code-block-driven plugin.** A ` ```dataview `
  fence is a code fence, so you get the query text in monospace. Which is
  correct behaviour for a converter that never ran Obsidian, but worth knowing
  before you convert a dashboard note.
- **Raw HTML** is escaped to visible literal text by default rather than
  dropped. There are three honest options — escape it, keep it verbatim in
  monospace, or delete it — and no fourth one, because Word has nowhere to put a
  `<div>`. I picked the one where you can see what happened.
- **Local image links** (`![[img.png]]`, or a relative path) cannot be read by a
  page in your browser — it has no access to your vault folder. You get a
  visible `[image: alt]` placeholder and a warning, not a silently missing
  picture. Remote images are off by default and can be switched on, though a lot
  of hosts refuse cross-origin requests anyway.
- **Code blocks arrive in one colour.** Right font, no syntax highlighting.

Anything the converter could not represent shows up in a diagnostics panel
afterwards with a line number, rather than being silently dropped. That was the
design goal: never lose text, always say what happened.

### Caveats worth saying out loud

- Input is capped at 2 MB of markdown. That is a very long note, but if you were
  planning to convert a whole vault export in one go, you cannot.
- **It has never been opened in Microsoft Word.** I do not own a copy. Every
  build is opened by headless LibreOffice in CI and converted to PDF, with a
  deliberately corrupt file as a negative control so the check cannot pass
  vacuously. So I know the file is structurally sound and one real word
  processor reads it cover to cover — and I know nothing at all about how it
  looks in Word, which is a much smaller claim than people assume when they see
  a green CI badge. If someone here has Word and ten minutes, I would love to
  hear what actually happens.
- MIT licensed, and the whole thing runs locally if a website is the part you do
  not like: https://github.com/kspr-technologies/downword. The CLI
  (`npx @ksprtech/downword-cli note.md`) does everything the web version does except
  mermaid, since mermaid needs a real browser to measure text.

I am the author. Ask me anything, including which of the "does not survive" rows
you think I got wrong — a couple of them are choices rather than limitations and
I would happily be argued out of them.

---

## Predictable questions here

**"Why not the Pandoc plugin?"** — Because it needs pandoc installed. Concede
immediately and completely: if you can install it, it is better. Do not litigate
this.

**"Will you support wikilinks / callouts?"** — Honest answer: both are tractable
as a pre-processing pass (wikilinks to plain text or to an internal bookmark,
callouts to a styled block-quote). Neither exists today. Do not promise a date.

**"Does it read my vault?"** — It is a web page. It has no filesystem access
whatsoever, which is exactly why local images do not resolve.

**"What about Obsidian's own Export to PDF?"** — Different output, and fine for
PDF. This is for the person who specifically needs `.docx` because someone
downstream is going to edit it.

## What NOT to say in this sub

- Do not pitch this as a replacement for anything in their workflow. It is a
  one-way export for one specific situation.
- Do not skip the frontmatter warning to make the list look shorter. It is the
  first thing that will happen to the first person who tries it.
- No marketing language about privacy. State the mechanism (it runs in the tab,
  there is no upload path) and let them verify it in devtools. This sub will.
