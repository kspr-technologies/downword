# r/Notion

> **Draft. Do not post until every gate in [`../README.md`](../README.md) is
> closed.** Post on **Day 11** — two days after r/ObsidianMD, because the
> audiences overlap and back-to-back posts read as a blast.

> **VERIFY BEFORE POSTING.** This draft makes specific claims about how Notion's
> markdown export is shaped: that there is no `.docx` option, that databases come
> out as `.csv` rather than markdown tables, that images land in a folder next to
> the `.md` as relative paths, and that callouts export as `<aside>` HTML. All of
> that was written on 2026-08-18 and Notion changes its exporter. **Do one real
> export of a page containing a callout, a database view and an image, open the
> zip, and correct anything below that has moved.** Getting a Notion fact wrong
> in r/Notion is the fastest way to lose the thread.

**Voice here:** workspace people, not developers. They want steps they can
follow, in order, and they want to know what will break before they spend twenty
minutes on it. Do not talk about parsers.

---

## Title

```
Notion still has no Word export. Here's the workaround I built — plus the three things that don't survive the trip.
```

Alternative:

```
I needed a Notion page as a real .docx for a client, found there's no export for it, and built the missing step
```

**Flair:** "Tools" / "Resource" / whatever the sub currently uses. Not "Question".

---

## Post body

---

Notion exports to PDF, HTML, and Markdown & CSV. There is no Word export, and
for a chunk of the world "send me the doc" still means a `.docx` that someone
else is going to open, track changes in, and send back. PDF does not do that
job.

I hit this with a client deliverable, could not find a way to do it that did not
involve uploading the page to a converter site, and ended up building the
missing step. Sharing it plus, more usefully, the exact workflow and its
potholes.

### The workflow

1. On the Notion page, **⋯ → Export → Markdown & CSV**. You get a download —
   usually a `.zip`, sometimes a single `.md` for a very simple page.
2. Unzip it, open the `.md` in any text editor, select all, copy.
3. Paste it into https://ksprtech.com/tools/markdown-to-word and click Download.

That is it. Free, no signup, and the conversion happens inside your browser tab
— the page contents are not uploaded anywhere, which is the bit I cared about,
because the document that started all this was under NDA. (Being precise: the
website itself runs Google Analytics like any normal site, so clicking Download
sends one event with the theme, a couple of settings and a duration. Your text
is not in it and never leaves the tab.)

If the Word file you need to end up in is already open, skip step 3's download
entirely: **Copy for Word / Docs** puts the formatted content on your clipboard
and you paste it in where you want it. Works into Google Docs and Outlook too.
(It needs the clipboard API, so on a few Firefox builds the button is disabled
and tells you why instead of failing silently.)

### What you get on the other side

The thing that matters, and the reason I did not just use one of the free
converter sites: **headings come out as real Word headings.** Not "text that has
been made big and bold" — actual `Heading 1` / `Heading 2` styles. Which means
your company template applies to it, the navigation pane works, and
References → Table of Contents builds a real TOC in one click instead of you
typing one.

Also survives: bullet and numbered lists (properly nested, and separate lists
restart at 1 instead of continuing from the previous one), tables with their
column alignment, bold/italic/strikethrough, inline code, quotes, links, and
footnotes.

### The three things that do not survive, in order of how much they will annoy you

**1. Images.** This is the big one. Notion's export puts your images in a folder
next to the `.md` and links them by relative path. A web page in your browser
cannot read files off your disk — that is a hard browser rule, not a missing
feature — so those links cannot resolve. You get a visible `[image: alt]`
placeholder in the Word file, in the right position, and a warning telling you
which ones. **Practical answer: convert first, then drag the images from the
unzipped folder into Word afterwards.** The placeholders tell you exactly where
they go. For a page with three screenshots that is fine; for a page with thirty
it is not, and I would rather say so than let you find out at image nineteen.

**2. Databases.** An inline database or a database view exports as a separate
`.csv` file, not as a markdown table, so it is not in the `.md` you are pasting
and it will not appear in the Word file. If you need the table, open the `.csv`
in Excel, copy the range, paste it into Word where you want it. Only plain
Notion tables (the simple block, not a database) come through as markdown.

**3. Callouts, and anything Notion exports as HTML.** In the exports I have
looked at, callouts come out as an `<aside>` HTML block. The converter turns raw
HTML into visible literal text rather than deleting it — so you will see the tag
in your document and you can delete the two lines. Deliberate: Word has nowhere
to put an `<aside>`, and the three options are "show it", "keep it verbatim" or
"silently bin your content". I picked the one where nothing disappears without
telling you.

Anything else the converter could not represent shows up in a panel afterwards
with a line number, rather than vanishing.

### Other caveats

- **2 MB of text maximum.** That is an enormous page; a whole workspace export
  pasted together is not.
- **Code blocks come out in the right font but with no colours.**
- **Nobody has put this in front of real Microsoft Word yet** — I do not own a
  copy. There is an automated check that opens each build in LibreOffice, so I
  can say the file is valid and opens cleanly. Whether every heading and table
  looks the way it should once Word gets hold of it is genuinely an open
  question, and I would rather tell you that than let you find out during a
  client review. If you do try it, tell me what you saw — good or bad, I want
  both.
- Free and MIT licensed, and all the code is public:
  https://github.com/kspr-technologies/downword

I built it, so ask me anything — including "why not just use X", which is a fair
question and I will give you a straight answer.

---

## Predictable questions here

**"Why not export to PDF and convert the PDF?"** — Because PDF-to-Word
conversion reconstructs the structure by guessing at it, and the whole point
here is that the structure is already known — the markdown says which lines are
headings. Converting through PDF throws that away and then tries to infer it
back.

**"Why not export HTML and open that in Word?"** — Word will open it, and it
mostly looks fine, but the formatting gets painted directly onto the text rather
than coming from styles, so the company template does not apply afterwards. Same
underlying problem as pasting.

**"Does it work for a whole workspace export?"** — Not in one go: it takes one
document at a time and caps at 2 MB. A workspace export is many `.md` files, and
if you have a terminal there is a command-line version that takes a glob. If you
do not, this is a page-at-a-time tool and I would say so up front.

**"Is there a Notion integration?"** — No, and there is deliberately no OAuth
button. Connecting it to your workspace would mean it reads your pages, which is
the thing it currently cannot do.

## What NOT to say in this sub

- Do not post it as a template or in a template megathread. It is not a template
  and it will be removed.
- Do not overstate the workflow. Three steps involving an unzip is not seamless,
  and pretending otherwise gets called out in the first comment.
- Do not mention pandoc unprompted — wrong audience. If someone raises it, be
  fair: it is better at this if you can install it.
- No developer detail here. Save it for [`r-webdev.md`](r-webdev.md).
