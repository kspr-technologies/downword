# r/ChatGPT

> **Draft. Do not post until every gate in [`../README.md`](../README.md) is
> closed.** Post on **Day 13** — last of the four subreddits, so you have
> answered comments elsewhere first and are not a brand-new account dropping a
> link.

**Before you post, check:** r/ChatGPT is one of the largest and most
self-promotion-hostile subs in this set, and its rules change. Read the current
rules and the sidebar the morning you post. If self-promotion is restricted to a
weekly or pinned thread, **use that thread** — it converts worse and costs you
nothing. If it is banned outright, r/ChatGPTPro and r/OpenAI are the fallbacks,
and this draft works in either with the title swapped.

**Voice here:** non-technical. Nobody in this sub cares about OOXML. Everybody
in this sub has hit the copy button and then had to fix the formatting by hand.
Lead with the annoyance, give away something useful whether or not they click,
and be visibly honest about the limits.

---

## Title

Pick one. Both are under Reddit's 300-character limit.

```
The Copy button gives you Markdown. The person who asked for the report wants Word. I built the missing step.
```

```
I got tired of pasting ChatGPT answers into Word and getting literal ## headings, so I made a converter that fixes it
```

**Flair:** whatever the sub currently uses for "Resources", "Tools" or
"Project". Do not post unflaired in a sub this size.

---

## Post body

---

Ask for a report, get a great answer, need to hand it to someone who lives in
Word. That last step is where twenty seconds of annoyance lives, every single
time.

Here is why it is annoying, because it took me a while to work out that there
are two different failure modes:

**If you use the Copy button**, you get the raw markdown. Word has no idea what
that is, so you get literal `## Heading` and literal `| pipes |` all down the
page, and you reformat forty headings by hand.

**If you select the answer on screen and copy that**, you get something that
looks much better — Word actually does understand headings from a web page, and
tables usually survive. But it also glues the website's exact fonts, sizes and
colours onto every single word. So when your boss says "use the company
template", nothing happens, because there is no style to change underneath — the
formatting is painted directly onto the text. Code blocks also lose their
monospace font and turn into ordinary sentences.

### The free tip, which works with no tool at all

If you only ever take one thing from this post: **paste the answer, then in Word
go to the Home tab, hit Clear All Formatting on the pasted block, and reapply
Heading 1 / Heading 2 from the Styles gallery yourself.** It is tedious, but
after that the document is a real Word document and templates, the navigation
pane and automatic tables of contents all start working. That is the actual
difference between "looks like a heading" and "is a heading", and it is why the
template thing never works otherwise.

### The tool

I got tired of doing that by hand, so I built it: paste the markdown, click
Download, get a `.docx` where every heading is a real Word heading.

https://ksprtech.com/tools/markdown-to-word

Free, no signup, no upload. The conversion runs inside your browser tab — your
text is not sent anywhere, which mattered to me because half of what I convert
is work stuff I should not be posting to a random converter site. (Full
disclosure so nobody catches me out on it: the _page_ is on a normal website
that loads Google Analytics, so a click on Download sends one event with the
theme name, a few yes/no settings and how long it took. Your text is not in it
and cannot be.)

Things that survive, which are the ones I actually needed: headings, tables,
bullet and numbered lists, bold and italic, footnotes, and equations. Equations
were the one I cared most about — `$x^2$` becomes a real Word equation you can
click and edit, not a picture of an equation. If you have ever had ChatGPT
explain some maths and then watched it turn into gibberish in Word, that is the
fix.

There is also a "Copy for Word / Docs" button if you just want to paste straight
into a document you already have open, without downloading a file. (It needs the
clipboard API, so on a few Firefox builds it is disabled and tells you why.)

### What it does not do, so you are not disappointed

- **Images from the web will not come through by default.** There is a toggle,
  but even with it on, a lot of image hosts refuse the request for reasons
  outside my control. You get a visible placeholder rather than a silently
  broken document.
- **Code blocks come out in one colour.** Correct font, no syntax colours.
- **Documents up to 2 MB of text.** Past that it gets slow enough that I would
  rather tell you than let it hang. 2 MB is a very long report.
- **I have not been able to test it in Microsoft Word itself yet** — I do not
  have a copy. It is tested automatically against LibreOffice on every change,
  which proves the file is valid and opens, but if you try it in real Word I
  would genuinely like to know what you see.

Built it for myself, it is free, MIT licensed, and there is nothing to sign up
for. Happy to answer anything.

---

## First comment (post it yourself, straight away)

> Built this myself and I am the developer, so ask me anything including the
> awkward questions. If you want to check the "nothing gets uploaded" claim
> rather than take my word for it: open devtools, go to the Network tab, and
> convert something. It should stay empty apart from the page itself.

## What NOT to say in this sub

- No OOXML, no `styles.xml`, no "Web Worker", no bundle sizes. Different post,
  different sub — that material is in [`r-webdev.md`](r-webdev.md).
- Do not mention pandoc here. This audience does not have a terminal open, and
  "you could also install this command-line tool" reads as noise. (Everywhere
  that pandoc _is_ relevant, be fair to it — see the invariants in
  [`../README.md`](../README.md).)
- Do not claim it is better than anything. This sub reacts badly to comparison,
  and the post does not need it.
- Do not post the GitHub link in the body. If someone asks, give it in a reply.
  In this sub the repo link makes it read like a launch; in r/webdev it is the
  point.
