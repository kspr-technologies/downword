# Launch kit

Drafts for a human to post. **Nothing in this directory has been posted, and no
agent may post it.** Every file is copy waiting on a person to read it, disagree
with half of it, and press the button themselves.

---

## DO NOT POST UNTIL

Every one of these is a human gate. The drafts are written to be true _after_
they are closed, and several of them are false today. Posting early is not
"launching early" — it is linking a Hacker News front page at a 404.

| #   | Gate                                                                                                                                                                                                                                                                                                                      | True today?                  |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------- |
| 1   | **`github.com/kspr-technologies/downword` exists and is public.** Every draft links to it, `package.json` already points at it, and the README's CI badge already resolves against it.                                                                                                                                    | **No — it 404s.**            |
| 2   | **`downword` and `downword-cli` are published to npm, and every quoted command has been run from a clean cache.** The drafts contain `npm install downword` and `npx downword-cli notes.md`. See the note below — the root README quotes a different invocation and one of the two is wrong.                              | **No.**                      |
| 3   | **A human has opened the output in real Microsoft Word** and filled in at least one column of [`docs/fidelity-matrix.md`](../fidelity-matrix.md), with a row in its verification log. Follow the protocol in that file — do not spot-check and call it done.                                                              | **No — zero Word coverage.** |
| 4   | **The demo GIF is recorded.** See the `TODO(hero)` block at the top of the root [`README.md`](../../README.md); it has the shot list and the ffmpeg command. HN tolerates no media; Reddit and LinkedIn convert far worse without it.                                                                                     | No.                          |
| 5   | **Third-party facts re-verified.** These drafts describe how Notion exports markdown, what the Obsidian Pandoc plugin is called, and how four directories accept submissions. All of that was written on 2026-08-18 and all of it can change. Re-check before you paste; see the per-file "verify before posting" notes.  | Written, not re-checked.     |
| 6   | **The GitHub repo description and topics are set deliberately.** At least one directory in [`listings.md`](listings.md) scrapes repo metadata and renders it as your listing. Whatever is in the description field becomes the pitch.                                                                                     | Blocked on gate 1.           |
| 7   | **You have decided what to do about analytics.** Every draft says plainly that the hosted page loads Google Analytics and that the download button fires one event. That sentence is load-bearing for the privacy claim's credibility. If you would rather not say it, remove the analytics — do not remove the sentence. | Decide before Day 0.         |

### Gate 2 has a trap in it

The `downword` **bin** is declared by the **`downword-cli`** package
(`packages/cli/package.json` → `"bin": { "downword": … }`). The `downword`
package — the library — declares no bin at all. So the root README's
`npx downword notes.md` resolves the _library_ package, which has nothing to
run. The drafts in this directory say `npx downword-cli notes.md` instead, which
should work because npx runs a package's sole binary even when the names differ.

**Neither form has been verified against a published package, because there is
no published package.** Before you post anything containing a command: publish,
then run both forms in a container with an empty npm cache, then fix whichever
of the two documents is wrong. A launch post whose first code block does not run
is the most expensive kind of typo.

### After gate 3, come back and edit

The drafts currently say the output has **never been opened in Microsoft Word**,
because that is true. If you do the fidelity pass, that sentence becomes false
and every file that carries it needs updating — to what you actually found,
including the parts that were wrong. Grep for `Word` before you post.

---

## The staggering rule

From `SEO-GROWTH-PLAN.md`, Part 2.2 and 2.3. The short version: one channel per
day, Hacker News first and alone, everything else spread across two weeks.

**Day 0 — Show HN, and nothing else.**

- Post **Tuesday or Wednesday, 06:00–08:00 EST**. Not Monday (weekend backlog),
  not Friday (nobody is there), not a US public holiday.
- **Be at the keyboard for the first four hours.** This is the whole rule. A
  Show HN that answers every comment inside ten minutes for four hours behaves
  completely differently from one posted and abandoned. Block the morning out.
- Post the first comment yourself, immediately, from the submitting account.
- **Do not post anything to Reddit the same day.** Not one subreddit. A Show HN
  and four Reddit posts in one day reads as a campaign, and both communities
  can see your posting history.
- Do not ask anyone to upvote. It is the one thing HN reliably detects and
  penalises, and it would sink the whole launch.

**Days 1–14 — one channel per day, at most.**

| Day | Channel                                  | File                                               | Why here                                                                                  |
| --- | ---------------------------------------- | -------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| 0   | Show HN                                  | [`show-hn.md`](show-hn.md)                         | Hardest audience, best backlink, and it must not share a day with anything.               |
| 1   | _Nothing._ Keep answering the HN thread. | —                                                  | The thread is still live on day 1. Feeding a second channel splits your attention.        |
| 2   | LinkedIn                                 | [`linkedin.md`](linkedin.md)                       | Your warmest audience, and the only channel where a launch post is expected.              |
| 4\* | r/webdev                                 | [`reddit/r-webdev.md`](reddit/r-webdev.md)         | Closest to the HN crowd, so go while the repo is still warm — but not on the same day.    |
| 7   | dev.to article                           | [`devto.md`](devto.md)                             | Needs the repo live and a week of settled facts. Also the piece most worth getting right. |
| 9   | r/ObsidianMD                             | [`reddit/r-obsidianmd.md`](reddit/r-obsidianmd.md) | Different audience, different pitch, far enough from r/webdev to not look like a blast.   |
| 11  | r/Notion                                 | [`reddit/r-notion.md`](reddit/r-notion.md)         | Same reasoning. Notion and Obsidian users overlap; two days apart is the minimum.         |
| 13  | r/ChatGPT                                | [`reddit/r-chatgpt.md`](reddit/r-chatgpt.md)       | Biggest and most self-promo-hostile sub. Go last, when you have replies to point at.      |
| 14+ | Directory submissions                    | [`listings.md`](listings.md)                       | Several directories want to see the project is real. Submit after the noise, not before.  |

**\* r/webdev is Saturday-only.** The sub restricts sharing your own project to
**Showoff Saturday**, with the matching flair. Day 4 works if Day 0 is a
Tuesday; if Day 0 is a Wednesday, r/webdev slides to the following Saturday and
r/ObsidianMD moves up to fill Day 4. Check the rule the morning you post — it is
enforced, and a mid-week project post gets removed.

**Rules that apply to every Reddit day:**

- Read the subreddit's current rules and its self-promotion policy the morning
  you post. Several of these subs require account age, comment karma, or a flair,
  and the thresholds change.
- One subreddit per day. Never cross-post the same text — there is no duplicated
  copy between the four drafts on purpose, and reusing one in the wrong sub is
  how a launch turns into a ban.
- Reply as the author, disclose that you built it in the post body itself, and
  answer the "is this an ad" comment plainly rather than defensively.
- If a sub's rules say self-promotion goes in a weekly thread, use the weekly
  thread. It converts worse and costs nothing.

---

## What is in here

| File                                               | What it is                                                                                         |
| -------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| [`show-hn.md`](show-hn.md)                         | The Show HN title, the first comment, and prepared answers for the questions HN will actually ask. |
| [`reddit/r-chatgpt.md`](reddit/r-chatgpt.md)       | r/ChatGPT. Non-technical, about the copy button.                                                   |
| [`reddit/r-obsidianmd.md`](reddit/r-obsidianmd.md) | r/ObsidianMD. Vault-shaped markdown, and what does not survive it.                                 |
| [`reddit/r-notion.md`](reddit/r-notion.md)         | r/Notion. Notion has no Word export; this is the workaround and its caveats.                       |
| [`reddit/r-webdev.md`](reddit/r-webdev.md)         | r/webdev. The build, the CI fidelity gate, and a request for Word testers.                         |
| [`devto.md`](devto.md)                             | Technical article: TeX → MathML → OMML, and why `docx`'s Math builders could not do the job.       |
| [`linkedin.md`](linkedin.md)                       | Short post, with the campaign UTM on the link.                                                     |
| [`listings.md`](listings.md)                       | Directory submission tracker, with the exact text to paste into each one.                          |

---

## Honesty invariants

These are the claims that took work to get right. Keep them exactly this
precise when you edit — every one of them has a looser version that is false,
and the looser version is what a reader will hold you to.

1. **"The converter makes no network requests."** True: the conversion path
   cannot reach the network unless remote images are switched on. **"The page
   makes no network requests" is false** — the hosted page loads Google
   Analytics and Tag Manager like the rest of the site. Say which one you mean,
   every time.

2. **What the download event contains.** Clicking Download fires one analytics
   beacon carrying the theme name, three booleans, two byte counts, a duration
   and a count of notices. Numbers and closed enums only. It does not contain
   your text, your filename, or any part of the document. Say that concretely
   rather than saying "anonymous".

3. **`npm install downword` does not work yet** (gate 2), and
   `github.com/kspr-technologies/downword` **404s** (gate 1). No draft may imply
   otherwise before those gates close.

4. **Equations really are native.** `$…$` becomes `m:oMath` — editable in
   Word's equation editor, not a picture. This one you can state flatly.
   The n-ary caveat (∫ and ∑ draw an empty placeholder box next to a complete,
   editable equation) is a real defect and every technical draft volunteers it.

5. **Code blocks in the hosted tool have no syntax colours.** The library
   supports highlighting through an opt-in `highlight.js` adapter; the website
   passes no highlighter, so fenced code arrives as `CodeBlock`-styled monospace
   in one colour. Do not describe the tool as if it inherited the library's
   feature list.

6. **The hosted tool emits no TOC field.** Headings are real `Heading1`–`Heading6`
   styles, so Word's References → Table of Contents builds one in a click — which
   is the thing worth saying. The CLI has `--toc`; the web tool does not expose it.

7. **Raw HTML is escaped to literal text by default, not dropped.** `<aside>`,
   `<br>`, `<div align="center">` all arrive as visible text. That is a
   deliberate choice among three bad options and it surprises people, so say it
   before they find it.

8. **Remote images are refused by default** and degrade to a visible
   `[image: alt]` placeholder plus one error notice. Turning them on is a
   cross-origin `fetch` that plenty of hosts refuse; the document is never lost.

9. **Mermaid needs a browser.** It renders in the web tool. In Node the fence
   stays a code block and raises a notice.

10. **Never opened in Microsoft Word.** CI opens every generated file in
    headless LibreOffice and converts it to PDF, with a positive control on the
    import filter and a negative control on a corrupt file. That proves the
    container parses. It does not prove a heading looks like a heading. Until
    gate 3 closes, say exactly that.

11. **Be fair to pandoc.** Every draft that mentions it says some version of: if
    you can install pandoc, install pandoc. It is twenty years old, it is better
    at this in almost every dimension that is not "runs in a browser tab", and
    `--reference-doc` — restyling from your own corporate template — is a thing
    downword cannot do at all. Punching at pandoc reads as a tell; conceding to
    it reads as confidence.

12. **No invented numbers, ever.** Every figure in these drafts is measured, and
    each one carries the machine and the corpus it was measured on. There are no
    star counts, no download counts, no user counts, no testimonials, and no
    benchmark against a named competitor that was not actually run. If you want
    a number that is not here, measure it first.

---

## Links, and tracking parameters on them

- **Hacker News and Reddit: use the bare URL.** `https://ksprtech.com/tools/markdown-to-word`.
  Both audiences notice `utm_` parameters, and a privacy pitch arriving with
  campaign tracking attached is an own goal for a rounding error of attribution.
  You will see the referrer anyway.
- **LinkedIn: use the UTM.** It is in [`linkedin.md`](linkedin.md) and it is the
  one channel where the plan explicitly asks for measurement.
- **dev.to: canonical URL back to the tool page**, plus the article's own UTM if
  you want it. Normal there, and expected.
