# Directory and listing submissions

Submission tracker for downword. Every row has the exact text to paste, so the
job is copy, paste, set the status.

> **Blocked until gates 1 and 2 in [`README.md`](README.md) are closed.** Almost
> every target here wants a repository URL, and several read the repo's metadata
> directly. Submitting before `github.com/kspr-technologies/downword` resolves
> gets you a rejected entry and, on the awesome lists, a closed PR you cannot
> reopen cleanly.
>
> Submit these on **Day 14 or later**, after the launch posts. Two of the lists
> below explicitly weigh whether a project looks alive.

## About the character limits below

Each entry says whether its limit is **verified** or **unverified**. Nothing is
guessed silently. Where a form's cap is not documented anywhere I could check, it
says so, and the text is sized to be comfortably short instead — with the exact
character count given, so you can trim against whatever the form actually
accepts. **Check the form before you paste**, and if a limit turns out to be
real, record it here for next time.

## Status legend

| Value                  | Meaning                                          |
| ---------------------- | ------------------------------------------------ |
| `blocked`              | Waiting on a gate in `README.md`.                |
| `todo`                 | Gates clear, not yet submitted.                  |
| `submitted YYYY-MM-DD` | Sent, awaiting moderation.                       |
| `live <url>`           | Published. Record the URL — it is the backlink.  |
| `rejected — why`       | Record the reason. It usually applies elsewhere. |
| `skipped — why`        | Deliberately not submitted.                      |

---

## The tracker

| #   | Target                                          | Submit at                                                                                               | Category / section                                | Exact text                                                                                         | Status                 |
| --- | ----------------------------------------------- | ------------------------------------------------------------------------------------------------------- | ------------------------------------------------- | -------------------------------------------------------------------------------------------------- | ---------------------- |
| 1   | AlternativeTo — alternative to **pandoc**       | <https://alternativeto.net/manage-app/> (sign in, "Add application")                                    | App listing + "Alternative to" link to Pandoc     | [§1](#1-alternativeto--the-app-listing) + [§2](#2-alternativeto--why-its-an-alternative-to-pandoc) | `blocked`              |
| 2   | AlternativeTo — alternative to **CloudConvert** | Same listing, second "Alternative to" link                                                              | Existing listing, add the link                    | [§3](#3-alternativeto--why-its-an-alternative-to-cloudconvert)                                     | `blocked`              |
| 3   | AlternativeTo — "**Word's own paste**"          | **Not submittable as a row.** See [§4](#4-alternativeto--the-word-paste-angle-read-this-before-you-try) | —                                                 | [§4](#4-alternativeto--the-word-paste-angle-read-this-before-you-try)                              | `skipped — not an app` |
| 4   | **awesome-markdown**                            | <https://github.com/BubuAnabelas/awesome-markdown> — PR                                                 | Tools → converters (verify current heading)       | [§5](#5-awesome-markdown)                                                                          | `blocked`              |
| 5   | **awesome-privacy**                             | <https://github.com/Lissy93/awesome-privacy> — PR                                                       | Document / file tooling (verify current section)  | [§6](#6-awesome-privacy--read-the-warning-first)                                                   | `blocked`              |
| 6   | **Markdown Guide — Tools**                      | <https://github.com/mattcone/markdown-guide> — PR                                                       | Tools directory entry                             | [§7](#7-markdown-guide--tools-directory)                                                           | `blocked`              |
| 7   | **Tiny Helpers**                                | <https://github.com/stefanjudis/tiny-helpers> — PR                                                      | Data entry; tag under text / content              | [§8](#8-tiny-helpers)                                                                              | `blocked`              |
| 8   | **OpenAlternative**                             | <https://openalternative.co/submit>                                                                     | Open-source alternative to CloudConvert           | [§9](#9-openalternative)                                                                           | `blocked`              |
| 9   | **awesome-cli-apps** (for the CLI)              | <https://github.com/agarrharr/awesome-cli-apps> — PR                                                    | Productivity / documents (verify current heading) | [§10](#10-awesome-cli-apps--for-downword-cli)                                                      | `blocked`              |
| 10  | npm keywords + GitHub topics                    | The repo itself                                                                                         | Not a directory — the input to several of these   | [§11](#11-npm-keywords-and-github-topics--not-a-directory-but-do-it-first)                         | `blocked`              |

Targets deliberately **not** on this list, and why: [see below](#do-not-submit-to-these).

---

## 1. AlternativeTo — the app listing

**Limit: unverified.** AlternativeTo does not document a character cap for the
app description that I could confirm, and the field renders truncated on some
views. Two lengths are given. Start with the long one; if the form or the
rendered page cuts it, fall back.

**Long — 426 characters:**

```text
downword converts Markdown into a real Word .docx entirely inside your browser — nothing is uploaded. Headings become genuine Word heading styles, so templates, the navigation pane and automatic tables of contents all work. Tables, lists and footnotes survive, and $x^2$ becomes a native, editable Word equation rather than a picture of one. Free, no signup, MIT licensed, with a command-line version and a JavaScript library.
```

**Short — 268 characters:**

```text
downword converts Markdown into a real Word .docx entirely inside your browser — nothing is uploaded. Real Word heading styles, so templates and tables of contents work; tables and footnotes survive; equations arrive native and editable. Free, no signup, MIT licensed.
```

**One-liner, if there is a tagline field — 77 characters:**

```text
Markdown to a real Word .docx, entirely in your browser. Nothing is uploaded.
```

**Other fields:**

| Field      | Value                                                                                  |
| ---------- | -------------------------------------------------------------------------------------- |
| Name       | `downword`                                                                             |
| Website    | `https://ksprtech.com/tools/markdown-to-word`                                          |
| Repository | `https://github.com/kspr-technologies/downword`                                        |
| License    | Open source — MIT                                                                      |
| Pricing    | Free                                                                                   |
| Platforms  | Web, plus Node.js/CLI (self-hosted). **Not** Windows/Mac/Linux desktop — it is a page. |
| Tags       | `markdown`, `docx`, `word`, `converter`, `privacy`, `offline`, `open-source`, `cli`    |

Do **not** tick a platform because the CLI runs there in a terminal — AlternativeTo
readers filter on platform expecting a native app, and a wrong tick reads as
padding.

## 2. AlternativeTo — why it's an alternative to pandoc

Paste into the "why is this an alternative" field on the Pandoc link.
**Limit: unverified. 324 characters.**

```text
For the case pandoc cannot serve: a browser tab, on a machine where you cannot install a binary. pandoc is more capable in almost every other direction — dozens of formats, --reference-doc templates, citations — and if you can install it, you should. downword does one conversion, markdown to .docx, with nothing to install.
```

Conceding to pandoc in the pitch is deliberate. This is a community that knows
pandoc well, and a listing claiming to beat it would be marked down instantly.

## 3. AlternativeTo — why it's an alternative to CloudConvert

**Limit: unverified. 219 characters.**

```text
The same conversion without the upload. CloudConvert processes your file on its own servers; downword's conversion runs in your browser tab and the document never leaves it. Free, unlimited, no account, and open source.
```

That is a factual statement about how a server-side converter necessarily works,
not a claim about CloudConvert's conduct. Keep it that way — do not editorialise
about what they do with files, because you do not know.

## 4. AlternativeTo — the "Word paste" angle, read this before you try

**There is no honest row here, and it is worth being explicit about why.**

AlternativeTo indexes _applications_. "Pasting into Word" is a behaviour, not an
app, so there is nothing to link to. The only listing that exists in that
direction is Microsoft Word itself, and **listing downword as an alternative to
Microsoft Word would be false** — it is a one-way exporter, not a word
processor. Submitting that gets the entry flagged and costs you credibility on a
site where the community moderates entries.

So the copy-paste angle belongs in the description, not in the graph. If you
want it visible, add this sentence to the long description in §1 (it pushes the
total to roughly 640 characters, so only do it if the field allows it):

```text
It also solves the thing everyone tries first: pasting a rendered answer into Word looks fine but stamps the web page's fonts and colours directly onto the text, so applying your own template afterwards does nothing — there is no style underneath to change.
```

**If you want a third "alternative to" link**, look for existing listings for
other markdown converters — Zamzar and Dillinger are plausible candidates. **I
have not verified that either listing exists**, so search the site first, and
only link to something you have actually opened.

## 5. awesome-markdown

Repo: <https://github.com/BubuAnabelas/awesome-markdown>

**Format: verified by convention, not by a character cap.** Awesome lists are
linted by `awesome-lint`, which requires the exact shape
`- [Name](link) - Description.` — description starts with a capital letter and
ends with a full stop. There is no documented character limit; the line below is
147 characters, which is in line with its neighbours.

```text
- [downword](https://github.com/kspr-technologies/downword) - Convert Markdown to a real Word .docx in the browser, with native editable equations.
```

**Before opening the PR:**

- Read the repo's `contributing.md` — some awesome lists require the project to
  have existed for a minimum period, or to have a certain amount of
  documentation. downword has a substantial README and a CLI, which is normally
  what is being checked for.
- Put it in the converters or tools section. **Check the current headings** —
  they get reorganised, and a PR into a section that no longer exists is a wasted
  round trip.
- One entry per PR. Keep the PR description to one sentence and a link.

## 6. awesome-privacy — read the warning first

Repo: <https://github.com/Lissy93/awesome-privacy>

> **This submission is gated on a decision, not just on the repo existing.** The
> hosted tool page loads Google Analytics and Tag Manager, like the rest of
> ksprtech.com. Submitting an analytics-loading page to a privacy list is the
> single most likely way to get publicly corrected, and being corrected on that
> list is worse than not being on it. **Two honest options:**
>
> 1. **Remove analytics from `/tools/markdown-to-word`** and then submit the
>    hosted page. Cleanest, and it makes the privacy claim unqualified.
> 2. **Submit the project, not the page** — point the entry at the GitHub repo
>    and describe the library and CLI, which genuinely make no requests at all,
>    and mention that the hosted demo is a normal website.
>
> Pick one before you open the PR. Do not submit the hosted page as-is and hope.

**Format: verified as "not raw markdown".** This repo keeps its entries as
structured data rather than as README lines, so the entry is a data record with
named fields. **Do not guess the field names from this document** — open the
repo's contributing guide and the existing data file, and match the schema you
find there exactly. The prose below is what goes in whichever field holds the
description.

**Limit: unverified. 246 characters.**

```text
downword — a Markdown to Word (.docx) converter whose conversion runs entirely client-side and makes no network requests, so documents are never uploaded. MIT licensed. Usable as a hosted page, a CLI, or a library, and works with the tab offline.
```

If option 2 above is chosen, append:

```text
The hosted demo is on a normal marketing site that loads analytics; the library and CLI make no requests at all.
```

Volunteering that is not a weakness in a privacy listing. It is the thing that
makes the rest of the entry believable.

## 7. Markdown Guide — Tools directory

Site: <https://www.markdownguide.org/tools/> · Repo:
<https://github.com/mattcone/markdown-guide>

The tools directory is generated from files in the repository, so the submission
is a PR adding one. **Limit: unverified**, and **the field names are unverified
too** — open an existing tool file and copy its structure rather than trusting
the shape below. The content is what matters:

| Field             | Value                                                          |
| ----------------- | -------------------------------------------------------------- |
| Name              | `downword`                                                     |
| Link              | `https://ksprtech.com/tools/markdown-to-word`                  |
| Platforms         | Web (plus a Node.js CLI)                                       |
| Categories / tags | Whatever the existing entries use for converters and exporters |
| Description       | The 180-character text below                                   |

```text
Converts Markdown to a Microsoft Word .docx entirely in the browser, using real Word heading styles and native editable equations. Also available as a CLI and a JavaScript library.
```

This is the highest-relevance directory on the list — its whole audience is
people looking for markdown tooling — so it is worth spending the extra ten
minutes matching the house style of the existing entries.

## 8. Tiny Helpers

Site: <https://tiny-helpers.dev> · Repo:
<https://github.com/stefanjudis/tiny-helpers>

A curated collection of single-purpose online tools for web developers.
Submission is a PR adding an entry to the site's data.

> **Fit is borderline and you should check before spending the effort.** The
> collection is aimed at web-development helpers; markdown → `.docx` is adjacent
> rather than central. **Read the scope statement in the contributing
> guidelines** and only open the PR if it clearly qualifies. A rejected PR here
> costs a maintainer's time, which is the currency you are asking for.

**Limit: unverified. 139 characters**, sized to match the terse house style of
the existing entries.

```text
Paste Markdown, get a real .docx. Runs entirely in the browser — real Word heading styles, tables, footnotes and native editable equations.
```

Suggested tags: `markdown`, `converter`, `documents`, `privacy`. Match whatever
vocabulary the data file already uses rather than inventing tags.

## 9. OpenAlternative

Submit at: <https://openalternative.co/submit>

A directory of open-source alternatives to proprietary software. The proprietary
counterpart to name is **CloudConvert** (a hosted converter you upload to), not
Microsoft Word.

> **This one reads your repository.** It pulls metadata — description, stars,
> license, language — from the GitHub API, so **your repo description is your
> listing copy.** Close gate 6 in [`README.md`](README.md) before submitting:
> set the repo description and topics deliberately, then submit. Verify what it
> actually rendered afterwards, and correct the repo rather than the listing if
> it is wrong.

**Repo description to set — 110 characters** (GitHub's own limit is 350, so this
is comfortably inside it):

```text
Markdown to Word (.docx), entirely in your browser. Real heading styles, native editable equations, no upload.
```

**Other fields:** website `https://ksprtech.com/tools/markdown-to-word`, repo
`https://github.com/kspr-technologies/downword`, alternative to CloudConvert.
Any free-text field can take the 268-character short description from §1.

## 10. awesome-cli-apps — for `@ksprtech/downword-cli`

Repo: <https://github.com/agarrharr/awesome-cli-apps>

This one is for the **command-line package**, not the web tool, and the entry
should read that way — a reader of this list has a terminal and does not care
that there is a hosted page.

**Format: same `awesome-lint` shape as §5.** 119 characters.

```text
- [downword](https://github.com/kspr-technologies/downword) - Convert Markdown files to Microsoft Word .docx documents.
```

**Check the current headings** before opening the PR — the productivity and
documents sections have been reorganised before. And be honest in the PR
description that this is your own project; these maintainers ask.

## 11. npm keywords and GitHub topics — not a directory, but do it first

Not a submission, but it is the input to §9 and to npm's and GitHub's own
search, which between them will out-refer several of the lists above.

`packages/core/package.json` already carries:

```
markdown, docx, word, converter, office, ooxml, browser, chatgpt, claude
```

**GitHub topics to set on the repo** (gate 6):

```
markdown  docx  word  ooxml  converter  browser  client-side  privacy
typescript  markdown-to-word  office-open-xml  no-upload
```

Keep the two lists aligned, and do not stuff either with terms the project does
not deliver on.

---

## Do not submit to these

Recording the reasoning so nobody re-litigates it in three months.

- **awesome-selfhosted.** Its guidelines are for network-facing software you run
  on your own server. downword is a client-side library and a static page; there
  is nothing to host. A PR would be closed, correctly.
- **"Top 50 free online tools" SEO farms.** The listing pages that appear when
  you search for this category are mostly link farms. A link from one is worth
  approximately nothing and being adjacent to them is a mild negative signal.
- **Anything that charges for placement**, unless you have decided separately
  that it is an advertising spend. It is not a listing at that point and it does
  not belong in this tracker.
- **Product Hunt** is not on the list because it is a launch, not a listing: it
  needs its own assets, its own day, and someone at the keyboard all of it. If
  you want to do it, plan it separately — do not tack it onto a directory
  session.

## Before you submit anything

1. Gates 1 and 2 in [`README.md`](README.md) are closed.
2. The repo description and topics are set (§11), because §9 reads them.
3. The decision in §6 about analytics has been made.
4. Every claim in the text you are pasting still matches the
   [honesty invariants](README.md#honesty-invariants). The descriptions above
   say the conversion runs in the browser and nothing is uploaded — both true —
   and none of them says the page makes no network requests, which is false.
5. Update the Status column the moment you submit. An untracked submission gets
   submitted twice.
