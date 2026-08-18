# downword-cli

> Markdown to Word (`.docx`) at the command line, from
> [downword](https://www.npmjs.com/package/downword).

```sh
npx downword-cli notes.md        # -> notes.docx
```

Real Word constructs, not a hand-rolled XML file: heading styles, numbering
definitions, tables, footnotes, a TOC field, page numbers and embedded images —
all restylable in Word afterwards, because the theme lives in `styles.xml`
rather than in the text.

## Install

```sh
npm install -g downword-cli     # then the command is: downword …
                                # one-shot, no install: npx downword-cli …
```

Nothing else is required. Two features have optional peer dependencies, and each
says so if you ask for it without having installed them:

| flag                           | needs                     |
| ------------------------------ | ------------------------- |
| `--math omml` / `--math image` | `npm i temml mathml2omml` |
| `--highlight`                  | `npm i highlight.js`      |

## Usage

```sh
downword [options] <input.md>...
downword [options] "docs/**/*.md" --outdir build/
cat in.md | downword > out.docx
```

- **One file** — `downword notes.md` writes `notes.docx` beside it. `-o` puts it
  somewhere else; `-o -` writes it to stdout.
- **Many files** — quote the pattern so downword expands it rather than the
  shell (`*`, `?`, `**`, `[abc]` and `{a,b}` all work) and give it `--outdir`.
  One `.docx` per input; a document that fails does not stop the others.
- **A pipe** — with no arguments and something on stdin, the document goes to
  stdout. Writing a zip into a terminal is refused rather than done.

`downword --help` lists every flag. The ones worth knowing:

```
      --theme <name>     default | github | academic | academic-double | print
      --page-size <size> A4 | Letter
      --toc              a Word table-of-contents field
      --page-numbers     a page number in every footer
      --title / --author document properties
      --math <mode>      omml | image | off              (default: off)
      --highlight        colour fenced code blocks
  -q, --quiet            /  --verbose
```

## Three defaults worth knowing

**The network is off.** A conversion cannot make a request unless you pass
`--remote-images`, and `--no-remote-images` exists so that the default has a
name of its own. When you do turn it on, each fetch is capped at 10 MiB and 10
seconds — the same caps the browser build uses — and hosts that resolve into
private, loopback or link-local space are refused.

**Images are read relative to the markdown file, and cannot escape it.**
`![](./diagram.png)` beside `docs/notes.md` means `docs/diagram.png`;
`![](../../.ssh/id_rsa)` is refused. Pass `--allow-outside-images` for documents
you wrote yourself.

**Maths is off**, because `$` is a currency symbol far more often than it opens
an equation. `--math omml` turns `$x^2$` and `$$…$$` into native, editable Word
equations.

## Mermaid

` ```mermaid ` fences are **not** drawn. mermaid measures text by laying it
out in a DOM, so rendering a diagram needs a browser; the CLI leaves the fence
as a code block and names it on stderr:

````
downword: docs/architecture.md:12: mermaid/no-dom: left the ```mermaid fence
  "The request pipeline" as a code block: drawing a diagram needs a browser
  DOM, which a CLI does not have. …
````

Pre-render the diagram and reference the picture with `![](diagram.png)`, or
convert in the browser with `downword/plugins/mermaid`.

## Diagnostics and exit codes

Warnings go to **stderr** and never change the exit code: a document with one
broken image link is still a document. By default the CLI prints the warnings
that cost the document content and counts the rest; `--verbose` prints
everything, `--quiet` prints nothing but fatal errors.

| code | meaning                                                                           |
| ---: | --------------------------------------------------------------------------------- |
|    0 | the document was written                                                          |
|   64 | bad usage: an unknown flag, a bad value, or a combination that cannot be honoured |
|   66 | an input is missing, unreadable, or a pattern matched nothing                     |
|   70 | the conversion failed                                                             |
|   73 | the output could not be written                                                   |

Source, issues and docs: <https://github.com/kspr-technologies/downword>

MIT © 2026 KSPR Technologies
