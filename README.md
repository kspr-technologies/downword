# downword

> Markdown to Word (.docx) converter that runs entirely in your browser.
> Paste from ChatGPT/Claude, get a real Word document.

<!-- TODO(README): this is a developer-facing README. The product page —
     positioning, screenshots, the browser demo link, comparison table — is a
     later task. The API section below is real and is type-checked by
     packages/core/tests/readme.test.ts. -->

> [!WARNING]
> **Status: pre-release.** The library and the CLI are real and green — parser,
> renderer, images, highlighting, the public API and the `downword` binary all
> land here — but nothing is published to npm yet. Math (`downword/plugins/math`) produces native,
> editable Word equations and needs two optional peer dependencies (`temml`,
> `mathml2omml`). Mermaid diagrams (`downword/plugins/mermaid`) are real but
> need a DOM: in Node they degrade to the fenced source plus a warning rather
> than rendering.

## Packages

| Package                          | npm            | Description                     |
| -------------------------------- | -------------- | ------------------------------- |
| [`packages/core`](packages/core) | `downword`     | The library. Browser and Node.  |
| [`packages/cli`](packages/cli)   | `downword-cli` | `downword` command line binary. |

## Install

```sh
npm install downword
```

## Usage

```ts
import { convert } from "downword";

const bytes = await convert("# Hello\n\nFrom **downword**.");
// bytes is a Uint8Array containing a real .docx
```

In the browser:

```ts
import { convertToBlob } from "downword";

const blob = await convertToBlob("# Hello");
const url = URL.createObjectURL(blob);
```

With syntax highlighting (highlight.js is an optional peer dependency, loaded
lazily and only when a fence needs it):

````ts
import { convert } from "downword";
import { createHighlighter } from "downword/highlight";

const bytes = await convert("```ts\nconst x = 1;\n```", {
  highlighter: createHighlighter(),
});
````

Full API, options and behaviour: **[packages/core/README.md](packages/core/README.md)**.

At the command line:

```sh
npx downword notes.md                        # -> notes.docx
npx downword "docs/**/*.md" --outdir build/  # one .docx per file
cat notes.md | npx downword > notes.docx     # stdin -> stdout
```

Flags, defaults and exit codes:
**[packages/cli/README.md](packages/cli/README.md)**.

## Why

Every "markdown to Word" tool either uploads your document to somebody's server
or hands you an HTML file with a `.doc` extension that Word complains about.
downword builds a real OOXML package in your browser; nothing leaves the tab.

Concretely, that means:

- **No network by default.** The default image resolver reads `data:` and
  `blob:` URLs only. Fetching over HTTP is one explicit opt-in
  (`allowRemoteImages`), capped at 10 MiB and 10 s per image, and a test with a
  `fetch` spy proves the default cannot reach the network at all.
- **A real `styles.xml`.** Headings, quotes and code are styled by _style_, not
  by inline formatting, so the document restyles like one Word wrote.
- **Real list numbering.** Every list that counts gets its own numbering
  instance, so consecutive lists restart instead of continuing each other —
  including the nested ones, which is the bug every other converter ships with.
- **Deterministic output.** The same markdown produces the same bytes, apart
  from the creation timestamp and one relationship id per external link, which
  `docx` mints itself. That is what makes the golden tests in
  `packages/core/tests` possible.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md). Short version:

```sh
pnpm install
pnpm lint && pnpm typecheck && pnpm test && pnpm build && pnpm size
```

## License

[MIT](LICENSE) © 2026 KSPR Technologies
