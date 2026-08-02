# downword

> Markdown to Word (.docx) converter that runs entirely in your browser.
> Paste from ChatGPT/Claude, get a real Word document.

<!-- TODO(README): this is a placeholder. The real product README — positioning,
     screenshots, the browser demo link, feature matrix, comparison table, and
     usage docs — is a later task. Do not treat anything below as final copy. -->

> [!WARNING]
> **Status: T0 scaffold.** The toolchain, CI and release pipeline are real and
> green, but the markdown renderer is not implemented yet. `convert()` currently
> emits each blank-line-separated block as a plain paragraph and returns a
> warning saying so. Nothing here is published to npm yet.

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

```js
import { convert } from "downword";

const { bytes } = await convert("# Hello\n\nFrom **downword**.");
// bytes is a Uint8Array containing a real .docx
```

In the browser:

```js
const { bytes, mimeType } = await convert(markdown);
const url = URL.createObjectURL(new Blob([bytes], { type: mimeType }));
```

Subpath entry points (stubs today):

```js
import { mathPlugin } from "downword/plugins/math";
import { mermaidPlugin } from "downword/plugins/mermaid";
```

## Why

Every "markdown to Word" tool either uploads your document to somebody's server
or hands you an HTML file with a `.doc` extension that Word complains about.
downword builds a real OOXML package in your browser; nothing leaves the tab.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md). Short version:

```sh
pnpm install
pnpm test
pnpm build
```

## License

[MIT](LICENSE) © 2026 KSPR Technologies
