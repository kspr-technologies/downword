# downword

> Markdown to Word (.docx) converter that runs entirely in your browser.
> Paste from ChatGPT/Claude, get a real Word document.

<!-- TODO(README): placeholder. The full package README is written in a later task. -->

> **Status: T0 scaffold.** The markdown renderer is not implemented yet.
> `convert()` currently emits each blank-line-separated block as a plain
> paragraph and returns a warning saying so.

```sh
npm install downword
```

```js
import { convert } from "downword";

const { bytes, mimeType, warnings } = await convert("# Hello\n\nFrom **downword**.");

// browser
const url = URL.createObjectURL(new Blob([bytes], { type: mimeType }));

// node
await writeFile("out.docx", bytes);
```

Subpath entry points (stubs today):

```js
import { mathPlugin } from "downword/plugins/math";
import { mermaidPlugin } from "downword/plugins/mermaid";
```

- Source, issues and docs: <https://github.com/kspr-technologies/downword>
- CLI: [`downword-cli`](https://www.npmjs.com/package/downword-cli)

MIT © 2026 KSPR Technologies
