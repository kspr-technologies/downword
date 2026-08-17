// ESM half of the dual-format smoke test.
//
//   pnpm --filter @downword/example-basic-node start
//
// Writes examples/basic-node/out/basic-node.docx, which CI's LibreOffice gate
// (scripts/docx-validity.mjs) then converts to PDF.

import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { convert, VERSION } from "downword";

const markdown = `# Hello from ESM

downword ${VERSION} generated this file with \`import { convert } from "downword"\`.

- Paste your ChatGPT or Claude answer
- Get a real Word document

| Runs offline | Real .docx |
| ------------ | ---------- |
| yes          | yes        |
`;

const warnings = [];
const bytes = await convert(markdown, {
  metadata: { title: "downword ESM example", author: "KSPR Technologies" },
  onWarning: (warning) => warnings.push(warning),
});

const outDir = fileURLToPath(new URL("./out/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outFile = `${outDir}basic-node.docx`;
await writeFile(outFile, bytes);

console.log(`[esm] downword ${VERSION} -> ${outFile} (${bytes.byteLength} bytes)`);
for (const warning of warnings) {
  console.log(`[esm] ${warning.stage} warning (${warning.code}): ${warning.message}`);
}
