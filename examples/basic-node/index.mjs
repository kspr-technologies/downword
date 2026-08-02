// ESM half of the dual-format smoke test.
//
//   pnpm --filter @downword/example-basic-node start
//
// Writes examples/basic-node/out/basic-node.docx, which CI's LibreOffice gate
// (scripts/docx-validity.mjs) then converts to PDF.

import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { convert, DOCX_MIME_TYPE, VERSION } from "downword";

const markdown = `# Hello from ESM

downword ${VERSION} generated this file with \`import { convert } from "downword"\`.

Paste your ChatGPT or Claude answer here and you get a real Word document.
`;

const { bytes, mimeType, warnings } = await convert(markdown, {
  title: "downword ESM example",
  creator: "KSPR Technologies",
});

if (mimeType !== DOCX_MIME_TYPE) {
  throw new Error(`unexpected mime type: ${mimeType}`);
}

const outDir = fileURLToPath(new URL("./out/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outFile = `${outDir}basic-node.docx`;
await writeFile(outFile, bytes);

console.log(`[esm] downword ${VERSION} -> ${outFile} (${bytes.byteLength} bytes)`);
for (const warning of warnings) {
  console.log(`[esm] warning: ${warning}`);
}
