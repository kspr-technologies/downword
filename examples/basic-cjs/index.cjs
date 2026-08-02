// CommonJS half of the dual-format smoke test.
//
//   pnpm --filter @downword/example-basic-cjs start
//
// Proves the "require" condition of downword's exports map resolves to a real
// CJS build. Writes examples/basic-cjs/out/basic-cjs.docx for the LibreOffice
// validity gate.

"use strict";

const { mkdirSync, writeFileSync } = require("node:fs");
const path = require("node:path");

const { convert, DOCX_MIME_TYPE, VERSION } = require("downword");

const markdown = `# Hello from CommonJS

downword ${VERSION} generated this file with \`require("downword")\`.
`;

async function main() {
  const { bytes, mimeType, warnings } = await convert(markdown, {
    title: "downword CJS example",
    creator: "KSPR Technologies",
  });

  if (mimeType !== DOCX_MIME_TYPE) {
    throw new Error(`unexpected mime type: ${mimeType}`);
  }

  const outDir = path.join(__dirname, "out");
  mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, "basic-cjs.docx");
  writeFileSync(outFile, bytes);

  console.log(`[cjs] downword ${VERSION} -> ${outFile} (${bytes.byteLength} bytes)`);
  for (const warning of warnings) {
    console.log(`[cjs] warning: ${warning}`);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
