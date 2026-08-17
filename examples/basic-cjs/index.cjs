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

> Nothing left the process to build it.
`;

async function main() {
  if (typeof DOCX_MIME_TYPE !== "string" || !DOCX_MIME_TYPE.includes("wordprocessingml")) {
    throw new Error(`unexpected mime type constant: ${DOCX_MIME_TYPE}`);
  }

  const warnings = [];
  const bytes = await convert(markdown, {
    metadata: { title: "downword CJS example", author: "KSPR Technologies" },
    onWarning: (warning) => warnings.push(warning),
  });

  const outDir = path.join(__dirname, "out");
  mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, "basic-cjs.docx");
  writeFileSync(outFile, bytes);

  console.log(`[cjs] downword ${VERSION} -> ${outFile} (${bytes.byteLength} bytes)`);
  for (const warning of warnings) {
    console.log(`[cjs] ${warning.stage} warning (${warning.code}): ${warning.message}`);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
