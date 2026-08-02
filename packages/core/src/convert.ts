import { Document, Packer, Paragraph, TextRun } from "docx";

import {
  DOCX_MIME_TYPE,
  type ConvertOptions,
  type ConvertResult,
  type DownwordPluginContext,
} from "./types.js";

/** Version of this package, injected at build time. See `types/globals.d.ts`. */
export const VERSION: string = __DOWNWORD_VERSION__;

const PLACEHOLDER_WARNING =
  "downword: the markdown renderer is not implemented yet; input was emitted as plain paragraphs.";

/**
 * Splits raw text into paragraph-sized blocks on blank lines.
 *
 * Placeholder for the real block parser (see CONTRIBUTING.md -> "Adding a
 * renderer node type"). Deliberately dependency-free and deterministic so the
 * golden tests are byte-stable.
 */
function splitBlocks(input: string): string[] {
  return input
    .replace(/\r\n?/g, "\n")
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter((block) => block.length > 0);
}

/**
 * Converts markdown to a Word `.docx` document.
 *
 * T0 scaffold: the markdown is **not** parsed yet. Each blank-line-separated
 * block becomes one plain paragraph, and a warning is returned. The pipeline is
 * wired end to end (docx writer -> real OOXML package -> LibreOffice validity
 * gate) so that every later change is exercised by CI from day one.
 */
export async function convert(
  markdown: string,
  options: ConvertOptions = {},
): Promise<ConvertResult> {
  const warnings: string[] = [PLACEHOLDER_WARNING];

  const context: DownwordPluginContext = {
    version: VERSION,
    warn: (message: string) => {
      warnings.push(message);
    },
  };

  for (const plugin of options.plugins ?? []) {
    plugin.setup?.(context);
  }

  const blocks = splitBlocks(markdown);
  const children =
    blocks.length > 0
      ? blocks.map((block) => new Paragraph({ children: [new TextRun({ text: block })] }))
      : [new Paragraph({ children: [] })];

  const doc = new Document({
    sections: [{ children }],
    ...(options.title !== undefined ? { title: options.title } : {}),
    ...(options.creator !== undefined ? { creator: options.creator } : {}),
    ...(options.description !== undefined ? { description: options.description } : {}),
  });

  const arrayBuffer = await Packer.toArrayBuffer(doc);

  return {
    bytes: new Uint8Array(arrayBuffer),
    mimeType: DOCX_MIME_TYPE,
    warnings,
  };
}
