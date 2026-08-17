/**
 * Naming the ```` ```mermaid ```` fences the CLI cannot draw.
 *
 * mermaid measures text by laying it out in a DOM, so there is no headless path
 * that is not a headless *browser*. `downword/plugins/mermaid` says so itself
 * and degrades cleanly — but only if you run it, and running it means importing
 * a browser-only entry point into a Node process to be told "not here". The CLI
 * does not: it finds the fences during the parse it was already doing, leaves
 * them as code blocks (which is what the renderer does with any fence), and
 * says which one it skipped.
 *
 * The detection is a **markdown-it core rule**, added through
 * `ConvertOptions.plugins`, which is the supported extension point. That buys
 * two things a regular expression over the source cannot:
 *
 *  - **no false positives** — a ```` ```mermaid ```` inside a four-backtick
 *    fence, or inside an indented code block, is text and is not reported; and
 *  - **a line number**, from the token's source map, so the message points at
 *    the fence rather than at the file.
 *
 * Nothing in this module imports mermaid, `downword/plugins/mermaid`, or
 * anything else that needs a DOM; `tests/startup.test.ts` proves it by listing
 * every script a real conversion loads.
 */

import type { MarkdownItPlugin } from "downword";

/** One ```` ```mermaid ```` fence, as found in the source. */
export interface MermaidFence {
  /** 1-based line of the opening fence. */
  readonly line: number;
  /** The rest of the info string (`mermaid The request pipeline`), or `null`. */
  readonly caption: string | null;
  /** First non-blank, non-comment line of the diagram, e.g. `"flowchart LR"`. */
  readonly opening: string;
}

/** Whether an info string's first word is `mermaid`, however it is cased. */
function isMermaidInfo(info: string): boolean {
  const first = info.trim().split(/\s+/, 1)[0] ?? "";
  return first.toLowerCase() === "mermaid";
}

/** The caption an info string carries after the language, or `null`. */
function captionOf(info: string): string | null {
  const rest = info
    .trim()
    .replace(/^\S+\s*/, "")
    .trim();
  return rest === "" ? null : rest;
}

/** The first line that says what kind of diagram this is. */
function openingOf(source: string): string {
  for (const line of source.split("\n")) {
    const trimmed = line.trim();
    // `%%` is a mermaid comment, and `%%{init: …}%%` a directive; neither names
    // the diagram. `---` opens the YAML frontmatter mermaid 10+ accepts.
    if (trimmed === "" || trimmed.startsWith("%%") || trimmed === "---") continue;
    return trimmed;
  }
  return "";
}

/**
 * A markdown-it plugin that reports every mermaid fence it sees.
 *
 * @param onFence - Called once per fence, in source order.
 * @returns The plugin, for `ConvertOptions.plugins`.
 */
export function scanMermaidFences(onFence: (fence: MermaidFence) => void): MarkdownItPlugin {
  return (md) => {
    md.core.ruler.push("downword-cli:mermaid-scan", (state) => {
      for (const token of state.tokens) {
        if (token.type !== "fence" || !isMermaidInfo(token.info)) continue;
        onFence({
          // `map` is [startLine, endLine), 0-based, and is set on every block
          // token markdown-it produces from real source.
          line: (token.map?.[0] ?? 0) + 1,
          caption: captionOf(token.info),
          opening: openingOf(token.content),
        });
      }
    });
  };
}

/**
 * What to tell the user about a fence that stayed a fence.
 *
 * Names the diagram (its caption, or its opening line, or its position), says
 * why, and gives the two ways out.
 *
 * @param fence - From {@link scanMermaidFences}.
 * @param index - 0-based position among the document's fences, for the fallback name.
 * @returns The message, without the `downword:` prefix or a trailing newline.
 */
export function describeSkippedFence(fence: MermaidFence, index: number): string {
  const name =
    fence.caption !== null
      ? `"${fence.caption}"`
      : fence.opening !== ""
        ? `"${fence.opening}"`
        : `#${index + 1}`;
  return (
    `left the \`\`\`mermaid fence ${name} as a code block: ` +
    `drawing a diagram needs a browser DOM, which a CLI does not have. ` +
    `Render it to an image and reference it with ![](diagram.png), ` +
    `or convert in the browser with downword/plugins/mermaid.`
  );
}
