/**
 * Which fences are diagrams, and what to call them.
 *
 * ```` ```mermaid ```` is the only spelling anyone writes (it is what GitHub,
 * GitLab, Obsidian, Notion and every LLM emit), matched case-insensitively
 * because `Mermaid` happens. The info string's *remainder* becomes the caption,
 * which is the one piece of authored text a diagram can carry:
 *
 * ````text
 * ```mermaid The request pipeline
 * flowchart LR
 *   A --> B
 * ```
 * ````
 *
 * A caption can also come from mermaid's own YAML frontmatter (`title:`), which
 * is where mermaid 10+ puts it, so a diagram authored for github.com keeps its
 * title here without being rewritten.
 */

import type { BlockNode, CodeBlockNode } from "../model.js";

/** Whether a block is a ```` ```mermaid ```` fence. */
export function isMermaidFence(node: BlockNode): node is CodeBlockNode {
  return node.type === "codeBlock" && node.lang !== null && node.lang.toLowerCase() === "mermaid";
}

/**
 * Human names for the diagram kinds mermaid can parse.
 *
 * Used for the image's alt text, which is what a screen reader announces and
 * what Word shows in the Alt Text pane — so "Mermaid sequence diagram" rather
 * than "sequenceDiagram", and never the raw source, which is not a description.
 */
const DIAGRAM_KINDS: Readonly<Record<string, string>> = {
  flowchart: "flowchart",
  "flowchart-elk": "flowchart",
  graph: "flowchart",
  sequencediagram: "sequence diagram",
  classdiagram: "class diagram",
  "classdiagram-v2": "class diagram",
  statediagram: "state diagram",
  "statediagram-v2": "state diagram",
  erdiagram: "entity-relationship diagram",
  journey: "user journey diagram",
  gantt: "Gantt chart",
  pie: "pie chart",
  quadrantchart: "quadrant chart",
  requirementdiagram: "requirement diagram",
  gitgraph: "Git graph",
  mindmap: "mind map",
  timeline: "timeline",
  zenuml: "ZenUML sequence diagram",
  "sankey-beta": "Sankey diagram",
  "xychart-beta": "XY chart",
  "block-beta": "block diagram",
  "packet-beta": "packet diagram",
  architecture: "architecture diagram",
  c4context: "C4 context diagram",
};

/** Strips mermaid's `%%{init: …}%%` directives and `%%` comments from a line. */
function isNoise(line: string): boolean {
  return line === "" || line.startsWith("%%");
}

/**
 * The `---` … `---` YAML block mermaid 10+ accepts at the top of a diagram.
 *
 * Only `title` is read, and only as a flat scalar: this is a caption lookup,
 * not a YAML parser, and dragging one in for a single key would be a strange
 * dependency for a package this size.
 */
function frontmatterTitle(lines: readonly string[]): string | null {
  if ((lines[0] ?? "").trim() !== "---") return null;
  for (let index = 1; index < lines.length; index += 1) {
    const line = (lines[index] ?? "").trim();
    if (line === "---") return null;
    const match = /^title\s*:\s*(.+)$/.exec(line);
    if (match !== null) return unquote((match[1] ?? "").trim());
  }
  return null;
}

/** Removes one layer of matching quotes, which an info string often carries. */
function unquote(value: string): string {
  const first = value.charAt(0);
  if (value.length >= 2 && (first === '"' || first === "'") && value.endsWith(first)) {
    return value.slice(1, -1).trim();
  }
  return value;
}

/** What a fence says about itself. */
export interface MermaidFenceInfo {
  /** The diagram source, with trailing whitespace trimmed. */
  readonly source: string;
  /** The authored caption, or `null` when the fence carries none. */
  readonly caption: string | null;
  /** Human name of the diagram kind, e.g. `"sequence diagram"`. */
  readonly kind: string;
  /** Alt text: the caption if there is one, otherwise a description of the kind. */
  readonly alt: string;
}

/**
 * Reads a mermaid fence.
 *
 * @param node - A code block for which {@link isMermaidFence} is true.
 * @returns The source, its caption and a description for the alt text. `source`
 *   is empty when the fence held nothing but whitespace, which the caller
 *   treats as `empty-diagram`.
 */
export function readMermaidFence(node: CodeBlockNode): MermaidFenceInfo {
  const source = node.value.replace(/\s+$/, "");
  const lines = source.split("\n");

  const meta = node.meta === null ? "" : unquote(node.meta.trim());
  const caption = meta !== "" ? meta : frontmatterTitle(lines);

  const first = lines.find((line) => !isNoise(line.trim()))?.trim() ?? "";
  // `flowchart LR`, `stateDiagram-v2`, `sankey-beta` - the keyword is the first
  // token, and the direction or configuration follows it.
  const keyword = (/^([A-Za-z][\w-]*)/.exec(first)?.[1] ?? "").toLowerCase();
  const kind = DIAGRAM_KINDS[keyword] ?? "diagram";

  return {
    source,
    caption,
    kind,
    alt: caption ?? `Mermaid ${kind}`,
  };
}
