import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import type { MarkdownIt, StateBlock, StateInline } from "markdown-it";
import { describe, expect, it } from "vitest";

import {
  assertNever,
  bold,
  inlineCode,
  italic,
  link,
  nodesOfType,
  strikethrough,
  text,
  type BlockNode,
  type DocumentNode,
  type InlineNode,
  type ListNode,
  type Node,
} from "../src/model.js";
import {
  parseMarkdown,
  type MarkdownItPlugin,
  type ParseOptions,
  type ParseWarning,
} from "../src/parse/index.js";

/* -------------------------------------------------------------------------- */
/* Fixtures                                                                    */
/* -------------------------------------------------------------------------- */

const FIXTURE_DIR = fileURLToPath(new URL("./__fixtures__/markdown/", import.meta.url));

/** Every fixture file, in filename order (they are numbered for that reason). */
const FIXTURES: readonly string[] = readdirSync(FIXTURE_DIR)
  .filter((name) => name.endsWith(".md"))
  .sort();

function readFixture(name: string): string {
  return readFileSync(join(FIXTURE_DIR, name), "utf8");
}

/**
 * A minimal `math_inline` / `math_block` plugin.
 *
 * Real math comes from `@vscode/markdown-it-katex`, which this package
 * deliberately does not depend on (see `ParseOptions.plugins`). This stub emits
 * exactly the two token types that plugin emits — `content` holding the raw TeX
 * — which is the whole contract the parser cares about, and it keeps the test
 * suite free of a 500 kB dependency.
 */
const mathPlugin: MarkdownItPlugin = (md: MarkdownIt): void => {
  md.inline.ruler.before("escape", "math_inline", (state: StateInline, silent: boolean) => {
    const start = state.pos;
    if (state.src[start] !== "$") return false;

    const end = state.src.indexOf("$", start + 1);
    if (end < 0 || end === start + 1) return false;

    const content = state.src.slice(start + 1, end);
    if (/^\s|\s$/.test(content) || content.includes("\n")) return false;

    if (!silent) {
      const token = state.push("math_inline", "math", 0);
      token.markup = "$";
      token.content = content;
    }
    state.pos = end + 1;
    return true;
  });

  /** One source line of a block state; `noUncheckedIndexedAccess` needs the guards. */
  const lineText = (state: StateBlock, line: number): string => {
    const begin = state.bMarks[line] ?? 0;
    const shift = state.tShift[line] ?? 0;
    const end = state.eMarks[line] ?? begin;
    return state.src.slice(begin + shift, end);
  };

  md.block.ruler.before(
    "fence",
    "math_block",
    (state: StateBlock, startLine: number, endLine: number, silent: boolean) => {
      if (lineText(state, startLine).trim() !== "$$") return false;
      if (silent) return true;

      let line = startLine + 1;
      while (line < endLine) {
        if (lineText(state, line).trim() === "$$") break;
        line += 1;
      }

      const token = state.push("math_block", "math", 0);
      token.block = true;
      token.markup = "$$";
      token.content = state.getLines(startLine + 1, line, state.blkIndent, false);
      token.map = [startLine, line + 1];
      state.line = line + 1;
      return true;
    },
  );
};

/** Per-fixture option overrides. Everything not listed uses the defaults. */
const FIXTURE_OPTIONS: Readonly<Record<string, ParseOptions>> = {
  "24-html-blocks-and-inline.md": { html: true },
  "25-math.md": { plugins: [mathPlugin] },
  "29-llm-readme-output.md": { html: true },
};

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

/** Parses a string, collecting warnings instead of discarding them. */
function parseWithWarnings(
  source: string,
  options: ParseOptions = {},
): { readonly document: DocumentNode; readonly warnings: readonly ParseWarning[] } {
  const warnings: ParseWarning[] = [];
  const document = parseMarkdown(source, {
    ...options,
    onWarning: (warning) => warnings.push(warning),
  });
  return { document, warnings };
}

/* -------------------------------------------------------------------------- */
/* Outline rendering                                                           */
/* -------------------------------------------------------------------------- */

/**
 * Renders a document as a compact indented outline.
 *
 * The full-model snapshots are exhaustive but ~10k lines, which nobody reviews.
 * This second, lossy-but-readable view is what a human actually diffs when a
 * fixture's output changes: one line per node, with the fields that carry the
 * parsing decisions (list kind/tightness/start, table alignment, mark
 * composition, footnote numbering) spelled out inline.
 */
function outline(document: DocumentNode): string {
  const lines: string[] = [];
  for (const child of document.children) writeNode(child, 0, lines);
  return lines.length > 0 ? lines.join("\n") : "(empty document)";
}

function writeNode(node: Node, depth: number, lines: string[]): void {
  const pad = "  ".repeat(depth);

  switch (node.type) {
    case "paragraph":
      lines.push(`${pad}p ${inlineSummary(node.children)}`);
      return;
    case "heading":
      lines.push(`${pad}h${String(node.level)} #${node.id} ${inlineSummary(node.children)}`);
      return;
    case "list": {
      const start = node.start === null ? "" : ` start=${String(node.start)}`;
      lines.push(`${pad}list ${node.kind}${start} ${node.tight ? "tight" : "loose"}`);
      break;
    }
    case "listItem": {
      const box = node.checked === null ? "" : node.checked ? "[x] " : "[ ] ";
      lines.push(`${pad}item ${box}`.trimEnd());
      break;
    }
    case "table":
      lines.push(`${pad}table align=[${node.align.join(",")}]`);
      break;
    case "tableRow":
      lines.push(`${pad}row${node.header ? " header" : ""}`);
      break;
    case "tableCell":
      lines.push(`${pad}cell ${inlineSummary(node.children)}`);
      return;
    case "codeBlock": {
      const lang = node.lang === null ? "-" : node.lang;
      const meta = node.meta === null ? "" : ` meta=${JSON.stringify(node.meta)}`;
      lines.push(`${pad}code lang=${lang}${meta} ${JSON.stringify(node.value)}`);
      return;
    }
    case "blockquote":
      lines.push(`${pad}quote`);
      break;
    case "thematicBreak":
      lines.push(`${pad}hr`);
      return;
    case "htmlBlock":
      lines.push(`${pad}html ${JSON.stringify(node.value)}`);
      return;
    case "footnoteDefinition":
      lines.push(
        `${pad}footnote[${String(node.number)}] id=${node.identifier} label=${node.label}`,
      );
      break;
    case "mathBlock":
      lines.push(`${pad}math ${JSON.stringify(node.value)}`);
      return;
    case "document":
      break;
    case "text":
    case "hardBreak":
    case "softBreak":
    case "image":
    case "htmlInline":
    case "footnoteReference":
    case "mathInline":
      lines.push(`${pad}${inlineSummary([node])}`);
      return;
    default:
      return assertNever(node, "node");
  }

  for (const child of childrenOfNode(node)) writeNode(child, depth + 1, lines);
}

/** Block/structural children only; inline children go through `inlineSummary`. */
function childrenOfNode(node: Node): readonly Node[] {
  switch (node.type) {
    case "document":
    case "list":
    case "listItem":
    case "table":
    case "tableRow":
    case "blockquote":
    case "footnoteDefinition":
      return node.children;
    default:
      return [];
  }
}

/** One-line rendering of a run of inline nodes, marks included. */
function inlineSummary(nodes: readonly InlineNode[]): string {
  return nodes.map(inlineToken).join(" ");
}

function inlineToken(node: InlineNode): string {
  const marks = node.marks
    .map((mark) =>
      mark.type === "link"
        ? `link:${mark.href}${mark.title === null ? "" : ` title=${JSON.stringify(mark.title)}`}`
        : mark.type,
    )
    .join("+");
  const prefix = marks.length > 0 ? `<${marks}>` : "";

  switch (node.type) {
    case "text":
      return prefix + JSON.stringify(node.value);
    case "hardBreak":
      return `${prefix}<BR>`;
    case "softBreak":
      return `${prefix}<NL>`;
    case "image":
      return `${prefix}img(src=${node.src} alt=${JSON.stringify(node.alt)} title=${JSON.stringify(node.title)})`;
    case "htmlInline":
      return `${prefix}raw(${JSON.stringify(node.value)})`;
    case "footnoteReference":
      return `${prefix}fnref(${String(node.number)}:${node.identifier})`;
    case "mathInline":
      return `${prefix}math(${JSON.stringify(node.value)})`;
    default:
      return assertNever(node, "inline node");
  }
}

/** The single block at `index`, asserted to exist. */
function blockAt(document: DocumentNode, index: number): BlockNode {
  const block = document.children[index];
  if (block === undefined) throw new Error(`no block at index ${String(index)}`);
  return block;
}

/** The first list in a document, asserted to exist. */
function firstList(document: DocumentNode): ListNode {
  const [list] = nodesOfType(document, "list");
  if (list === undefined) throw new Error("document contains no list");
  return list;
}

/* -------------------------------------------------------------------------- */
/* Fixture snapshots                                                           */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: fixtures", () => {
  it("found every fixture", () => {
    expect(FIXTURES.length).toBeGreaterThanOrEqual(25);
    expect(FIXTURES).toMatchSnapshot("fixture list");
  });

  it.each(FIXTURES)("%s", (name) => {
    const { document, warnings } = parseWithWarnings(
      readFixture(name),
      FIXTURE_OPTIONS[name] ?? {},
    );

    expect(outline(document)).toMatchSnapshot("outline");
    expect(document).toMatchSnapshot("model");
    expect(warnings).toMatchSnapshot("warnings");
  });

  it("is deterministic: parsing every fixture twice gives identical models", () => {
    for (const name of FIXTURES) {
      const source = readFixture(name);
      const options = FIXTURE_OPTIONS[name] ?? {};
      expect(parseMarkdown(source, options)).toEqual(parseMarkdown(source, options));
    }
  });

  it("never leaks a markdown-it token into the model", () => {
    for (const name of FIXTURES) {
      const document = parseMarkdown(readFixture(name), FIXTURE_OPTIONS[name] ?? {});
      const serialised = JSON.stringify(document);
      expect(serialised).not.toContain('"nesting"');
      expect(serialised).not.toContain('"markup"');
    }
  });
});

/* -------------------------------------------------------------------------- */
/* Documents                                                                   */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: document", () => {
  it("returns an empty document for empty input", () => {
    expect(parseMarkdown("")).toEqual({
      type: "document",
      metadata: {
        title: null,
        author: null,
        description: null,
        keywords: [],
        date: null,
        custom: {},
      },
      children: [],
    });
  });

  it("returns an empty document for whitespace-only input", () => {
    expect(parseMarkdown("   \n\n\t\n").children).toEqual([]);
  });

  it("carries injected metadata through untouched", () => {
    const document = parseMarkdown("# Hi", {
      metadata: {
        title: "T",
        author: "A",
        description: null,
        keywords: ["k"],
        date: "2026-08-02",
        custom: { x: "y" },
      },
    });

    expect(document.metadata.title).toBe("T");
    expect(document.metadata.keywords).toEqual(["k"]);
  });

  it("accepts CRLF input", () => {
    expect(parseMarkdown("# A\r\n\r\nB\r\n")).toEqual(parseMarkdown("# A\n\nB\n"));
  });
});

/* -------------------------------------------------------------------------- */
/* Headings and slugs                                                          */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: headings", () => {
  it("maps h1..h6", () => {
    const document = parseMarkdown("# a\n\n## b\n\n### c\n\n#### d\n\n##### e\n\n###### f\n");
    expect(nodesOfType(document, "heading").map((node) => node.level)).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it("makes duplicate heading slugs unique and deterministic", () => {
    const document = parseMarkdown("# Setup\n\n# Setup\n\n# Setup\n\n# setup\n");
    expect(nodesOfType(document, "heading").map((node) => node.id)).toEqual([
      "setup",
      "setup-1",
      "setup-2",
      "setup-3",
    ]);
  });

  it("strips trailing punctuation and emoji from slugs", () => {
    // `slugify` trims before it drops punctuation, so the space that preceded
    // the emoji survives as a trailing hyphen — the same thing GitHub does.
    const document = parseMarkdown("## Setup ⚙️!\n\n## Setup?\n");
    expect(nodesOfType(document, "heading").map((node) => node.id)).toEqual(["setup-", "setup"]);
  });

  it("falls back to a fixed slug when nothing survives", () => {
    const document = parseMarkdown("## 🚀\n\n## ✨\n");
    expect(nodesOfType(document, "heading").map((node) => node.id)).toEqual([
      "section",
      "section-1",
    ]);
  });

  it("does not include image alt text in a slug", () => {
    const document = parseMarkdown("## Build ![status](s.svg) badge\n");
    expect(nodesOfType(document, "heading")[0]?.id).toBe("build-badge");
  });

  it("restarts slug allocation for each document", () => {
    expect(nodesOfType(parseMarkdown("# Setup"), "heading")[0]?.id).toBe("setup");
    expect(nodesOfType(parseMarkdown("# Setup"), "heading")[0]?.id).toBe("setup");
  });

  it("parses setext headings", () => {
    const document = parseMarkdown("Title\n=====\n\nSub\n---\n");
    expect(nodesOfType(document, "heading").map((node) => node.level)).toEqual([1, 2]);
  });
});

/* -------------------------------------------------------------------------- */
/* Marks                                                                       */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: mark composition", () => {
  it("composes bold inside a link", () => {
    const document = parseMarkdown("[**a**](https://x.io)");
    expect(blockAt(document, 0)).toEqual({
      type: "paragraph",
      children: [text("a", [link("https://x.io"), bold()])],
    });
  });

  it("composes code inside bold inside a link", () => {
    const document = parseMarkdown("[**`npm i`**](https://x.io)");
    expect(blockAt(document, 0)).toEqual({
      type: "paragraph",
      children: [text("npm i", [link("https://x.io"), bold(), inlineCode()])],
    });
  });

  it("normalises mark order regardless of nesting order", () => {
    const outer = parseMarkdown("***a***");
    const inner = parseMarkdown("*__a__*");
    expect(outer).toEqual(inner);
    expect(blockAt(outer, 0)).toEqual({
      type: "paragraph",
      children: [text("a", [bold(), italic()])],
    });
  });

  it("carries a link title", () => {
    const document = parseMarkdown('[a](https://x.io "T")');
    expect(blockAt(document, 0)).toEqual({
      type: "paragraph",
      children: [text("a", [link("https://x.io", "T")])],
    });
  });

  it("composes strikethrough with bold", () => {
    expect(blockAt(parseMarkdown("~~**a**~~"), 0)).toEqual({
      type: "paragraph",
      children: [text("a", [bold(), strikethrough()])],
    });
  });

  it("drops the zero-length text runs that ***x*** produces", () => {
    const document = parseMarkdown("***x***");
    expect(nodesOfType(document, "text").map((node) => node.value)).toEqual(["x"]);
  });

  it("keeps intraword underscores literal", () => {
    expect(nodesOfType(parseMarkdown("snake_case_name"), "text").map((n) => n.value)).toEqual([
      "snake_case_name",
    ]);
  });

  it("marks an image that sits inside a link", () => {
    const document = parseMarkdown("[![badge](b.svg)](https://ci.example.com)");
    expect(nodesOfType(document, "image")[0]).toEqual({
      type: "image",
      src: "b.svg",
      alt: "badge",
      title: null,
      resolved: null,
      marks: [link("https://ci.example.com")],
    });
  });
});

/* -------------------------------------------------------------------------- */
/* Lists                                                                       */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: lists", () => {
  it("detects a tight list", () => {
    expect(firstList(parseMarkdown("- a\n- b\n")).tight).toBe(true);
  });

  it("detects a loose list", () => {
    expect(firstList(parseMarkdown("- a\n\n- b\n")).tight).toBe(false);
  });

  it("does not let a nested loose list make its tight parent loose", () => {
    const document = parseMarkdown("- a\n  - x\n\n    y\n- b\n");
    const [outer, inner] = nodesOfType(document, "list");
    expect(outer?.tight).toBe(true);
    expect(inner?.tight).toBe(false);
  });

  it("reads an ordered list's start ordinal", () => {
    expect(firstList(parseMarkdown("5. five\n6. six\n")).start).toBe(5);
  });

  it("defaults the start ordinal to 1", () => {
    expect(firstList(parseMarkdown("1. one\n")).start).toBe(1);
  });

  it("reads a zero start ordinal", () => {
    expect(firstList(parseMarkdown("0. zero\n")).start).toBe(0);
  });

  it("restarts a second list rather than continuing the first", () => {
    const document = parseMarkdown("1. a\n\ntext\n\n1. b\n");
    expect(nodesOfType(document, "list").map((node) => node.start)).toEqual([1, 1]);
  });

  it("forces start to null for bullet lists", () => {
    expect(firstList(parseMarkdown("- a\n")).start).toBeNull();
  });

  it("nests four levels of mixed ordered and bullet lists", () => {
    const document = parseMarkdown("1. a\n   - b\n     1. c\n        - d\n");
    expect(nodesOfType(document, "list").map((node) => node.kind)).toEqual([
      "ordered",
      "bullet",
      "ordered",
      "bullet",
    ]);
  });

  it("keeps a fenced code block inside an ordered list item", () => {
    const document = parseMarkdown("1. Install:\n\n   ```bash\n   npm i\n   ```\n");
    const [item] = nodesOfType(document, "listItem");
    expect(item?.children.map((child) => child.type)).toEqual(["paragraph", "codeBlock"]);
    expect(nodesOfType(document, "codeBlock")[0]).toEqual({
      type: "codeBlock",
      lang: "bash",
      meta: null,
      value: "npm i\n",
      highlights: null,
    });
  });

  it("keeps a blockquote inside a list item", () => {
    const document = parseMarkdown("1. a\n\n   > quoted\n");
    const [item] = nodesOfType(document, "listItem");
    expect(item?.children.map((child) => child.type)).toEqual(["paragraph", "blockquote"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Task lists                                                                  */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: task lists", () => {
  it("reads mixed checked state and strips the marker", () => {
    const document = parseMarkdown("- [x] done\n- [ ] todo\n- plain\n");
    const list = firstList(document);

    expect(list.kind).toBe("task");
    expect(list.children.map((item) => item.checked)).toEqual([true, false, null]);
    expect(nodesOfType(document, "text").map((node) => node.value)).toEqual([
      "done",
      "todo",
      "plain",
    ]);
  });

  it("accepts an uppercase X", () => {
    expect(firstList(parseMarkdown("- [X] done\n")).children[0]?.checked).toBe(true);
  });

  it("accepts a checkbox with no text after it", () => {
    const list = firstList(parseMarkdown("- [x]\n"));
    expect(list.children[0]?.checked).toBe(true);
    expect(list.children[0]?.children).toEqual([{ type: "paragraph", children: [] }]);
  });

  it("rejects a checkbox with no space after the bracket", () => {
    expect(firstList(parseMarkdown("- [ ]no space\n")).children[0]?.checked).toBeNull();
  });

  it("rejects an empty bracket pair", () => {
    expect(firstList(parseMarkdown("- [] nope\n")).children[0]?.checked).toBeNull();
  });

  it("rejects a marked-up checkbox", () => {
    expect(firstList(parseMarkdown("- **[x] nope**\n")).children[0]?.checked).toBeNull();
  });

  it("nests task lists", () => {
    const document = parseMarkdown("- [x] outer\n  - [ ] inner\n");
    expect(nodesOfType(document, "list").map((node) => node.kind)).toEqual(["task", "task"]);
  });

  it("keeps an ordered list ordered even when its items are checkboxes", () => {
    const list = firstList(parseMarkdown("1. [x] a\n2. [ ] b\n"));
    expect(list.kind).toBe("ordered");
    expect(list.start).toBe(1);
    expect(list.children.map((item) => item.checked)).toEqual([true, false]);
  });

  it("keeps marks that follow the stripped marker", () => {
    const document = parseMarkdown("- [x] done **now**\n");
    expect(blockAt(document, 0)).toMatchObject({
      children: [
        {
          checked: true,
          children: [{ children: [text("done "), text("now", [bold()])] }],
        },
      ],
    });
  });
});

/* -------------------------------------------------------------------------- */
/* Tables                                                                      */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: tables", () => {
  it("reads per-column alignment once, on the table", () => {
    const document = parseMarkdown("| a | b | c | d |\n|:--|:-:|--:|---|\n| 1 | 2 | 3 | 4 |\n");
    const [table] = nodesOfType(document, "table");
    expect(table?.align).toEqual(["left", "center", "right", "none"]);
  });

  it("marks the header row and only the header row", () => {
    const document = parseMarkdown("| a |\n| - |\n| 1 |\n| 2 |\n");
    expect(nodesOfType(document, "tableRow").map((row) => row.header)).toEqual([
      true,
      false,
      false,
    ]);
  });

  it("keeps inline formatting inside cells", () => {
    const document = parseMarkdown("| a |\n| - |\n| **b** `c` [d](https://x.io) |\n");
    const cells = nodesOfType(document, "tableCell");
    expect(cells[1]?.children).toEqual([
      text("b", [bold()]),
      text(" "),
      text("c", [inlineCode()]),
      text(" "),
      text("d", [link("https://x.io")]),
    ]);
  });

  it("unescapes a pipe inside a cell", () => {
    const document = parseMarkdown("| a |\n| - |\n| x \\| y |\n");
    expect(nodesOfType(document, "tableCell")[1]?.children).toEqual([text("x | y")]);
  });

  it("handles a header-only table", () => {
    const document = parseMarkdown("| a | b |\n| - | - |\n");
    const [table] = nodesOfType(document, "table");
    expect(table?.children.length).toBe(1);
    expect(table?.align).toEqual(["none", "none"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Code blocks                                                                 */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: code blocks", () => {
  it("splits a fence info string into lang and meta", () => {
    const document = parseMarkdown('```ts title="a.ts" {1,3}\nx\n```\n');
    expect(nodesOfType(document, "codeBlock")[0]).toEqual({
      type: "codeBlock",
      lang: "ts",
      meta: 'title="a.ts" {1,3}',
      value: "x\n",
      highlights: null,
    });
  });

  it("leaves lang and meta null for a bare fence", () => {
    const document = parseMarkdown("```\nx\n```\n");
    expect(nodesOfType(document, "codeBlock")[0]?.lang).toBeNull();
    expect(nodesOfType(document, "codeBlock")[0]?.meta).toBeNull();
  });

  it("recovers from an unclosed fence at EOF", () => {
    const document = parseMarkdown("```python\ndef f():\n    return 1\n");
    expect(nodesOfType(document, "codeBlock")[0]).toEqual({
      type: "codeBlock",
      lang: "python",
      meta: null,
      value: "def f():\n    return 1\n",
      highlights: null,
    });
  });

  it("reads an indented code block", () => {
    const document = parseMarkdown("    indented\n    lines\n");
    expect(nodesOfType(document, "codeBlock")[0]).toEqual({
      type: "codeBlock",
      lang: null,
      meta: null,
      value: "indented\nlines\n",
      highlights: null,
    });
  });
});

/* -------------------------------------------------------------------------- */
/* Breaks, entities, HTML                                                      */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: inline details", () => {
  it("distinguishes hard breaks from soft breaks", () => {
    const document = parseMarkdown("a  \nb\\\nc\nd\n");
    expect(blockAt(document, 0)).toEqual({
      type: "paragraph",
      children: [
        text("a"),
        { type: "hardBreak", marks: [] },
        text("b"),
        { type: "hardBreak", marks: [] },
        text("c"),
        { type: "softBreak", marks: [] },
        text("d"),
      ],
    });
  });

  it("turns every newline into a hard break when breaks: true", () => {
    const document = parseMarkdown("a\nb\n", { breaks: true });
    expect(nodesOfType(document, "softBreak")).toEqual([]);
    expect(nodesOfType(document, "hardBreak").length).toBe(1);
  });

  it("decodes HTML entities exactly once", () => {
    expect(nodesOfType(parseMarkdown("&amp;amp; &copy;"), "text").map((n) => n.value)).toEqual([
      "&amp; ©",
    ]);
  });

  it("keeps entities literal inside code", () => {
    expect(nodesOfType(parseMarkdown("`&amp;`"), "text")[0]?.value).toBe("&amp;");
  });

  it("leaves raw HTML as text when html is false", () => {
    const document = parseMarkdown("a <br> b\n");
    expect(nodesOfType(document, "htmlInline")).toEqual([]);
    expect(nodesOfType(document, "text")[0]?.value).toBe("a <br> b");
  });

  it("models raw HTML when html is true", () => {
    const { document, warnings } = parseWithWarnings("a <br> b\n\n<div>x</div>\n", { html: true });
    expect(nodesOfType(document, "htmlInline").map((node) => node.value)).toEqual(["<br>"]);
    expect(nodesOfType(document, "htmlBlock").map((node) => node.value)).toEqual([
      "<div>x</div>\n",
    ]);
    expect(warnings.map((warning) => warning.code)).toEqual(["raw-html", "raw-html"]);
  });

  it("linkifies bare URLs by default", () => {
    const document = parseMarkdown("see https://bare.example.com now\n");
    expect(nodesOfType(document, "text")[1]).toEqual(
      text("https://bare.example.com", [link("https://bare.example.com")]),
    );
  });

  it("does not linkify when linkify is false", () => {
    const document = parseMarkdown("see https://bare.example.com now\n", { linkify: false });
    expect(nodesOfType(document, "text")[0]?.marks).toEqual([]);
  });

  it("keeps an image title and alt", () => {
    const document = parseMarkdown('![the alt](x.png "The Title")');
    expect(nodesOfType(document, "image")[0]).toEqual({
      type: "image",
      src: "x.png",
      alt: "the alt",
      title: "The Title",
      resolved: null,
      marks: [],
    });
  });

  it("flattens markup out of image alt text", () => {
    expect(nodesOfType(parseMarkdown("![**b** and `c`](x.png)"), "image")[0]?.alt).toBe("b and c");
  });
});

/* -------------------------------------------------------------------------- */
/* Footnotes                                                                   */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: footnotes", () => {
  it("numbers references and definitions from 1", () => {
    const document = parseMarkdown("a[^x] b[^y]\n\n[^x]: X\n\n[^y]: Y\n");

    expect(nodesOfType(document, "footnoteReference").map((node) => node.number)).toEqual([1, 2]);
    expect(nodesOfType(document, "footnoteDefinition").map((node) => node.number)).toEqual([1, 2]);
  });

  it("numbers by first reference, not by definition order", () => {
    const document = parseMarkdown("a[^b] c[^a]\n\n[^a]: A\n\n[^b]: B\n");
    expect(
      nodesOfType(document, "footnoteDefinition").map((node) => [node.identifier, node.number]),
    ).toEqual([
      ["b", 1],
      ["a", 2],
    ]);
  });

  it("gives repeated references to one note the same number", () => {
    const document = parseMarkdown("a[^x] b[^x]\n\n[^x]: X\n");
    expect(nodesOfType(document, "footnoteReference").map((node) => node.number)).toEqual([1, 1]);
    expect(nodesOfType(document, "footnoteDefinition").length).toBe(1);
  });

  it("keeps definitions in document.children, at the end", () => {
    const document = parseMarkdown("a[^x]\n\n[^x]: X\n");
    expect(document.children.map((child) => child.type)).toEqual([
      "paragraph",
      "footnoteDefinition",
    ]);
  });

  it("supports block content inside a definition", () => {
    const document = parseMarkdown("a[^x]\n\n[^x]: one\n\n    - two\n");
    const [definition] = nodesOfType(document, "footnoteDefinition");
    expect(definition?.children.map((child) => child.type)).toEqual(["paragraph", "list"]);
  });

  it("drops the back-link anchor markdown-it-footnote appends", () => {
    const document = parseMarkdown("a[^x]\n\n[^x]: X\n");
    const [definition] = nodesOfType(document, "footnoteDefinition");
    expect(definition?.children).toEqual([{ type: "paragraph", children: [text("X")] }]);
  });

  it("labels an unlabelled inline footnote with its number", () => {
    const document = parseMarkdown("a^[inline]\n");
    expect(nodesOfType(document, "footnoteReference")[0]).toMatchObject({
      identifier: "1",
      label: "1",
      number: 1,
    });
  });

  it("leaves footnote syntax as text when footnotes are disabled", () => {
    const document = parseMarkdown("a[^x]\n\n[^x]: X\n", { footnotes: false });
    expect(nodesOfType(document, "footnoteReference")).toEqual([]);
    expect(nodesOfType(document, "footnoteDefinition")).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Math                                                                        */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: math", () => {
  const options: ParseOptions = { plugins: [mathPlugin] };

  it("maps math_inline to a mathInline node", () => {
    const document = parseMarkdown("x $a+b$ y", options);
    expect(nodesOfType(document, "mathInline")[0]).toEqual({
      type: "mathInline",
      value: "a+b",
      omml: null,
      marks: [],
    });
  });

  it("maps math_block to a mathBlock node", () => {
    const document = parseMarkdown("$$\na+b\n$$\n", options);
    expect(nodesOfType(document, "mathBlock")[0]).toEqual({
      type: "mathBlock",
      value: "a+b",
      omml: null,
    });
  });

  it("carries marks onto inline math", () => {
    const document = parseMarkdown("**$a$**", options);
    expect(nodesOfType(document, "mathInline")[0]?.marks).toEqual([bold()]);
  });

  it("emits no math nodes when no math plugin is injected", () => {
    const document = parseMarkdown("x $a+b$ y");
    expect(nodesOfType(document, "mathInline")).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Warnings                                                                    */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: warnings", () => {
  it("stays silent on clean input", () => {
    const { warnings } = parseWithWarnings("# a\n\n- b\n\n| c |\n| - |\n| d |\n");
    expect(warnings).toEqual([]);
  });

  it("reports raw HTML with a source line", () => {
    const { warnings } = parseWithWarnings("para\n\n<div>\nx\n</div>\n", { html: true });
    expect(warnings).toEqual([
      // Raw HTML is a `notice`: the markup survives as text, nothing is lost.
      { code: "raw-html", severity: "notice", message: "raw HTML block: <div> x </div>", line: 3 },
    ]);
  });

  it("reports an unsupported token from an injected plugin", () => {
    const definitionList: MarkdownItPlugin = (md: MarkdownIt): void => {
      md.core.ruler.push("inject_unknown", (state) => {
        state.tokens.push(new state.Token("dl_open", "dl", 1));
        state.tokens.push(new state.Token("dl_close", "dl", -1));
      });
    };

    const { warnings } = parseWithWarnings("a\n", { plugins: [definitionList] });
    expect(warnings).toEqual([
      // A skipped token is an `error`: whatever it opened never reached the model.
      {
        code: "unsupported-token",
        severity: "error",
        message: 'unsupported block token "dl_open"',
        line: 1,
      },
    ]);
  });

  it("survives a plugin that emits an unbalanced inline close token", () => {
    const unbalanced: MarkdownItPlugin = (md: MarkdownIt): void => {
      md.core.ruler.push("inject_unbalanced", (state) => {
        for (const token of state.tokens) {
          if (token.type === "inline" && token.children !== null) {
            token.children.unshift(new state.Token("strong_close", "strong", -1));
          }
        }
      });
    };

    const { document, warnings } = parseWithWarnings("a\n", { plugins: [unbalanced] });
    expect(document.children).toEqual([{ type: "paragraph", children: [text("a")] }]);
    expect(warnings.map((warning) => warning.code)).toEqual(["unexpected-token"]);
  });

  it("does not build warnings at all when no handler is supplied", () => {
    expect(() => parseMarkdown("<div>x</div>", { html: true })).not.toThrow();
  });
});

/* -------------------------------------------------------------------------- */
/* Blockquotes                                                                 */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: blockquotes", () => {
  it("nests blockquotes", () => {
    const document = parseMarkdown("> a\n>\n> > b\n");
    expect(blockAt(document, 0)).toEqual({
      type: "blockquote",
      children: [
        { type: "paragraph", children: [text("a")] },
        { type: "blockquote", children: [{ type: "paragraph", children: [text("b")] }] },
      ],
    });
  });

  it("keeps a list containing code inside a blockquote", () => {
    const document = parseMarkdown("> - a\n>\n>   ```js\n>   x()\n>   ```\n");
    const [item] = nodesOfType(document, "listItem");
    expect(item?.children.map((child) => child.type)).toEqual(["paragraph", "codeBlock"]);
    expect(nodesOfType(document, "codeBlock")[0]?.value).toBe("x()\n");
  });
});

/* -------------------------------------------------------------------------- */
/* Termination                                                                 */
/* -------------------------------------------------------------------------- */

describe("parseMarkdown: termination", () => {
  it("terminates on nesting far deeper than markdown-it's own limit", () => {
    const document = parseMarkdown(`${"> ".repeat(200)}deep\n`);
    expect(nodesOfType(document, "blockquote").length).toBeGreaterThan(0);
  });

  it("terminates on a 5,000-item list", () => {
    const source = Array.from({ length: 5000 }, (_unused, index) => `- item ${String(index)}`).join(
      "\n",
    );
    expect(firstList(parseMarkdown(source)).children.length).toBe(5000);
  });

  it("terminates when a plugin leaves an opening token unclosed", () => {
    const truncate: MarkdownItPlugin = (md: MarkdownIt): void => {
      md.core.ruler.push("drop_closers", (state) => {
        state.tokens = state.tokens.filter((token) => token.type !== "blockquote_close");
      });
    };

    const document = parseMarkdown("> a\n\nb\n", { plugins: [truncate] });
    expect(blockAt(document, 0).type).toBe("blockquote");
    expect(document.children.length).toBe(1);
  });

  it("terminates when a plugin leaves a closing token orphaned", () => {
    const orphan: MarkdownItPlugin = (md: MarkdownIt): void => {
      md.core.ruler.push("orphan_closer", (state) => {
        state.tokens.push(new state.Token("blockquote_close", "blockquote", -1));
      });
    };

    const { document, warnings } = parseWithWarnings("a\n", { plugins: [orphan] });
    expect(document.children.map((child) => child.type)).toEqual(["paragraph"]);
    expect(warnings).toEqual([
      {
        code: "unexpected-token",
        severity: "error",
        message: 'orphaned "blockquote_close" with no matching opener',
        line: 1,
      },
    ]);
  });
});
