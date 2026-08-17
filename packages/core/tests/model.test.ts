import { describe, expect, it } from "vitest";

import {
  assertNever,
  blockquote,
  bold,
  bulletList,
  childrenOf,
  codeBlock,
  createSlugger,
  doc,
  FALLBACK_SLUG,
  findMark,
  footnoteDefinition,
  footnoteReference,
  hardBreak,
  hasMark,
  headerRow,
  heading,
  highlight,
  htmlBlock,
  htmlInline,
  image,
  inlineCode,
  isBlockNode,
  isInlineNode,
  isStructuralNode,
  italic,
  link,
  list,
  listItem,
  mathBlock,
  mathInline,
  metadata,
  nodeCategory,
  nodesOfType,
  nodeText,
  NODE_TYPES,
  NODE_TYPES_ARE_EXHAUSTIVE,
  normalizeMarks,
  orderedList,
  paragraph,
  slugify,
  softBreak,
  strikethrough,
  subscript,
  superscript,
  table,
  tableCell,
  tableRow,
  taskList,
  text,
  thematicBreak,
  visit,
  type DocumentNode,
  type Node,
  type NodeType,
} from "../src/model.js";

/* -------------------------------------------------------------------------- */
/* A fixture that contains every node type exactly once (or more).             */
/* -------------------------------------------------------------------------- */

/**
 * Builds a document exercising all 21 node kinds. Doubles as the readability
 * demo for the builders and as the input for the traversal tests.
 */
function everythingDocument(): DocumentNode {
  return doc(
    [
      heading(1, [text("Downword "), text("Model", [inlineCode()])]),
      paragraph([
        text("A "),
        text("composed", [bold(), italic(), strikethrough(), link("https://example.com", "t")]),
        text(" run, "),
        text("x", [superscript()]),
        text("y", [subscript()]),
        text("!", [highlight()]),
        hardBreak(),
        text("after the break"),
        softBreak(),
        image("logo.png", { alt: "Logo", marks: [link("https://example.com")] }),
        htmlInline("<kbd>"),
        footnoteReference("1", 1),
        mathInline("x^2"),
      ]),
      blockquote([paragraph([text("quoted")]), blockquote([paragraph([text("deeper")])])]),
      bulletList([
        listItem([
          paragraph([text("outer")]),
          orderedList([listItem([paragraph([text("inner")])])], { start: 3 }),
        ]),
      ]),
      taskList([listItem([paragraph([text("done")])], { checked: true })]),
      table([
        headerRow([tableCell([text("Key")]), tableCell([text("Value")])]),
        tableRow([tableCell([text("a")]), tableCell([text("1")])]),
      ]),
      codeBlock("const a = 1;\n", { lang: "ts" }),
      thematicBreak(),
      htmlBlock("<details><summary>x</summary></details>"),
      mathBlock("\\int_0^1 x"),
      footnoteDefinition("1", 1, [paragraph([text("The note.")])]),
    ],
    metadata({ title: "Everything", keywords: ["a", "b"] }),
  );
}

/* -------------------------------------------------------------------------- */
/* Marks                                                                       */
/* -------------------------------------------------------------------------- */

describe("marks", () => {
  it("composes rather than enumerating combinations", () => {
    const node = text("x", [bold(), italic(), inlineCode(), link("https://example.com")]);

    expect(node.marks.map((mark) => mark.type)).toEqual(["link", "bold", "italic", "inlineCode"]);
  });

  it("normalises to a canonical order so equal formatting is structurally equal", () => {
    expect(text("x", [italic(), bold()])).toEqual(text("x", [bold(), italic()]));
  });

  it("de-duplicates by mark type, last one winning", () => {
    const marks = normalizeMarks([
      link("https://first.example", null),
      bold(),
      link("https://second.example", "second"),
      bold(),
    ]);

    expect(marks).toEqual([
      { type: "link", href: "https://second.example", title: "second" },
      { type: "bold" },
    ]);
  });

  it("returns the shared empty list for no marks", () => {
    expect(normalizeMarks([])).toEqual([]);
    expect(text("x").marks).toBe(normalizeMarks([]));
  });

  it("reads marks back with hasMark and findMark", () => {
    const marks = text("x", [bold(), link("#anchor", "go")]).marks;

    expect(hasMark(marks, "bold")).toBe(true);
    expect(hasMark(marks, "highlight")).toBe(false);

    const found = findMark(marks, "link");
    expect(found?.href).toBe("#anchor");
    expect(found?.title).toBe("go");
    expect(findMark(marks, "italic")).toBeNull();
  });

  it("defaults a link title to null", () => {
    expect(link("https://example.com")).toEqual({
      type: "link",
      href: "https://example.com",
      title: null,
    });
  });

  it("marks every inline node kind, not just text", () => {
    expect(hardBreak([bold()]).marks).toHaveLength(1);
    expect(softBreak([bold()]).marks).toHaveLength(1);
    expect(htmlInline("<br>", [bold()]).marks).toHaveLength(1);
    expect(image("a.png", { marks: [link("u")] }).marks).toHaveLength(1);
    expect(footnoteReference("1", 1, { marks: [bold()] }).marks).toHaveLength(1);
    expect(mathInline("x", { marks: [bold()] }).marks).toHaveLength(1);
  });
});

/* -------------------------------------------------------------------------- */
/* Builders                                                                    */
/* -------------------------------------------------------------------------- */

describe("builders", () => {
  it("produces total nodes: no field is ever absent", () => {
    expect(image("a.png")).toEqual({
      type: "image",
      src: "a.png",
      alt: "",
      title: null,
      resolved: null,
      marks: [],
    });

    expect(codeBlock("x")).toEqual({
      type: "codeBlock",
      lang: null,
      meta: null,
      value: "x",
      highlights: null,
    });

    expect(mathBlock("x")).toEqual({ type: "mathBlock", value: "x", omml: null });

    expect(metadata()).toEqual({
      title: null,
      author: null,
      description: null,
      keywords: [],
      date: null,
      custom: {},
    });
  });

  it("derives a heading id from its text content", () => {
    expect(heading(2, [text("Getting "), text("Started", [bold()])]).id).toBe("getting-started");
  });

  it("accepts an explicit heading id", () => {
    expect(heading(3, [text("Anything")], { id: "custom-id" }).id).toBe("custom-id");
  });

  it("gives ordered lists a start and denies one to bullet/task lists", () => {
    expect(orderedList([]).start).toBe(1);
    expect(orderedList([], { start: 5 }).start).toBe(5);
    expect(bulletList([], { start: 5 }).start).toBeNull();
    expect(taskList([], { start: 5 }).start).toBeNull();
    expect(list("ordered", []).kind).toBe("ordered");
  });

  it("defaults lists to tight and items to non-task", () => {
    expect(bulletList([]).tight).toBe(true);
    expect(bulletList([], { tight: false }).tight).toBe(false);
    expect(listItem([]).checked).toBeNull();
    expect(listItem([], { checked: false }).checked).toBe(false);
    expect(listItem([], { checked: true }).checked).toBe(true);
  });

  it("pads table alignment to the widest row", () => {
    const built = table(
      [
        headerRow([tableCell([]), tableCell([]), tableCell([])]),
        tableRow([tableCell([]), tableCell([])]),
      ],
      { align: ["right"] },
    );

    expect(built.align).toEqual(["right", "none", "none"]);
    expect(built.children[0]?.header).toBe(true);
    expect(built.children[1]?.header).toBe(false);
  });

  it("keeps an over-long alignment array rather than truncating it", () => {
    expect(table([tableRow([tableCell([])])], { align: ["left", "center"] }).align).toEqual([
      "left",
      "center",
    ]);
  });

  it("defaults footnote labels to the identifier", () => {
    expect(footnoteReference("note-a", 2)).toEqual({
      type: "footnoteReference",
      identifier: "note-a",
      label: "note-a",
      number: 2,
      marks: [],
    });

    expect(footnoteDefinition("note-a", 2, [], { label: "a" }).label).toBe("a");
  });

  it("nests to arbitrary depth", () => {
    const deep = blockquote([
      bulletList([
        listItem([
          bulletList([
            listItem([blockquote([paragraph([text("bottom", [bold(), link("#top")])])])]),
          ]),
        ]),
      ]),
    ]);

    expect(nodeText(deep)).toBe("bottom");
    expect(nodesOfType(deep, "list")).toHaveLength(2);
    expect(nodesOfType(deep, "blockquote")).toHaveLength(2);
  });
});

/* -------------------------------------------------------------------------- */
/* Slugs                                                                       */
/* -------------------------------------------------------------------------- */

describe("slugify", () => {
  it("lowercases, drops punctuation and hyphenates whitespace", () => {
    expect(slugify("Hello, World!")).toBe("hello-world");
    expect(slugify("  API   reference  ")).toBe("api-reference");
    expect(slugify("snake_case-and-dash")).toBe("snake_case-and-dash");
  });

  it("keeps non-latin letters and digits", () => {
    expect(slugify("Café 42 日本語")).toBe("café-42-日本語");
  });

  it("falls back when nothing survives", () => {
    expect(slugify("!!!")).toBe(FALLBACK_SLUG);
    expect(slugify("")).toBe(FALLBACK_SLUG);
  });
});

describe("createSlugger", () => {
  it("suffixes collisions the way GitHub does", () => {
    const slugger = createSlugger();

    expect(slugger.slug("Setup")).toBe("setup");
    expect(slugger.slug("Setup")).toBe("setup-1");
    expect(slugger.slug("setup")).toBe("setup-2");
    expect(slugger.slug("Other")).toBe("other");
  });

  it("does not hand out a suffix that was already claimed literally", () => {
    const slugger = createSlugger();

    expect(slugger.slug("a-1")).toBe("a-1");
    expect(slugger.slug("a")).toBe("a");
    expect(slugger.slug("a")).toBe("a-2");
  });

  it("resets", () => {
    const slugger = createSlugger();

    expect(slugger.slug("x")).toBe("x");
    slugger.reset();
    expect(slugger.slug("x")).toBe("x");
  });
});

/* -------------------------------------------------------------------------- */
/* Traversal                                                                   */
/* -------------------------------------------------------------------------- */

describe("childrenOf", () => {
  it("returns children for containers", () => {
    const node = paragraph([text("a"), text("b")]);
    expect(childrenOf(node)).toHaveLength(2);
  });

  it("returns one shared empty array for every leaf", () => {
    expect(childrenOf(text("a"))).toEqual([]);
    expect(childrenOf(text("a"))).toBe(childrenOf(thematicBreak()));
  });

  it("handles every node type without throwing", () => {
    const seen = new Set<NodeType>();
    visit(everythingDocument(), (node) => {
      seen.add(node.type);
      expect(() => childrenOf(node)).not.toThrow();
    });

    expect([...seen].sort()).toEqual([...NODE_TYPES].sort());
  });
});

describe("nodeCategory", () => {
  it("classifies every node type", () => {
    expect(nodeCategory(doc([]))).toBe("document");
    expect(nodeCategory(paragraph([]))).toBe("block");
    expect(nodeCategory(listItem([]))).toBe("structural");
    expect(nodeCategory(text("x"))).toBe("inline");
  });

  it("backs the type guards", () => {
    const nodes: readonly Node[] = [doc([]), paragraph([]), tableRow([]), text("x")];

    expect(nodes.filter(isBlockNode)).toHaveLength(1);
    expect(nodes.filter(isStructuralNode)).toHaveLength(1);
    expect(nodes.filter(isInlineNode)).toHaveLength(1);
  });
});

describe("visit", () => {
  it("walks depth-first in pre-order", () => {
    const tree = doc([heading(1, [text("H")]), paragraph([text("a"), text("b")])]);
    const order: string[] = [];

    visit(tree, (node) => {
      order.push(node.type === "text" ? `text:${node.value}` : node.type);
    });

    expect(order).toEqual(["document", "heading", "text:H", "paragraph", "text:a", "text:b"]);
  });

  it("reports parent, index and ancestors", () => {
    const inner = text("b");
    const para = paragraph([text("a"), inner]);
    const tree = doc([para]);

    const seen: { parent: string | null; index: number; ancestors: string[] }[] = [];

    visit(tree, (node, context) => {
      if (node === inner) {
        seen.push({
          parent: context.parent?.type ?? null,
          index: context.index,
          ancestors: context.ancestors.map((ancestor) => ancestor.type),
        });
      }
    });

    expect(seen).toEqual([
      {
        parent: "paragraph",
        index: 1,
        ancestors: ["document", "paragraph"],
      },
    ]);
  });

  it("gives the root a null parent and index -1", () => {
    const tree = doc([]);
    const contexts: number[] = [];

    visit(tree, (node, context) => {
      expect(node).toBe(tree);
      expect(context.parent).toBeNull();
      expect(context.ancestors).toEqual([]);
      contexts.push(context.index);
    });

    expect(contexts).toEqual([-1]);
  });

  it("prunes a subtree when enter returns 'skip', but still leaves the node", () => {
    const tree = doc([paragraph([text("hidden")]), paragraph([text("shown")])]);
    const entered: string[] = [];
    const left: string[] = [];

    visit(
      tree,
      (node, context) => {
        entered.push(node.type);
        return node.type === "paragraph" && context.index === 0 ? "skip" : undefined;
      },
      (node) => {
        left.push(node.type);
      },
    );

    expect(entered).toEqual(["document", "paragraph", "paragraph", "text"]);
    expect(left).toEqual(["paragraph", "text", "paragraph", "document"]);
  });

  it("calls leave in post-order", () => {
    const tree = doc([paragraph([text("a")])]);
    const left: string[] = [];

    visit(
      tree,
      () => undefined,
      (node) => {
        left.push(node.type);
      },
    );

    expect(left).toEqual(["text", "paragraph", "document"]);
  });

  it("hands out a fresh ancestor snapshot per node", () => {
    const snapshots: (readonly Node[])[] = [];
    const tree = doc([paragraph([text("a")])]);

    visit(tree, (_node, context) => {
      snapshots.push(context.ancestors);
    });

    expect(snapshots.map((snapshot) => snapshot.length)).toEqual([0, 1, 2]);
  });
});

describe("nodesOfType", () => {
  it("collects in document order, root included", () => {
    const tree = everythingDocument();

    expect(nodesOfType(tree, "document")).toHaveLength(1);
    expect(nodesOfType(tree, "heading").map((node) => node.id)).toEqual(["downword-model"]);
    expect(nodesOfType(tree, "tableCell")).toHaveLength(4);
    expect(nodesOfType(tree, "footnoteDefinition").map((node) => node.number)).toEqual([1]);
  });

  it("returns an empty list when nothing matches", () => {
    expect(nodesOfType(paragraph([text("x")]), "table")).toEqual([]);
  });
});

describe("nodeText", () => {
  it("concatenates text, code and math, and turns breaks into newlines", () => {
    const node = paragraph([
      text("a"),
      hardBreak(),
      text("b"),
      softBreak(),
      mathInline("x^2"),
      htmlInline("<br>"),
      image("i.png", { alt: "ignored" }),
    ]);

    expect(nodeText(node)).toBe("a\nb\nx^2");
  });

  it("reads through a whole document", () => {
    expect(nodeText(doc([heading(1, [text("T")]), paragraph([text("body")])]))).toBe("Tbody");
  });
});

/* -------------------------------------------------------------------------- */
/* Exhaustiveness                                                              */
/* -------------------------------------------------------------------------- */

describe("exhaustiveness", () => {
  it("lists every node type exactly once in NODE_TYPES", () => {
    expect(NODE_TYPES).toHaveLength(21);
    expect(new Set(NODE_TYPES).size).toBe(NODE_TYPES.length);
    expect(NODE_TYPES_ARE_EXHAUSTIVE).toBe(true);
  });

  it("throws from assertNever when a switch is reached with an impossible value", () => {
    const impossible = { type: "callout" } as unknown as never;

    expect(() => assertNever(impossible, "node")).toThrow(
      /downword: unhandled node: \{"type":"callout"\}/,
    );
    expect(() => childrenOf(impossible)).toThrow(/unhandled node/);
    expect(() => nodeCategory(impossible)).toThrow(/unhandled node/);
  });

  it("defaults the assertNever context", () => {
    expect(() => assertNever(1 as unknown as never)).toThrow("downword: unhandled value: 1");
  });
});
