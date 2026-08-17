import { describe, expect, it } from "vitest";

import {
  AUTO_DETECT_LANGUAGES,
  createHighlighter,
  LANGUAGE_ALIASES,
  LANGUAGE_DEPENDENCIES,
  LANGUAGE_LOADERS,
  normalizeLanguageId,
  PLAIN_LANGUAGES,
  PRINT_CODE_PALETTE,
  PRINT_INK,
  PRINT_SCOPE_STYLES,
  resolveLanguageId,
  resolveScopeStyle,
  SpanEmitter,
  toHighlightSpans,
  type HighlighterOptions,
  type HighlightWarning,
  type LanguageLoader,
} from "../src/highlight/index.js";
import { THEMES } from "../src/index.js";
import { codeBlock, doc } from "../src/model.js";
import {
  DEFAULT_THEME,
  prepareHighlights,
  type HighlightSpan,
  type Theme,
} from "../src/render/index.js";
import { contrast } from "./helpers/contrast.js";
import { renderParts } from "./helpers/docx.js";
import { paragraphs, pStyle, textOf } from "./helpers/xml.js";

/* -------------------------------------------------------------------------- */
/* Fixtures                                                                    */
/* -------------------------------------------------------------------------- */

const TS = `const answer: number = 42; // the answer\n`;
const PY = `# greet\ndef greet(name: str) -> str:\n    return f"hello {name}"\n`;
const SQL = `SELECT id, name FROM users WHERE id = 1;\n`;
const BASH = `#!/usr/bin/env bash\necho "hello" # noise\n`;
const HTML = `<div id="app"><script>const n = 1;</script></div>\n`;

/** The renderer's CodeBlock shading: the real background these inks sit on. */
const CODE_BACKGROUND = DEFAULT_THEME.colors.codeBackground;

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

/** Collects the warnings a highlighter raises, alongside the highlighter. */
function withWarnings(options: Omit<HighlighterOptions, "onWarning"> = {}): {
  highlight: (code: string, lang: string | null) => Promise<readonly HighlightSpan[]>;
  warnings: HighlightWarning[];
} {
  const warnings: HighlightWarning[] = [];
  const highlighter = createHighlighter({ ...options, onWarning: (w) => warnings.push(w) });
  return {
    highlight: async (code, lang) => highlighter.highlight(code, lang),
    warnings,
  };
}

/** Looks a span up by its exact text. Fails loudly, with the whole span list. */
function span(spans: readonly HighlightSpan[], text: string): HighlightSpan {
  const found = spans.find((candidate) => candidate.text === text);
  if (found === undefined) {
    throw new Error(
      `no span with text ${JSON.stringify(text)}; saw ${JSON.stringify(spans.map((s) => s.text))}`,
    );
  }
  return found;
}

/** Concatenated span text: must always reproduce the highlighter's input. */
function joined(spans: readonly HighlightSpan[]): string {
  return spans.map((s) => s.text).join("");
}

/** How many spans carry exactly the formatting of the span before them. */
function adjacentDuplicates(spans: readonly HighlightSpan[]): number {
  let count = 0;
  for (let index = 1; index < spans.length; index += 1) {
    const current = spans[index];
    const previous = spans[index - 1];
    if (current === undefined || previous === undefined) continue;
    if (
      current.color === previous.color &&
      current.bold === previous.bold &&
      current.italic === previous.italic &&
      // Part of the identity: two scopes sharing an ink here may not share one
      // once the renderer's theme has had its say.
      current.scope === previous.scope
    ) {
      count += 1;
    }
  }
  return count;
}

/* --- OOXML run probing ---------------------------------------------------- */

/** One `<w:r>` reduced to the four things highlighting can change. */
interface RunXml {
  readonly text: string;
  readonly color: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
}

function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/** Every run of an element, with its direct character formatting. */
function runs(xml: string): RunXml[] {
  return [...xml.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((match) => {
    const body = match[1] ?? "";
    const props = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(body)?.[1] ?? "";
    return {
      text: decodeXml(
        [...body.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((t) => t[1] ?? "").join(""),
      ),
      color: /<w:color w:val="([^"]*)"\/>/.exec(props)?.[1] ?? null,
      bold: props.includes("<w:b/>"),
      italic: props.includes("<w:i/>"),
    };
  });
}

/** Every run of every `CodeBlock`-styled paragraph, in document order. */
function codeRuns(documentXml: string): RunXml[] {
  return paragraphs(documentXml)
    .filter((p) => pStyle(p) === "CodeBlock")
    .flatMap(runs);
}

function runWithText(all: readonly RunXml[], text: string): RunXml {
  const found = all.find((run) => run.text === text);
  if (found === undefined) {
    throw new Error(
      `no run with text ${JSON.stringify(text)}; saw ${JSON.stringify(all.map((r) => r.text))}`,
    );
  }
  return found;
}

/* -------------------------------------------------------------------------- */
/* Info strings                                                                */
/* -------------------------------------------------------------------------- */

describe("normalizeLanguageId", () => {
  it.each([
    ["ts", "ts"],
    ["  TS  ", "ts"],
    ["TypeScript", "typescript"],
    ["ts title=example.ts", "ts"],
    ["python {highlight: [1,2]}", "python"],
    ["js{1,3-5}", "js"],
    ["{.python .numberLines}", "python"],
    ["{ .sql }", "sql"],
    ["language-python", "python"],
    ["lang-rb", "rb"],
    ["highlight-source-js", "js"],
    ["json,foo", "json"],
    ["yaml:config", "yaml"],
    ["c++", "c++"],
    ["f#", "f#"],
    ["obj-c", "obj-c"],
    ["cmake.in", "cmake.in"],
    ["ts.", "ts"],
  ])("normalizes %o to %o", (info, expected) => {
    expect(normalizeLanguageId(info)).toBe(expected);
  });

  it.each([null, undefined, "", "   ", "{}", "=nope", "  {  }  "])(
    "returns null for %o",
    (info) => {
      expect(normalizeLanguageId(info)).toBeNull();
    },
  );

  it("does not strip a prefix that is the whole id", () => {
    // `lang-` alone is a language nobody has, but truncating it to "" would
    // turn a bad info string into an unlabelled fence instead of an unknown one.
    expect(normalizeLanguageId("lang-")).toBe("lang-");
  });
});

/* -------------------------------------------------------------------------- */
/* Registry                                                                    */
/* -------------------------------------------------------------------------- */

describe("language registry", () => {
  it.each([
    ["ts", "typescript"],
    ["tsx", "typescript"],
    ["py", "python"],
    ["sh", "bash"],
    ["zsh", "bash"],
    ["console", "shell"],
    ["yml", "yaml"],
    ["html", "xml"],
    ["toml", "ini"],
    ["c++", "cpp"],
    ["c#", "csharp"],
    ["objective-c", "objectivec"],
    ["golang", "go"],
    ["typescript", "typescript"],
  ])("resolves %o to the %o grammar", (id, expected) => {
    const resolution = resolveLanguageId(id, LANGUAGE_LOADERS, LANGUAGE_ALIASES);
    expect(resolution).toMatchObject({ kind: "language", name: expected });
  });

  it.each([...PLAIN_LANGUAGES])("treats %o as an explicit request for plain text", (id) => {
    expect(resolveLanguageId(id, LANGUAGE_LOADERS, LANGUAGE_ALIASES).kind).toBe("plain");
  });

  it.each(["klingon", "brainfuck", "cobol"])("reports %o as unknown", (id) => {
    expect(resolveLanguageId(id, LANGUAGE_LOADERS, LANGUAGE_ALIASES).kind).toBe("unknown");
  });

  it("points every alias at a grammar that exists", () => {
    const dangling = Object.entries(LANGUAGE_ALIASES).filter(
      ([, canonical]) => LANGUAGE_LOADERS[canonical] === undefined,
    );
    expect(dangling).toEqual([]);
  });

  it("never lets an alias shadow a canonical name", () => {
    const shadowed = Object.keys(LANGUAGE_ALIASES).filter(
      (alias) => LANGUAGE_LOADERS[alias] !== undefined,
    );
    expect(shadowed).toEqual([]);
  });

  it("never aliases a plain-text id", () => {
    const clashes = [...PLAIN_LANGUAGES].filter(
      (id) => LANGUAGE_ALIASES[id] !== undefined || LANGUAGE_LOADERS[id] !== undefined,
    );
    expect(clashes).toEqual([]);
  });

  it("keeps every auto-detect candidate and dependency loadable", () => {
    const referenced = [...AUTO_DETECT_LANGUAGES, ...Object.values(LANGUAGE_DEPENDENCIES).flat()];
    const missing = referenced.filter((name) => LANGUAGE_LOADERS[name] === undefined);
    expect(missing).toEqual([]);
  });

  it("normalises every id it can then resolve", () => {
    // A table key that does not survive normalisation is unreachable.
    const unreachable = [...Object.keys(LANGUAGE_LOADERS), ...Object.keys(LANGUAGE_ALIASES)].filter(
      (id) => normalizeLanguageId(id) !== id,
    );
    expect(unreachable).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Palette                                                                     */
/* -------------------------------------------------------------------------- */

describe("print palette", () => {
  // Pinned so a palette edit has to restate the contrast it is trading away.
  // Both backgrounds matter: F6F8FA is the CodeBlock shading, FFFFFF is what a
  // reader gets if the shading is restyled away or the printer drops it.
  it.each([
    ["text", PRINT_INK.text, 13.76, 14.65],
    ["keyword", PRINT_INK.keyword, 7.39, 7.87],
    ["string", PRINT_INK.string, 12.03, 12.81],
    ["comment", PRINT_INK.comment, 7.32, 7.79],
    ["number", PRINT_INK.number, 7.13, 7.59],
    ["type", PRINT_INK.type, 7.67, 8.17],
    ["callable", PRINT_INK.callable, 8.15, 8.68],
    ["markup", PRINT_INK.markup, 8.52, 9.07],
    ["meta", PRINT_INK.meta, 9.19, 9.79],
    ["removed", PRINT_INK.removed, 9.87, 10.51],
  ])(
    "%s (#%s) contrasts %f:1 on the code shading and %f:1 on white",
    (_role, ink, shaded, white) => {
      expect(contrast(ink, CODE_BACKGROUND)).toBeCloseTo(shaded, 2);
      expect(contrast(ink, "FFFFFF")).toBeCloseTo(white, 2);
    },
  );

  it("clears WCAG AAA (7:1) for every scope, on shading and on white", () => {
    const failures = Object.entries(PRINT_SCOPE_STYLES)
      .flatMap(([scope, style]) => (style.color === undefined ? [] : [[scope, style.color]]))
      .map(([scope, color]) => ({
        scope,
        onShading: contrast(color ?? "", CODE_BACKGROUND),
        onWhite: contrast(color ?? "", "FFFFFF"),
      }))
      .filter((row) => row.onShading < 7 || row.onWhite < 7);

    expect(failures).toEqual([]);
  });

  it("writes colours the way docx wants them: bare uppercase RRGGBB", () => {
    const malformed = Object.values(PRINT_SCOPE_STYLES)
      .map((style) => style.color)
      .filter((color) => color !== undefined && !/^[0-9A-F]{6}$/.test(color));

    expect(malformed).toEqual([]);
  });

  it("agrees with the renderer on the default code colour", () => {
    // PRINT_SCOPE_STYLES leaves ordinary text uncoloured on purpose; that is
    // only correct while the CodeBlock style already paints it this colour.
    expect(PRINT_INK.text).toBe(DEFAULT_THEME.colors.codeText);
  });

  it("carries a non-colour signal on the two scopes greyscale cannot separate", () => {
    // keyword (7.39) and comment (7.32) are the same grey once printed mono.
    expect(contrast(PRINT_INK.keyword, PRINT_INK.comment)).toBeLessThan(1.2);
    expect(PRINT_SCOPE_STYLES["keyword"]).toMatchObject({ bold: true });
    expect(PRINT_SCOPE_STYLES["comment"]).toMatchObject({ italic: true });
  });

  it("resolves dotted scopes longest-prefix-first", () => {
    expect(resolveScopeStyle(PRINT_SCOPE_STYLES, "title.function")).toBe(
      PRINT_SCOPE_STYLES["title.function"],
    );
    // Invented leaf, known family: inherits rather than rendering plain.
    expect(resolveScopeStyle(PRINT_SCOPE_STYLES, "title.function.invented")).toBe(
      PRINT_SCOPE_STYLES["title.function"],
    );
    expect(resolveScopeStyle(PRINT_SCOPE_STYLES, "meta.something")).toBe(
      PRINT_SCOPE_STYLES["meta"],
    );
    expect(resolveScopeStyle(PRINT_SCOPE_STYLES, "nonsense")).toBeNull();
  });

  it("exports a colour-only view usable as the renderer's theme palette", () => {
    expect(PRINT_CODE_PALETTE["keyword"]).toBe(PRINT_INK.keyword);
    expect(PRINT_CODE_PALETTE["comment"]).toBe(PRINT_INK.comment);
    // Style-only entries carry no colour and must not appear.
    expect(PRINT_CODE_PALETTE["strong"]).toBeUndefined();
    expect(Object.values(PRINT_CODE_PALETTE).every((c) => /^[0-9A-F]{6}$/.test(c))).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* Span conversion                                                             */
/* -------------------------------------------------------------------------- */

describe("span conversion", () => {
  it("merges adjacent spans that resolve to the same formatting", () => {
    const spans = toHighlightSpans(
      [
        // Different scopes, same ink: still one run.
        { scope: "attr", text: "a" },
        { scope: "number", text: "1" },
        { scope: null, text: " " },
        { scope: null, text: "b" },
        { scope: "keyword", text: "if" },
      ],
      "a1 bif",
      PRINT_SCOPE_STYLES,
    );

    expect(spans.map((s) => s.text)).toEqual(["a1", " b", "if"]);
    expect(spans[0]?.color).toBe(PRINT_INK.number);
    expect(spans[1]?.color).toBeUndefined();
    expect(spans[2]).toMatchObject({ color: PRINT_INK.keyword, bold: true });
  });

  it("appends the tail when a grammar stops early", () => {
    // What an `illegal` match leaves behind: a correct prefix and nothing else.
    const spans = toHighlightSpans(
      [{ scope: "keyword", text: "if" }],
      "if (x) {",
      PRINT_SCOPE_STYLES,
    );

    expect(joined(spans)).toBe("if (x) {");
    expect(spans[0]).toMatchObject({ color: PRINT_INK.keyword });
    expect(spans[1]).toMatchObject({ text: " (x) {", color: undefined });
  });

  it("discards a token stream that does not match the source at all", () => {
    const spans = toHighlightSpans(
      [{ scope: "keyword", text: "totally wrong" }],
      "if",
      PRINT_SCOPE_STYLES,
    );

    expect(spans).toEqual([{ text: "if", color: undefined, bold: undefined, italic: undefined }]);
  });

  it("emits nothing for empty input", () => {
    expect(toHighlightSpans([], "", PRINT_SCOPE_STYLES)).toEqual([]);
  });
});

describe("SpanEmitter", () => {
  it("collapses nested scopes to the innermost one", () => {
    const emitter = new SpanEmitter({});
    emitter.openNode("function");
    emitter.addText("def ");
    emitter.openNode("title.function");
    emitter.addText("f");
    emitter.closeNode();
    emitter.addText("()");
    emitter.closeNode();
    emitter.finalize();

    expect(emitter.spans).toEqual([
      { scope: "function", text: "def " },
      { scope: "title.function", text: "f" },
      { scope: "function", text: "()" },
    ]);
  });

  it("gives a sub-language's unscoped text the outer scope", () => {
    const outer = new SpanEmitter({});
    const inner = new SpanEmitter({});
    inner.addText("var ");
    inner.startScope("number");
    inner.addText("1");
    inner.endScope();

    outer.openNode("tag");
    outer.__addSublanguage(inner, "javascript");
    outer.closeNode();

    expect(outer.spans).toEqual([
      { scope: "tag", text: "var " },
      { scope: "number", text: "1" },
    ]);
  });

  it("produces no HTML at all", () => {
    const emitter = new SpanEmitter({});
    emitter.addText("<script>");
    expect(emitter.toHTML()).toBe("");
    // The text is kept verbatim - never entity-encoded.
    expect(emitter.spans).toEqual([{ scope: null, text: "<script>" }]);
  });
});

/* -------------------------------------------------------------------------- */
/* The adapter                                                                 */
/* -------------------------------------------------------------------------- */

describe("createHighlighter", () => {
  it("colours TypeScript", async () => {
    const spans = await createHighlighter().highlight(TS, "ts");

    expect(joined(spans)).toBe(TS);
    expect(span(spans, "const")).toMatchObject({ color: PRINT_INK.keyword, bold: true });
    expect(span(spans, "number")).toMatchObject({ color: PRINT_INK.type });
    expect(span(spans, "42")).toMatchObject({ color: PRINT_INK.number });
    expect(span(spans, "// the answer")).toMatchObject({
      color: PRINT_INK.comment,
      italic: true,
    });
  });

  it("colours Python", async () => {
    const spans = await createHighlighter().highlight(PY, "py");

    expect(joined(spans)).toBe(PY);
    expect(span(spans, "# greet")).toMatchObject({ color: PRINT_INK.comment, italic: true });
    expect(span(spans, "def")).toMatchObject({ color: PRINT_INK.keyword, bold: true });
    expect(span(spans, "greet")).toMatchObject({ color: PRINT_INK.callable });
    expect(span(spans, 'f"hello ')).toMatchObject({ color: PRINT_INK.string });
  });

  it("colours SQL", async () => {
    const spans = await createHighlighter().highlight(SQL, "sql");

    expect(joined(spans)).toBe(SQL);
    for (const keyword of ["SELECT", "FROM", "WHERE"]) {
      expect(span(spans, keyword)).toMatchObject({ color: PRINT_INK.keyword, bold: true });
    }
    expect(span(spans, "1")).toMatchObject({ color: PRINT_INK.number });
  });

  it("colours Bash", async () => {
    const spans = await createHighlighter().highlight(BASH, "bash");

    expect(joined(spans)).toBe(BASH);
    expect(span(spans, "#!/usr/bin/env bash")).toMatchObject({ color: PRINT_INK.meta });
    expect(span(spans, "echo")).toMatchObject({ color: PRINT_INK.type });
    expect(span(spans, '"hello"')).toMatchObject({ color: PRINT_INK.string });
    expect(span(spans, "# noise")).toMatchObject({ color: PRINT_INK.comment, italic: true });
  });

  it("resolves the language through the alias table", async () => {
    const viaAlias = await createHighlighter().highlight(TS, "ts");
    const viaName = await createHighlighter().highlight(TS, "typescript");
    expect(viaAlias).toEqual(viaName);
  });

  it("ignores extra info-string attributes", async () => {
    const plain = await createHighlighter().highlight(TS, "ts");
    const decorated = await createHighlighter().highlight(TS, "ts title=answer.ts showLineNumbers");
    expect(decorated).toEqual(plain);
  });

  it("colours code embedded in another language", async () => {
    const spans = await createHighlighter().highlight(HTML, "html");

    expect(joined(spans)).toBe(HTML);
    // From the xml grammar...
    expect(span(spans, "div")).toMatchObject({ color: PRINT_INK.markup });
    // ...and from the javascript grammar it hands `<script>` bodies to, which
    // only works because LANGUAGE_DEPENDENCIES registered it too.
    expect(span(spans, "const")).toMatchObject({ color: PRINT_INK.keyword, bold: true });
  });

  it.each([
    ["ts", TS],
    ["py", PY],
    ["sql", SQL],
    ["bash", BASH],
    ["html", HTML],
  ])("emits no two consecutive %s runs with identical formatting", async (lang, code) => {
    const spans = await createHighlighter().highlight(code, lang);

    expect(joined(spans)).toBe(code);
    expect(adjacentDuplicates(spans)).toBe(0);
  });

  it("can be told not to load a grammar's embedded languages", async () => {
    const spans = await createHighlighter({ dependencies: {} }).highlight(HTML, "html");

    expect(joined(spans)).toBe(HTML);
    expect(span(spans, "div")).toMatchObject({ color: PRINT_INK.markup });
    // The <script> body is now one unhighlighted lump.
    expect(spans.some((s) => s.text.includes("const n = 1;"))).toBe(true);
  });

  it("honours a replacement scope palette", async () => {
    const spans = await createHighlighter({
      scopeStyles: { keyword: { color: "112233" } },
    }).highlight(TS, "ts");

    expect(joined(spans)).toBe(TS);
    expect(span(spans, "const")).toMatchObject({ color: "112233", bold: undefined });
    // Replacement, not merge: every other scope now resolves to nothing, so the
    // rest of the line coalesces into a single unformatted run.
    expect(spans).toHaveLength(2);
    expect(spans[1]?.color).toBeUndefined();
  });

  it("honours a narrowed language registry", async () => {
    const { highlight, warnings } = withWarnings({
      languages: { typescript: LANGUAGE_LOADERS["typescript"] as LanguageLoader },
    });

    expect(span(await highlight(TS, "ts"), "const").color).toBe(PRINT_INK.keyword);
    expect(await highlight(PY, "py")).toEqual([
      { text: PY, color: undefined, bold: undefined, italic: undefined },
    ]);
    expect(warnings.map((w) => w.code)).toEqual(["language-unknown"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Falling back                                                                */
/* -------------------------------------------------------------------------- */

describe("fallbacks", () => {
  const plain = (code: string): readonly HighlightSpan[] => [
    { text: code, color: undefined, bold: undefined, italic: undefined },
  ];

  it("falls back to plain monospace for an unknown language, and says so", async () => {
    const { highlight, warnings } = withWarnings();
    expect(await highlight(TS, "klingon")).toEqual(plain(TS));
    expect(warnings).toEqual([
      { code: "language-unknown", message: expect.stringContaining("klingon"), lang: "klingon" },
    ]);
  });

  it("stays silent for an unlabelled fence", async () => {
    const { highlight, warnings } = withWarnings();
    expect(await highlight(TS, null)).toEqual(plain(TS));
    expect(warnings).toEqual([]);
  });

  it("stays silent - and loads nothing - for an explicit plain-text fence", async () => {
    let loads = 0;
    const { highlight, warnings } = withWarnings({
      load: async () => {
        loads += 1;
        return (await import("highlight.js/lib/core")).default;
      },
    });

    for (const id of ["text", "plaintext", "none", "mermaid"]) {
      expect(await highlight(TS, id)).toEqual(plain(TS));
    }
    expect(warnings).toEqual([]);
    expect(loads).toBe(0);
  });

  it("returns an empty span list for empty code", async () => {
    expect(await createHighlighter().highlight("", "ts")).toEqual([]);
  });

  it("survives an unavailable engine and warns exactly once", async () => {
    const { highlight, warnings } = withWarnings({
      load: () => Promise.reject(new Error("no highlight.js here")),
    });

    expect(await highlight(TS, "ts")).toEqual(plain(TS));
    expect(await highlight(PY, "py")).toEqual(plain(PY));
    expect(warnings.map((w) => w.code)).toEqual(["engine-unavailable"]);
  });

  it("survives a grammar that fails to load", async () => {
    const { highlight, warnings } = withWarnings({
      languages: { typescript: () => Promise.reject(new Error("chunk 404")) },
    });

    expect(await highlight(TS, "ts")).toEqual(plain(TS));
    expect(warnings).toEqual([
      { code: "language-load-failed", message: expect.stringContaining("typescript"), lang: "ts" },
    ]);
  });

  it("survives a grammar module with no usable export", async () => {
    const { highlight, warnings } = withWarnings({
      languages: { typescript: () => Promise.resolve({ default: { default: 42 } as never }) },
    });

    expect(await highlight(TS, "ts")).toEqual(plain(TS));
    expect(warnings.map((w) => w.code)).toEqual(["language-load-failed"]);
  });

  it("refuses a fence larger than maxLength", async () => {
    const { highlight, warnings } = withWarnings({ maxLength: 10 });

    expect(await highlight(TS, "ts")).toEqual(plain(TS));
    expect(warnings).toEqual([
      { code: "code-too-large", message: expect.stringContaining("maxLength"), lang: "ts" },
    ]);
  });

  it("never throws, whatever the fence says", async () => {
    const highlighter = createHighlighter();
    const nonsense = [null, "", "   ", "{", " ", "language-", "…", "x".repeat(500)];

    for (const lang of nonsense) {
      const spans = await highlighter.highlight("body\n", lang);
      expect(joined(spans)).toBe("body\n");
    }
  });
});

/* -------------------------------------------------------------------------- */
/* Laziness                                                                    */
/* -------------------------------------------------------------------------- */

describe("laziness", () => {
  it("loads the engine once and each grammar once", async () => {
    let engineLoads = 0;
    const grammarLoads: string[] = [];
    const wrap = (name: string): LanguageLoader => {
      const loader = LANGUAGE_LOADERS[name];
      if (loader === undefined) throw new Error(`no loader for ${name}`);
      return () => {
        grammarLoads.push(name);
        return loader();
      };
    };

    const highlighter = createHighlighter({
      load: async () => {
        engineLoads += 1;
        return (await import("highlight.js/lib/core")).default;
      },
      languages: { typescript: wrap("typescript"), python: wrap("python") },
      dependencies: {},
    });

    await highlighter.highlight(TS, "ts");
    await highlighter.highlight(TS, "typescript");
    await highlighter.highlight(PY, "py");
    await highlighter.highlight(PY, "python");

    expect(engineLoads).toBe(1);
    expect(grammarLoads).toEqual(["typescript", "python"]);
  });

  it("loads a grammar once even when two fences race", async () => {
    let grammarLoads = 0;
    const loader = LANGUAGE_LOADERS["typescript"] as LanguageLoader;
    const highlighter = createHighlighter({
      languages: {
        typescript: () => {
          grammarLoads += 1;
          return loader();
        },
      },
      dependencies: {},
    });

    const [a, b] = await Promise.all([
      highlighter.highlight(TS, "ts"),
      highlighter.highlight(TS, "ts"),
    ]);

    expect(grammarLoads).toBe(1);
    expect(a).toEqual(b);
  });

  it("pulls in a grammar's embedded languages, and only those", async () => {
    const loaded: string[] = [];
    const languages: Record<string, LanguageLoader> = {};
    for (const name of ["xml", "javascript", "css", "python"]) {
      const loader = LANGUAGE_LOADERS[name] as LanguageLoader;
      languages[name] = () => {
        loaded.push(name);
        return loader();
      };
    }

    await createHighlighter({ languages }).highlight(HTML, "html");

    expect(loaded.sort()).toEqual(["css", "javascript", "xml"]);
  });
});

/* -------------------------------------------------------------------------- */
/* Auto-detection                                                              */
/* -------------------------------------------------------------------------- */

describe("auto-detection", () => {
  it("is off by default: an unlabelled fence stays plain", async () => {
    let engineLoads = 0;
    const { highlight } = withWarnings({
      load: async () => {
        engineLoads += 1;
        return (await import("highlight.js/lib/core")).default;
      },
    });

    expect(joined(await highlight(PY, null))).toBe(PY);
    expect(await highlight(PY, null)).toHaveLength(1);
    expect(engineLoads).toBe(0);
  });

  it("guesses when switched on", async () => {
    const spans = await createHighlighter({ autoDetect: ["python", "sql"] }).highlight(PY, null);

    expect(joined(spans)).toBe(PY);
    expect(span(spans, "def")).toMatchObject({ color: PRINT_INK.keyword, bold: true });
  });

  it("also rescues a fence whose named language is unsupported", async () => {
    const { highlight, warnings } = withWarnings({ autoDetect: ["python"] });

    expect(span(await highlight(PY, "python3"), "def").color).toBe(PRINT_INK.keyword);
    expect(warnings).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Scale                                                                       */
/* -------------------------------------------------------------------------- */

describe("large inputs", () => {
  it("highlights a 30 000-line fence without losing a character", { timeout: 60_000 }, async () => {
    const line = `const value = compute(1, "two"); // note\n`;
    const lines = 30_000;
    const code = line.repeat(lines);

    const started = Date.now();
    const spans = await createHighlighter().highlight(code, "ts");
    const elapsed = Date.now() - started;

    expect(joined(spans)).toBe(code);
    expect(joined(spans).split("\n")).toHaveLength(lines + 1);
    // Coalescing has to hold at scale too: every span becomes a <w:r> in
    // document.xml, so a merge missed here is a million redundant elements.
    expect(adjacentDuplicates(spans)).toBe(0);
    expect(span(spans, "const")).toMatchObject({ color: PRINT_INK.keyword, bold: true });
    expect(elapsed).toBeLessThan(20_000);
  });

  it("handles a single line of a million characters", async () => {
    const code = `const s = "${"x".repeat(1_000_000)}";\n`;
    const spans = await createHighlighter().highlight(code, "ts");

    expect(joined(spans)).toBe(code);
  });
});

/* -------------------------------------------------------------------------- */
/* End to end: real .docx bytes                                                */
/* -------------------------------------------------------------------------- */

describe("rendered .docx", () => {
  /**
   * Renders five fences under a chosen theme.
   *
   * The theme is a parameter because it is the theme, not the adapter, that
   * chooses the inks: `createHighlighter()` reports scope names and
   * `theme.codePalette` turns them into `<w:color>`.
   */
  async function renderFences(
    theme: Theme = THEMES.print,
  ): Promise<{ document: string; text: string }> {
    const blocks = [
      codeBlock(TS, { lang: "ts" }),
      codeBlock(PY, { lang: "python title=greet.py" }),
      codeBlock(SQL, { lang: "sql" }),
      codeBlock(BASH, { lang: "sh" }),
      codeBlock(TS, { lang: "klingon" }),
    ];
    const model = doc(blocks);
    const highlights = await prepareHighlights(model, createHighlighter());
    const parts = await renderParts(model, { theme, highlights });

    return { document: parts.document, text: textOf(parts.document) };
  }

  it("writes real w:color values for ts, py, sql and bash fences", async () => {
    const { document } = await renderFences();
    const all = codeRuns(document);

    // TypeScript
    expect(runWithText(all, "const")).toMatchObject({ color: PRINT_INK.keyword, bold: true });
    expect(runWithText(all, "42")).toMatchObject({ color: PRINT_INK.number, bold: false });
    expect(runWithText(all, "// the answer")).toMatchObject({
      color: PRINT_INK.comment,
      italic: true,
    });

    // Python (whose info string carried an attribute the adapter had to strip)
    expect(runWithText(all, "# greet")).toMatchObject({ color: PRINT_INK.comment, italic: true });
    expect(runWithText(all, "greet")).toMatchObject({ color: PRINT_INK.callable });

    // SQL
    expect(runWithText(all, "SELECT")).toMatchObject({ color: PRINT_INK.keyword, bold: true });

    // Bash (reached through the `sh` alias)
    expect(runWithText(all, "#!/usr/bin/env bash")).toMatchObject({ color: PRINT_INK.meta });
    expect(runWithText(all, '"hello"')).toMatchObject({ color: PRINT_INK.string });

    // The literal XML the AC is about, not just the parsed view of it.
    expect(document).toContain(`<w:color w:val="${PRINT_INK.keyword}"/>`);
    expect(document).toContain(`<w:color w:val="${PRINT_INK.string}"/>`);
    expect(document).toContain(`<w:color w:val="${PRINT_INK.comment}"/>`);
    expect(document).toContain(`<w:color w:val="${PRINT_INK.meta}"/>`);
  });

  it("takes those colours from the theme, not from the adapter", async () => {
    const { document } = await renderFences(THEMES.default);
    const all = codeRuns(document);

    // Same highlighter, same fences, different palette - which is the whole
    // point of reporting scopes instead of colours.
    expect(runWithText(all, "const")).toMatchObject({
      color: DEFAULT_THEME.codePalette["keyword"],
      bold: true,
    });
    expect(runWithText(all, "// the answer")).toMatchObject({
      color: DEFAULT_THEME.codePalette["comment"],
      italic: true,
    });
    expect(document).not.toContain(`<w:color w:val="${PRINT_INK.keyword}"/>`);
  });

  it("leaves an unknown-language fence as plain monospace", async () => {
    const { document } = await renderFences();

    // The klingon fence is the same source as the ts one, so the only way to
    // tell them apart is that its paragraph carries no colour at all.
    const uncoloured = paragraphs(document)
      .filter((p) => pStyle(p) === "CodeBlock")
      .filter((p) => runs(p).every((run) => run.color === null));

    expect(uncoloured).toHaveLength(1);
    expect(textOf(uncoloured[0] ?? "")).toBe(TS.trimEnd());
  });

  it("keeps every CodeBlock paragraph on the monospace style", async () => {
    const { document } = await renderFences();
    const codeParagraphs = paragraphs(document).filter((p) => pStyle(p) === "CodeBlock");

    // 1 (ts) + 3 (py) + 1 (sql) + 2 (bash) + 1 (klingon)
    expect(codeParagraphs).toHaveLength(8);
    // Colour is direct formatting; the font is never stamped on a run.
    expect(document).not.toContain("<w:rFonts");
  });

  it("reproduces the source text of every fence", async () => {
    const { text } = await renderFences();

    for (const source of [TS, PY, SQL, BASH]) {
      for (const line of source.trimEnd().split("\n")) {
        expect(text).toContain(line);
      }
    }
  });

  it("renders the same colours through the theme palette", async () => {
    // The renderer's other highlighting path: offset+scope spans coloured from
    // theme.codePalette. PRINT_CODE_PALETTE is what keeps the two in step.
    const block = codeBlock("SELECT 1;\n", {
      lang: "sql",
      highlights: [{ start: 0, end: 6, scope: "keyword" }],
    });
    const parts = await renderParts(doc([block]), {
      theme: { codePalette: PRINT_CODE_PALETTE },
    });

    expect(runWithText(codeRuns(parts.document), "SELECT").color).toBe(PRINT_INK.keyword);
  });
});
