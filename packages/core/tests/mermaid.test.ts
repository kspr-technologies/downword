import { fileURLToPath } from "node:url";

import { build, type Metafile } from "esbuild";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  blockquote,
  bulletList,
  codeBlock,
  doc,
  heading,
  listItem,
  paragraph,
  text,
  type DocumentNode,
} from "../src/model.js";
import { isDownwordError } from "../src/errors.js";
import type { ImageRasterizer, RasterizeRequest } from "../src/images/types.js";
import type { ResolvedRasterImage } from "../src/model.js";
import {
  countMermaidDiagrams,
  createBrowserMermaidRenderer,
  createCanvasRasterizer,
  detectDomSupport,
  isEngineUnavailable,
  isMermaidFence,
  MERMAID_DIAGNOSTIC_CODES,
  mermaidWarningSeverity,
  readMermaidFence,
  renderMermaid,
  unwrapMermaid,
  type MermaidRenderRequest,
  type MermaidRenderer,
  type MermaidWarning,
} from "../src/plugins/mermaid.js";
import { listDocxParts, renderParts, writeFixture } from "./helpers/docx.js";
import { paragraphs, pStyle, textOf } from "./helpers/xml.js";

/**
 * mermaid: what is genuinely tested, and what is stubbed.
 *
 * Being precise about this matters more here than anywhere else in the suite,
 * because mermaid needs a real DOM and CI has none. So:
 *
 * **Genuinely exercised, end to end, in Node.**
 * - Fence recognition, caption and alt-text derivation.
 * - The whole pass: sizing, the 2x raster rule, the model rewrite, node
 *   identity, every failure path, and the resulting `.docx` — including the
 *   drawing, the media part, the `wp:extent` and the alt text — through the
 *   real renderer and the real packer, with the two seams
 *   (`MermaidRenderer`, `ImageRasterizer`) injected. This is the same seam the
 *   image pipeline already defines for SVG, not one invented for the test.
 * - The **real Node path**: no stubs at all, real defaults, no DOM. It warns
 *   and degrades, and the fence survives into the document.
 * - The **real dynamic import** of the optional peer dependency, failing at
 *   module evaluation in a runtime with no `window` — the same one-fact-per-run
 *   path a browser with a missing dependency takes.
 * - The canvas rasteriser's own logic (scale, clamping, background, export
 *   paths, error handling) against a stubbed `document`.
 *
 * **Stubbed, and therefore not proof of anything.**
 * - mermaid itself. No test here parses a diagram or produces a real SVG: the
 *   fixtures are hand-written SVG documents shaped like mermaid's output.
 * - The raster. A fake `<canvas>` returns a known PNG rather than drawing
 *   pixels, so "the diagram looks right" is not a claim this file makes. What
 *   it does claim is that whatever the canvas returns is embedded correctly.
 */

/* -------------------------------------------------------------------------- */
/* Fixtures                                                                    */
/* -------------------------------------------------------------------------- */

/** A real 8x4 PNG (solid #2F5496), standing in for whatever a canvas produces. */
const PNG_BYTES = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAEUlEQVR42mPQD5mGFTFQTwIA0B4jIb0iwL8AAAAASUVORK5CYII=",
    "base64",
  ),
);

const FLOWCHART = `flowchart LR
  Browser --> Worker
  Worker --> Word`;

const SEQUENCE = `sequenceDiagram
  participant U as User
  participant D as downword
  U->>D: paste markdown
  D-->>U: .docx`;

/** Shaped like mermaid's own output: width/height plus a viewBox. */
function flowchartSvg(id: string): string {
  return (
    `<svg id="${id}" aria-roledescription="flowchart-v2" role="graphics-document document" ` +
    `viewBox="0 0 402 74" style="max-width: 402px;" width="402" height="74" ` +
    `xmlns="http://www.w3.org/2000/svg"><g class="nodes"><rect/></g></svg>`
  );
}

/** viewBox only - the other half of `probeSvg`'s intrinsic-size logic. */
function sequenceSvg(id: string): string {
  return (
    `<svg id="${id}" aria-roledescription="sequence" viewBox="0 0 640 480" ` +
    `xmlns="http://www.w3.org/2000/svg"><g class="actor"><line/></g></svg>`
  );
}

/** A renderer that returns a fixed SVG per diagram and records what it was asked. */
function fakeRenderer(
  svgFor: (request: MermaidRenderRequest) => string | Error,
  log: MermaidRenderRequest[] = [],
): MermaidRenderer & { readonly calls: MermaidRenderRequest[] } {
  return {
    calls: log,
    async render(request: MermaidRenderRequest) {
      log.push(request);
      const result = svgFor(request);
      if (result instanceof Error) throw result;
      return { svg: result };
    },
  };
}

/** A rasteriser that returns a known PNG at `scale` times the requested size. */
function fakeRasterizer(
  scale = 2,
  log: RasterizeRequest[] = [],
): ImageRasterizer & { readonly calls: RasterizeRequest[] } {
  return {
    calls: log,
    async rasterize(request: RasterizeRequest): Promise<ResolvedRasterImage> {
      log.push(request);
      return {
        format: "png",
        data: PNG_BYTES,
        width: Math.round((request.targetWidth ?? 300) * scale),
        height: Math.round((request.targetHeight ?? 150) * scale),
      };
    },
  };
}

function mermaidDoc(...sources: readonly string[]): DocumentNode {
  return doc(sources.map((source) => codeBlock(`${source}\n`, { lang: "mermaid" })));
}

/** Renders with the two seams injected, which is the only DOM-free way to do it. */
async function pass(document_: DocumentNode, options: Parameters<typeof renderMermaid>[1] = {}) {
  const renderer = fakeRenderer((request) =>
    request.source.startsWith("sequenceDiagram")
      ? sequenceSvg(request.id)
      : flowchartSvg(request.id),
  );
  const rasterizer = fakeRasterizer();
  const result = await renderMermaid(document_, { renderer, rasterizer, ...options });
  return { ...result, renderer, rasterizer };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

/* -------------------------------------------------------------------------- */
/* Fences                                                                      */
/* -------------------------------------------------------------------------- */

describe("mermaid: reading a fence", () => {
  it("recognises the fence, whatever case it is written in", () => {
    expect(isMermaidFence(codeBlock("graph TD", { lang: "mermaid" }))).toBe(true);
    expect(isMermaidFence(codeBlock("graph TD", { lang: "Mermaid" }))).toBe(true);
    expect(isMermaidFence(codeBlock("graph TD", { lang: "ts" }))).toBe(false);
    expect(isMermaidFence(codeBlock("graph TD"))).toBe(false);
    expect(isMermaidFence(paragraph([text("graph TD")]))).toBe(false);
  });

  it("takes the caption from the info string", () => {
    const info = readMermaidFence(
      codeBlock(`${FLOWCHART}\n`, { lang: "mermaid", meta: "The request pipeline" }),
    );

    expect(info.caption).toBe("The request pipeline");
    expect(info.alt).toBe("The request pipeline");
  });

  it("unquotes a quoted info string", () => {
    const info = readMermaidFence(
      codeBlock("graph TD\n", { lang: "mermaid", meta: '"A quoted caption"' }),
    );

    expect(info.caption).toBe("A quoted caption");
  });

  it("falls back to mermaid's own title: frontmatter", () => {
    const info = readMermaidFence(
      codeBlock(
        "---\ntitle: Deployment flow\nconfig:\n  theme: base\n---\nflowchart LR\n A-->B\n",
        {
          lang: "mermaid",
        },
      ),
    );

    expect(info.caption).toBe("Deployment flow");
  });

  it("describes the diagram kind when there is no caption", () => {
    const kinds: readonly (readonly [string, string])[] = [
      [FLOWCHART, "Mermaid flowchart"],
      [SEQUENCE, "Mermaid sequence diagram"],
      ["classDiagram\n  A <|-- B", "Mermaid class diagram"],
      ["erDiagram\n  A ||--o{ B : has", "Mermaid entity-relationship diagram"],
      ["gantt\n  title X", "Mermaid Gantt chart"],
      ["stateDiagram-v2\n  [*] --> A", "Mermaid state diagram"],
      ["mindmap\n  root", "Mermaid mind map"],
      ["notADiagram\n  x", "Mermaid diagram"],
    ];

    for (const [source, alt] of kinds) {
      expect(readMermaidFence(codeBlock(`${source}\n`, { lang: "mermaid" })).alt).toBe(alt);
    }
  });

  it("looks past directives and comments to find the kind", () => {
    const info = readMermaidFence(
      codeBlock("%%{init: {'theme':'forest'}}%%\n%% a note\nsequenceDiagram\n  A->>B: hi\n", {
        lang: "mermaid",
      }),
    );

    expect(info.kind).toBe("sequence diagram");
    expect(info.caption).toBeNull();
  });

  it("reports an empty fence as empty", () => {
    expect(readMermaidFence(codeBlock("   \n\n", { lang: "mermaid" })).source).toBe("");
  });

  it("counts diagrams without touching a DOM", () => {
    const document_ = doc([
      heading(1, [text("Diagrams")], { id: "diagrams" }),
      codeBlock(`${FLOWCHART}\n`, { lang: "mermaid" }),
      codeBlock("const x = 1;\n", { lang: "ts" }),
      blockquote([codeBlock(`${SEQUENCE}\n`, { lang: "mermaid" })]),
      bulletList([listItem([codeBlock("graph TD\n", { lang: "mermaid" })])]),
    ]);

    expect(countMermaidDiagrams(document_)).toBe(3);
    expect(countMermaidDiagrams(doc([paragraph([text("none here")])]))).toBe(0);
  });
});

/* -------------------------------------------------------------------------- */
/* The pass, with the seams injected                                           */
/* -------------------------------------------------------------------------- */

describe("mermaid: the pass", () => {
  it("turns a flowchart fence into a centred figure with a caption", async () => {
    const document_ = doc([
      codeBlock(`${FLOWCHART}\n`, { lang: "mermaid", meta: "The request pipeline" }),
    ]);
    const { document: rewritten, diagrams, rendered, warnings } = await pass(document_);

    expect({ diagrams, rendered, warnings }).toEqual({ diagrams: 1, rendered: 1, warnings: [] });

    const parts = await renderParts(rewritten);
    const found = paragraphs(parts.document);

    // A lone image is a `Figure`: centred, with the figure spacing. The caption
    // is an ordinary italic paragraph underneath.
    expect(pStyle(found[0] ?? "")).toBe("Figure");
    expect(found[0] ?? "").toContain("<w:drawing>");
    expect(pStyle(found[1] ?? "")).toBe("Normal");
    expect(textOf(found[1] ?? "")).toBe("The request pipeline");
    expect(found[1] ?? "").toContain("<w:i/>");

    // The bytes are really in the package, as a media part. docx names those
    // by content hash rather than by index, which is also how it dedupes the
    // same picture used twice - hence the pattern rather than a filename.
    const media = (await listDocxParts(parts.bytes)).filter((entry) =>
      /^word\/media\/[0-9a-f]{40}\.png$/.test(entry),
    );
    expect(media).toHaveLength(1);
    expect(parts.warnings).toEqual([]);
  });

  it("embeds the picture at its display size, not at the raster's", async () => {
    const { document: rewritten, rasterizer } = await pass(mermaidDoc(FLOWCHART));
    const parts = await renderParts(rewritten);

    // The SVG is 402x74 and fits the column, so that is the display size; the
    // rasteriser was asked for exactly that and returned twice as many pixels.
    const request = rasterizer.calls[0];
    expect(request?.targetWidth).toBe(402);
    expect(request?.targetHeight).toBe(74);
    expect(request?.format).toBe("svg");
    expect(request?.intrinsicWidth).toBe(402);

    // <wp:extent> is EMU: 9525 per CSS pixel. 402 px, not the raster's 804.
    expect(parts.document).toContain(`cx="${Math.round(402 * 9525)}"`);
    expect(parts.document).toContain(`cy="${Math.round(74 * 9525)}"`);
  });

  it("scales an oversized sequence diagram down to the text column", async () => {
    const { rasterizer, rendered } = await pass(mermaidDoc(SEQUENCE));

    // A4 minus one-inch margins is 9026 twips, which `twipsToPixels` floors to
    // 601 px. 640x480 scaled by 601/640 is 601x451, aspect ratio preserved.
    expect(rendered).toBe(1);
    expect(rasterizer.calls[0]?.targetWidth).toBe(601);
    expect(rasterizer.calls[0]?.targetHeight).toBe(451);
    // Measured from the viewBox, since this SVG declares no width/height.
    expect(rasterizer.calls[0]?.intrinsicWidth).toBe(640);
  });

  it("writes alt text a screen reader can use", async () => {
    const withCaption = await pass(
      doc([codeBlock(`${SEQUENCE}\n`, { lang: "mermaid", meta: "How a paste becomes a file" })]),
    );
    const without = await pass(mermaidDoc(SEQUENCE));

    const captioned = await renderParts(withCaption.document);
    const plain = await renderParts(without.document);

    expect(captioned.document).toContain('descr="How a paste becomes a file"');
    expect(captioned.document).toContain('title="How a paste becomes a file"');
    // No caption: the alt text describes the diagram rather than repeating its
    // source, which is what a `wp:docPr` is for.
    expect(plain.document).toContain('descr="Mermaid sequence diagram"');
  });

  it("renders every diagram in the document, in order, one at a time", async () => {
    const { renderer, rendered } = await pass(mermaidDoc(FLOWCHART, SEQUENCE, FLOWCHART));

    expect(rendered).toBe(3);
    expect(renderer.calls.map((call) => call.diagram)).toEqual([1, 2, 3]);
    // Deterministic ids: mermaid writes them into the SVG, so two conversions
    // of one document produce the same bytes.
    expect(renderer.calls.map((call) => call.id)).toEqual([
      "downword-mermaid-1",
      "downword-mermaid-2",
      "downword-mermaid-3",
    ]);
  });

  it("replaces fences wherever they are nested, and leaves everything else identical", async () => {
    const prose = paragraph([text("Before.")]);
    const other = codeBlock("const x = 1;\n", { lang: "ts" });
    const document_ = doc([
      prose,
      other,
      blockquote([codeBlock(`${FLOWCHART}\n`, { lang: "mermaid" })]),
      bulletList([
        listItem([paragraph([text("item")]), codeBlock(`${SEQUENCE}\n`, { lang: "mermaid" })]),
      ]),
    ]);

    const { document: rewritten, rendered } = await pass(document_);

    expect(rendered).toBe(2);
    // Untouched nodes come back as the same objects, so an ImageMap or
    // HighlightMap keyed by identity still resolves.
    expect(rewritten.children[0]).toBe(prose);
    expect(rewritten.children[1]).toBe(other);
    expect(rewritten).not.toBe(document_);

    const quote = rewritten.children[2];
    expect(quote?.type).toBe("blockquote");
    expect(quote?.type === "blockquote" && quote.children[0]?.type).toBe("paragraph");

    const list = rewritten.children[3];
    const item = list?.type === "list" ? list.children[0] : undefined;
    expect(item?.children[1]?.type).toBe("paragraph");
  });

  it("returns the very same document when there is nothing to do", async () => {
    const document_ = doc([paragraph([text("no diagrams here")])]);
    const result = await renderMermaid(document_);

    expect(result.document).toBe(document_);
    expect(result).toMatchObject({ diagrams: 0, rendered: 0, warnings: [] });
  });

  it("omits the caption paragraph when asked, keeping the alt text", async () => {
    const document_ = doc([codeBlock(`${FLOWCHART}\n`, { lang: "mermaid", meta: "A caption" })]);
    const { document: rewritten } = await pass(document_, { caption: false });

    const parts = await renderParts(rewritten);
    expect(paragraphs(parts.document)).toHaveLength(1);
    expect(parts.document).toContain('descr="A caption"');
  });

  it("ships the SVG and its raster twin when asked for a vector", async () => {
    const { document: rewritten } = await pass(mermaidDoc(FLOWCHART), { embed: "svg" });
    const parts = await renderParts(rewritten);
    const entries = await listDocxParts(parts.bytes);

    // OOXML stores both halves; Word 2016+ draws the vector, everything else
    // draws the twin. `ImageRun` refuses `type: "svg"` without one.
    expect(parts.document).toContain("asvg:svgBlip");
    expect(entries.filter((entry) => /\.svg$/.test(entry))).toHaveLength(1);
    expect(entries.filter((entry) => /^word\/media\/[0-9a-f]{40}\.png$/.test(entry))).toHaveLength(
      1,
    );
  });

  it("writes a .docx for the validity gate", async () => {
    const document_ = doc([
      heading(1, [text("Diagrams")], { id: "diagrams" }),
      codeBlock(`${FLOWCHART}\n`, { lang: "mermaid", meta: "The request pipeline" }),
      codeBlock(`${SEQUENCE}\n`, { lang: "mermaid" }),
    ]);
    const { document: rewritten } = await pass(document_);
    const parts = await renderParts(rewritten);

    await writeFixture("golden-mermaid.docx", parts.bytes);
    expect(parts.warnings).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* Degrading, loudly                                                           */
/* -------------------------------------------------------------------------- */

/** Every warning code, as a set, for terse assertions. */
function codes(warnings: readonly MermaidWarning[]): string[] {
  return warnings.map((warning) => warning.code);
}

describe("mermaid: no DOM (the real Node path, nothing stubbed)", () => {
  it("keeps every fence, warns once, and says what to do about it", async () => {
    const document_ = mermaidDoc(FLOWCHART, SEQUENCE);
    const seen: MermaidWarning[] = [];

    // No injected seams: this is exactly what a CLI gets.
    const result = await renderMermaid(document_, { onWarning: (w) => seen.push(w) });

    expect(result.document).toBe(document_);
    expect(result).toMatchObject({ diagrams: 2, rendered: 0 });
    expect(codes(result.warnings)).toEqual(["no-dom"]);
    expect(seen).toEqual(result.warnings);

    const warning = result.warnings[0];
    expect(warning?.severity).toBe("notice");
    expect(warning?.diagram).toBeNull();
    // Actionable: what happened, how many, why, and the two ways out.
    expect(warning?.message).toContain("2 mermaid diagrams");
    expect(warning?.message).toContain("fenced code blocks");
    expect(warning?.message).toContain("no DOM");
    expect(warning?.message).toContain("renderer");
  });

  it("leaves the diagram source readable in the .docx", async () => {
    const { document: rewritten } = await renderMermaid(mermaidDoc(FLOWCHART));
    const parts = await renderParts(rewritten);

    // The whole point of degrading rather than dropping: the reader still gets
    // the diagram, as its source.
    expect(pStyle(paragraphs(parts.document)[0] ?? "")).toBe("CodeBlock");
    expect(parts.document).toContain("flowchart LR");
    expect(parts.document).not.toContain("<w:drawing>");
  });

  it("detects the missing DOM rather than guessing at the runtime", () => {
    expect(detectDomSupport()).toEqual({
      renderer: false,
      rasterizer: false,
      reason: expect.stringContaining("no DOM"),
    });
  });
});

describe("mermaid: the engine cannot be loaded (the real dynamic import)", () => {
  /** Enough DOM for `detectDomSupport` to choose the real defaults. */
  function stubDom(): void {
    vi.stubGlobal("document", {
      createElement: (tag: string) => (tag === "canvas" ? { getContext: () => ({}) } : {}),
      body: {},
      getElementById: () => null,
    });
  }

  it("reports engine-unavailable once and stops, rather than failing per diagram", async () => {
    stubDom();
    // The real `import("mermaid")` is exercised here and really does fail, but
    // *not* because the package is absent: pnpm's `autoInstallPeers` installs
    // declared peers, so mermaid is present in this repository's node_modules.
    // It fails at module evaluation with `ReferenceError: window is not
    // defined` — `stubDom` supplies `document` and deliberately not `window`,
    // which is a fair likeness of a worker, and is a load failure either way.
    const result = await renderMermaid(mermaidDoc(FLOWCHART, SEQUENCE, FLOWCHART));

    expect(codes(result.warnings)).toEqual(["engine-unavailable"]);
    expect(result).toMatchObject({ diagrams: 3, rendered: 0 });
    expect(result.warnings[0]?.severity).toBe("notice");
    expect(result.warnings[0]?.message).toContain("npm install mermaid");
    expect(result.warnings[0]?.message).toContain("3 mermaid diagrams");
  });

  it("marks a load failure so it is not mistaken for a broken diagram", async () => {
    const renderer = createBrowserMermaidRenderer({
      load: () => Promise.reject(new Error("network down")),
    });
    const error = await renderer
      .render({ source: FLOWCHART, id: "x", diagram: 1 })
      .catch((e: unknown) => e);

    expect(isEngineUnavailable(error)).toBe(true);
    expect(String(error)).toContain("network down");
    expect(isEngineUnavailable(new Error("boom"))).toBe(false);
  });

  it("does not cache a failed load, so installing mermaid and retrying works", async () => {
    let attempts = 0;
    const api = {
      initialize: vi.fn(),
      render: vi.fn(async () => ({ svg: flowchartSvg("id") })),
    };
    const renderer = createBrowserMermaidRenderer({
      load: () => {
        attempts += 1;
        return attempts === 1 ? Promise.reject(new Error("not installed")) : Promise.resolve(api);
      },
    });

    await expect(renderer.render({ source: FLOWCHART, id: "a", diagram: 1 })).rejects.toThrow();
    await expect(
      renderer.render({ source: FLOWCHART, id: "b", diagram: 1 }),
    ).resolves.toMatchObject({ svg: expect.stringContaining("<svg") });
    expect(attempts).toBe(2);
    // Loaded once it succeeded, and initialised exactly once.
    expect(api.initialize).toHaveBeenCalledTimes(1);
  });

  it("initialises mermaid with startOnLoad off and its sanitiser on", async () => {
    const api = {
      initialize: vi.fn(),
      render: vi.fn(async (id: string) => ({ svg: flowchartSvg(id) })),
    };
    const renderer = createBrowserMermaidRenderer({
      load: () => Promise.resolve(api),
      config: { theme: "neutral", securityLevel: "loose" },
    });

    await renderer.render({ source: FLOWCHART, id: "downword-mermaid-1", diagram: 1 });

    // Defaults first, caller's config over the top: startOnLoad stays off
    // (rewriting the host page is not a conversion), and an explicit override
    // of securityLevel is honoured because it is the caller's document.
    expect(api.initialize).toHaveBeenCalledWith({
      startOnLoad: false,
      securityLevel: "loose",
      theme: "neutral",
    });
    expect(api.render).toHaveBeenCalledWith("downword-mermaid-1", FLOWCHART);
  });

  it("unwraps whichever module shape the bundler produces", () => {
    const api = { initialize: () => {}, render: async () => ({ svg: "" }) };

    expect(unwrapMermaid(api)).toBe(api);
    expect(unwrapMermaid({ default: api })).toBe(api);
    expect(() => unwrapMermaid({ nope: true })).toThrow(/does not look like mermaid/);
    expect(() => unwrapMermaid(null)).toThrow(/does not look like mermaid/);
  });
});

describe("mermaid: one broken diagram costs one diagram", () => {
  it("keeps the fence when mermaid rejects, and renders its neighbours", async () => {
    const renderer = fakeRenderer((request) =>
      request.diagram === 2 ? new Error("Parse error on line 2") : flowchartSvg(request.id),
    );
    const result = await renderMermaid(mermaidDoc(FLOWCHART, "flowchart LR\n  A --", FLOWCHART), {
      renderer,
      rasterizer: fakeRasterizer(),
    });

    expect(result).toMatchObject({ diagrams: 3, rendered: 2 });
    expect(codes(result.warnings)).toEqual(["render-failed"]);
    expect(result.warnings[0]).toMatchObject({ severity: "error", diagram: 2 });
    expect(result.warnings[0]?.message).toContain("Parse error on line 2");

    // The broken one is still a code block; the other two are pictures. (A
    // fence is one `CodeBlock` paragraph *per line* - that is how the renderer
    // keeps a listing's line breaks - and the broken source here has two.)
    const parts = await renderParts(result.document);
    const styles = paragraphs(parts.document).map((p) => pStyle(p));
    expect(styles).toEqual(["Figure", "CodeBlock", "CodeBlock", "Figure"]);
  });

  it("reports an empty fence without pretending to draw it", async () => {
    const result = await pass(doc([codeBlock("\n \n", { lang: "mermaid" })]));

    expect(codes(result.warnings)).toEqual(["empty-diagram"]);
    expect(result.warnings[0]?.severity).toBe("notice");
    expect(result).toMatchObject({ rendered: 0 });
  });

  it("rejects something that is not an SVG", async () => {
    const result = await renderMermaid(mermaidDoc(FLOWCHART), {
      renderer: fakeRenderer(() => "<html>error</html>"),
      rasterizer: fakeRasterizer(),
    });

    expect(codes(result.warnings)).toEqual(["invalid-svg"]);
    expect(result).toMatchObject({ rendered: 0 });
  });

  it("treats a rasteriser that throws, returns nothing, or returns nonsense as one lost diagram", async () => {
    const throwing: ImageRasterizer = {
      rasterize: () => Promise.reject(new Error("canvas is tainted")),
    };
    const empty: ImageRasterizer = { rasterize: async () => null };
    const nonsense: ImageRasterizer = {
      // Zero bytes: `isValidResolvedImage` is the same contract check the image
      // pass applies to a third-party resolver.
      rasterize: async () => ({ format: "png", data: new Uint8Array(0), width: 10, height: 10 }),
    };

    for (const rasterizer of [throwing, empty, nonsense]) {
      const result = await renderMermaid(mermaidDoc(FLOWCHART), {
        renderer: fakeRenderer((request) => flowchartSvg(request.id)),
        rasterizer,
      });

      expect(codes(result.warnings)).toEqual(["rasterize-failed"]);
      expect(result.warnings[0]?.severity).toBe("error");
      expect(result.document).toBe(result.document);
      expect(result.rendered).toBe(0);
    }
  });

  it("gives up on a renderer that never settles", async () => {
    const hanging: MermaidRenderer = { render: () => new Promise(() => {}) };
    const result = await renderMermaid(mermaidDoc(FLOWCHART), {
      renderer: hanging,
      rasterizer: fakeRasterizer(),
      timeoutMs: 20,
    });

    expect(codes(result.warnings)).toEqual(["render-failed"]);
    expect(result.warnings[0]?.message).toContain("longer than 20ms");
  });

  it("warns when there is a renderer but nothing to rasterise with", async () => {
    const result = await renderMermaid(mermaidDoc(FLOWCHART), {
      renderer: fakeRenderer((request) => flowchartSvg(request.id)),
    });

    expect(codes(result.warnings)).toEqual(["rasterizer-unavailable"]);
    expect(result.warnings[0]?.message).toContain("rasterizer");
    expect(result.warnings[0]?.message).toContain("resvg");
  });

  it("survives a warning handler that throws", async () => {
    const result = await renderMermaid(mermaidDoc(FLOWCHART), {
      onWarning: () => {
        throw new Error("my logger is broken");
      },
    });

    expect(codes(result.warnings)).toEqual(["no-dom"]);
  });

  it("rejects impossible options loudly, before doing any work", async () => {
    for (const options of [
      { scale: 0 },
      { scale: Number.NaN },
      { maxWidthTwips: -1 },
      { timeoutMs: -5 },
    ]) {
      const error = await renderMermaid(mermaidDoc(FLOWCHART), options).catch((e: unknown) => e);
      expect(isDownwordError(error) && error.code).toBe("invalid-options");
    }
  });

  it("keeps one severity per code, decided once", () => {
    expect([...MERMAID_DIAGNOSTIC_CODES].sort()).toEqual([
      "empty-diagram",
      "engine-unavailable",
      "invalid-svg",
      "no-dom",
      "rasterize-failed",
      "rasterizer-unavailable",
      "render-failed",
    ]);
    for (const code of MERMAID_DIAGNOSTIC_CODES) {
      expect(["error", "notice"]).toContain(mermaidWarningSeverity(code));
    }
    // The rule: could-never-have-worked is a notice, tried-and-failed is an error.
    expect(mermaidWarningSeverity("no-dom")).toBe("notice");
    expect(mermaidWarningSeverity("render-failed")).toBe("error");
  });
});

/* -------------------------------------------------------------------------- */
/* The canvas rasteriser                                                       */
/* -------------------------------------------------------------------------- */

interface FakeCanvas {
  width: number;
  height: number;
  readonly fills: string[];
  readonly drawn: (readonly [number, number])[];
}

/**
 * A `document` whose canvas records what it was asked to do and returns a known
 * PNG. Not a renderer: it proves the glue, not the pixels.
 */
function stubCanvasDom(options: { readonly export?: "blob" | "dataURL" | "none" } = {}): {
  readonly canvas: FakeCanvas;
  readonly imageSrc: () => string;
} {
  const canvas: FakeCanvas = { width: 0, height: 0, fills: [], drawn: [] };
  let src = "";

  const element = {
    get width() {
      return canvas.width;
    },
    set width(value: number) {
      canvas.width = value;
    },
    get height() {
      return canvas.height;
    },
    set height(value: number) {
      canvas.height = value;
    },
    getContext: () => ({
      fillStyle: "",
      fillRect: (_x: number, _y: number, w: number, h: number) => {
        canvas.fills.push(`${w}x${h}`);
      },
      drawImage: (_image: unknown, _x: number, _y: number, w: number, h: number) => {
        canvas.drawn.push([w, h]);
      },
    }),
    ...(options.export === "dataURL" || options.export === "none"
      ? {}
      : {
          toBlob: (callback: (blob: Blob | null) => void) => {
            callback(new Blob([PNG_BYTES], { type: "image/png" }));
          },
        }),
    ...(options.export === "dataURL"
      ? {
          toDataURL: () => `data:image/png;base64,${Buffer.from(PNG_BYTES).toString("base64")}`,
        }
      : {}),
  };

  vi.stubGlobal("document", {
    body: {},
    getElementById: () => null,
    createElement: (tag: string) => {
      if (tag === "canvas") return element;
      const image: Record<string, unknown> = {
        width: 0,
        height: 0,
        onload: null,
        onerror: null,
        decoding: "",
      };
      Object.defineProperty(image, "src", {
        set(value: string) {
          src = value;
          // A real <img> fires asynchronously; so does this one.
          queueMicrotask(() => {
            const onload = image["onload"];
            if (typeof onload === "function") onload();
          });
        },
        get() {
          return src;
        },
      });
      return image;
    },
  });

  return { canvas, imageSrc: () => src };
}

function svgRequest(overrides: Partial<RasterizeRequest> = {}): RasterizeRequest {
  return {
    data: new TextEncoder().encode(flowchartSvg("x")),
    format: "svg",
    src: "mermaid:diagram-1",
    intrinsicWidth: 402,
    intrinsicHeight: 74,
    targetWidth: 402,
    targetHeight: 74,
    ...overrides,
  };
}

describe("mermaid: the canvas rasteriser", () => {
  it("draws at twice the display size and returns PNG bytes", async () => {
    const { canvas } = stubCanvasDom();
    const raster = await createCanvasRasterizer().rasterize(svgRequest());

    expect(canvas.width).toBe(804);
    expect(canvas.height).toBe(148);
    expect(canvas.drawn).toEqual([[804, 148]]);
    expect(raster).toMatchObject({ format: "png", width: 804, height: 148 });
    expect(raster?.data).toEqual(PNG_BYTES);
  });

  it("honours an explicit scale and leaves the alpha channel alone by default", async () => {
    const { canvas } = stubCanvasDom();
    await createCanvasRasterizer({ scale: 1 }).rasterize(svgRequest());

    expect(canvas.width).toBe(402);
    expect(canvas.fills).toEqual([]);
  });

  it("paints a background when asked", async () => {
    const { canvas } = stubCanvasDom();
    await createCanvasRasterizer({ background: "#ffffff" }).rasterize(svgRequest());

    expect(canvas.fills).toEqual(["804x148"]);
  });

  it("reduces the scale rather than allocating a gigabyte", async () => {
    const { canvas } = stubCanvasDom();
    await createCanvasRasterizer({ scale: 8, maxPixels: 1_000_000 }).rasterize(
      svgRequest({ targetWidth: 1000, targetHeight: 1000 }),
    );

    // 1000x1000 at 8x would be 64 megapixels; clamped to the cap, never below 1x.
    expect(canvas.width * canvas.height).toBeLessThanOrEqual(1_000_000);
    expect(canvas.width).toBeGreaterThanOrEqual(1000);
  });

  it("never scales a WebP, whose raster size is its display size", async () => {
    const { canvas } = stubCanvasDom();
    await createCanvasRasterizer({ scale: 4 }).rasterize(
      svgRequest({ format: "webp", data: PNG_BYTES, targetWidth: 100, targetHeight: 50 }),
    );

    expect([canvas.width, canvas.height]).toEqual([100, 50]);
  });

  it("falls back to toDataURL when the canvas has no toBlob", async () => {
    stubCanvasDom({ export: "dataURL" });
    const raster = await createCanvasRasterizer().rasterize(svgRequest());

    expect(raster?.data).toEqual(PNG_BYTES);
  });

  it("fails loudly when the canvas can export nothing at all", async () => {
    stubCanvasDom({ export: "none" });

    await expect(createCanvasRasterizer().rasterize(svgRequest())).rejects.toThrow(
      /neither toBlob\(\) nor toDataURL\(\)/,
    );
  });

  it("hands the browser a URL for the source bytes", async () => {
    const { imageSrc } = stubCanvasDom();
    await createCanvasRasterizer().rasterize(svgRequest());

    // Node has both Blob and URL.createObjectURL, so this is the real path a
    // browser takes rather than a stub of it.
    expect(imageSrc()).toMatch(/^(blob:|data:image\/svg)/);
  });

  it("refuses to run without a DOM", async () => {
    await expect(createCanvasRasterizer().rasterize(svgRequest())).rejects.toThrow(/needs a DOM/);
  });

  it("rejects when the image cannot be decoded", async () => {
    vi.stubGlobal("document", {
      body: {},
      createElement: (tag: string) =>
        tag === "canvas"
          ? { width: 0, height: 0, getContext: () => ({}) }
          : {
              set src(_value: string) {
                queueMicrotask(() => {
                  const onerror = (this as unknown as { onerror?: () => void }).onerror;
                  onerror?.();
                });
              },
              onload: null,
              onerror: null,
            },
    });

    await expect(createCanvasRasterizer().rasterize(svgRequest())).rejects.toThrow(/could not/);
  });
});

/* -------------------------------------------------------------------------- */
/* Bundle discipline                                                           */
/* -------------------------------------------------------------------------- */

const PACKAGE_DIR = fileURLToPath(new URL("..", import.meta.url));

describe("@ksprtech/downword/plugins/mermaid: mermaid is loaded lazily or not at all", () => {
  it("has no static import of mermaid, only a dynamic one", async () => {
    const result = await build({
      absWorkingDir: PACKAGE_DIR,
      entryPoints: ["src/plugins/mermaid.ts"],
      bundle: true,
      format: "esm",
      target: "es2022",
      platform: "browser",
      packages: "external",
      write: false,
      metafile: true,
      logLevel: "silent",
      define: { __DOWNWORD_VERSION__: '"0.0.0-test"' },
    });

    const text_ = result.outputFiles?.[0]?.text ?? "";
    const staticImports = [...text_.matchAll(/(?:^|\n)import\s[^;]*?from\s*["']([^"']+)["']/g)].map(
      (match) => match[1],
    );
    const dynamicImports = [...text_.matchAll(/import\(\s*["']([^"']+)["']\s*\)/g)].map(
      (match) => match[1],
    );

    expect(staticImports).not.toContain("mermaid");
    expect(dynamicImports).toEqual(["mermaid"]);

    // And nothing heavy rode along: the plugin entry is the pass, the two seams
    // and the image helpers it reuses. `render/units.ts` is in the graph
    // because `images/layout.ts` re-exports the twip/pixel/EMU constants from
    // it - one table of numbers, and the single definition of each, which is
    // the whole reason that module owns them.
    const inputs = Object.keys((result.metafile as Metafile).inputs);
    expect(inputs.filter((input) => input.startsWith("src/render/"))).toEqual([
      "src/render/units.ts",
    ]);
    expect(inputs.filter((input) => input.startsWith("src/parse/"))).toEqual([]);
    expect(inputs.filter((input) => input.includes("node_modules"))).toEqual([]);
  });
});
