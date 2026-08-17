/**
 * "Is there a DOM here?", answered once and with a reason.
 *
 * Both defaults in this module — the mermaid renderer and the canvas
 * rasteriser — need a live document, and the interesting part of *not* having
 * one is being able to tell the user which piece is missing. A CLI that prints
 * "mermaid needs a DOM" is actionable; one that prints "diagram failed" is not.
 *
 * The checks are deliberately behavioural rather than a `typeof process` sniff:
 * jsdom, happy-dom, an Electron renderer and a browser tab all pass, and a
 * bundler that shims `document` to an empty object does not.
 */

/** Which capability is missing, and what the caller can do about it. */
export interface DomSupport {
  /** Whether mermaid can be driven here: an element can be created and attached. */
  readonly renderer: boolean;
  /** Whether a canvas can be rastered here. */
  readonly rasterizer: boolean;
  /** One sentence naming what is missing. Empty when everything is present. */
  readonly reason: string;
}

interface CanvasCapableDocument {
  createElement(tag: string): unknown;
  readonly body: unknown;
}

/** The `document` global, or `null` where there is none. */
function ambientDocument(): CanvasCapableDocument | null {
  if (typeof document === "undefined") return null;
  const candidate: unknown = document;
  if (typeof candidate !== "object" || candidate === null) return null;
  const withCreate = candidate as { readonly createElement?: unknown };
  if (typeof withCreate.createElement !== "function") return null;
  return candidate as unknown as CanvasCapableDocument;
}

/** Whether `document.createElement("canvas")` yields something with a 2d context. */
function hasCanvas(dom: CanvasCapableDocument): boolean {
  let element: unknown;
  try {
    element = dom.createElement("canvas");
  } catch {
    return false;
  }
  if (typeof element !== "object" || element === null) return false;
  const canvas = element as { readonly getContext?: unknown };
  return typeof canvas.getContext === "function";
}

/**
 * What this runtime can do.
 *
 * @returns Both capabilities and, when either is missing, a sentence that names
 *   the runtime rather than the symptom.
 */
export function detectDomSupport(): DomSupport {
  const dom = ambientDocument();
  if (dom === null) {
    return {
      renderer: false,
      rasterizer: false,
      reason:
        "this runtime has no DOM (`document` is undefined), and mermaid measures text by laying " +
        "it out, so it cannot run here",
    };
  }
  if (dom.body === null || dom.body === undefined) {
    return {
      renderer: false,
      rasterizer: hasCanvas(dom),
      reason:
        "there is a `document` but no `document.body` to render into yet; run the conversion " +
        "after the document has parsed",
    };
  }
  if (!hasCanvas(dom)) {
    // jsdom without the `canvas` package is exactly this: mermaid renders,
    // nothing can rasterise the result.
    return {
      renderer: true,
      rasterizer: false,
      reason:
        "this DOM has no working <canvas> (`getContext` is missing), so an SVG cannot be turned " +
        "into the raster OOXML needs",
    };
  }
  return { renderer: true, rasterizer: true, reason: "" };
}
