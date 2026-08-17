/**
 * Shared vocabulary for the mermaid pass.
 *
 * The pass has three stages, and each is an injectable seam rather than a hard
 * dependency, because every one of them needs something the core does not have:
 *
 * ```text
 *   ```mermaid fence ──▶ MermaidRenderer ──▶ SVG ──▶ ImageRasterizer ──▶ PNG ──▶ ImageNode
 *                        (mermaid + a DOM)          (a canvas, resvg, sharp…)
 * ```
 *
 * Both defaults are browser-only, which is the single most important fact about
 * this module: **mermaid needs a DOM**. It measures text by laying it out, so
 * there is no headless path that is not a headless *browser*. In Node the pass
 * therefore degrades — every fence stays a fenced code block, so the diagram's
 * source is still in the document — and says so through {@link MermaidWarning}.
 * Silently emitting nothing is the one outcome it will not produce.
 */

import type { BaseWarning, WarningSeverity } from "../warnings.js";

/* -------------------------------------------------------------------------- */
/* Diagnostics                                                                 */
/* -------------------------------------------------------------------------- */

/** Machine-readable reason a `mermaid` fence did not become a picture. */
export type MermaidDiagnosticCode =
  /** No DOM in this runtime (Node, a CLI, a worker without one). Expected, not a fault. */
  | "no-dom"
  /** `import("mermaid")` failed: the optional peer dependency is not installed. */
  | "engine-unavailable"
  /** mermaid rejected the diagram — almost always a syntax error in the source. */
  | "render-failed"
  /** mermaid produced something that is not an SVG document. */
  | "invalid-svg"
  /** No rasteriser, and no DOM to build the default canvas one with. */
  | "rasterizer-unavailable"
  /** The rasteriser returned nothing usable, so there are no bytes to embed. */
  | "rasterize-failed"
  /** The fence was empty or whitespace. Nothing to draw. */
  | "empty-diagram";

/**
 * How much each code matters, by one rule:
 *
 * **A fence that stays a fence is a `notice` when the environment could never
 * have rendered it, and an `error` when downword tried and failed.**
 *
 * The first group (`no-dom`, `engine-unavailable`) is a fact about where the
 * conversion is running — a CLI has no DOM, and that is not a defect in
 * anyone's document. The second (`render-failed`, `rasterize-failed`, …) means
 * a picture the author asked for is missing and something has to be fixed. Both
 * groups are reported; only the second implies the document is wrong.
 *
 * Either way the diagram *source* survives as a code block, which is why none
 * of this is fatal.
 */
const DIAGNOSTIC_SEVERITY: Readonly<Record<MermaidDiagnosticCode, WarningSeverity>> = {
  "no-dom": "notice",
  "engine-unavailable": "notice",
  "render-failed": "error",
  "invalid-svg": "error",
  "rasterizer-unavailable": "error",
  "rasterize-failed": "error",
  "empty-diagram": "notice",
};

/** Every {@link MermaidDiagnosticCode}, for a host building a per-code label table. */
export const MERMAID_DIAGNOSTIC_CODES = Object.keys(
  DIAGNOSTIC_SEVERITY,
) as readonly MermaidDiagnosticCode[];

/** The severity of a diagnostic code. Total by construction. */
export function mermaidWarningSeverity(code: MermaidDiagnosticCode): WarningSeverity {
  return DIAGNOSTIC_SEVERITY[code];
}

/** A non-fatal problem found while turning fences into diagrams. See {@link BaseWarning}. */
export interface MermaidWarning extends BaseWarning {
  readonly code: MermaidDiagnosticCode;
  /**
   * 1-based position of the diagram in the document, or `null` for a warning
   * about the run as a whole (no DOM, no engine — facts that are the same for
   * every fence and are therefore reported once).
   */
  readonly diagram: number | null;
}

/** Called once per {@link MermaidWarning}. Must not throw; if it does, it is ignored. */
export type MermaidWarningHandler = (warning: MermaidWarning) => void;

/* -------------------------------------------------------------------------- */
/* The renderer seam                                                           */
/* -------------------------------------------------------------------------- */

/** One diagram, as handed to a {@link MermaidRenderer}. */
export interface MermaidRenderRequest {
  /** The fence's contents, trimmed of the trailing newline. */
  readonly source: string;
  /**
   * A stable, unique DOM id for this diagram.
   *
   * mermaid renders into an element it creates from this id, and writes the id
   * into the SVG's own `id` and into the CSS selectors inside it. It is derived
   * from the diagram's position rather than from a random number so that
   * converting the same document twice produces the same SVG bytes.
   */
  readonly id: string;
  /** 1-based position in the document. */
  readonly diagram: number;
}

/** What a {@link MermaidRenderer} produces. */
export interface MermaidSvg {
  /** SVG markup. */
  readonly svg: string;
  /**
   * Intrinsic size in CSS pixels, when the renderer knows it.
   *
   * `null` is the normal case: the pass measures the markup itself with
   * `probeSvg`, the same reader the image pipeline uses, so a renderer only
   * needs to answer when it knows better.
   */
  readonly width?: number | null | undefined;
  readonly height?: number | null | undefined;
}

/**
 * Turns mermaid source into SVG markup.
 *
 * The seam that makes this module testable and host-replaceable. The default
 * implementation dynamically imports mermaid and drives it against the live
 * DOM; a host that already has mermaid (from a CDN, from its own bundle, from a
 * worker with an offscreen document) supplies its own and this package never
 * touches the specifier.
 *
 * A renderer may reject. The pass treats that as `render-failed` for that one
 * diagram, keeps the fence, and carries on.
 */
export interface MermaidRenderer {
  render(request: MermaidRenderRequest): Promise<MermaidSvg>;
}
