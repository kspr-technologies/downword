/**
 * Inline nodes -> `ParagraphChild`s.
 *
 * Two things here are less obvious than they look.
 *
 * **Links are marks, not wrappers.** The model applies `link` as a mark on each
 * inline node, so `[**bold** and _italic_](url)` arrives as three separate
 * nodes that happen to share an equal link mark. Emitting one `<w:hyperlink>`
 * per node would produce three adjacent hyperlinks — three relationship
 * entries, three underline segments with visible seams, and three separate
 * click targets. {@link segmentByLink} therefore coalesces runs of adjacent
 * nodes carrying an equal link mark into a single hyperlink, which is both what
 * the author meant and what Word writes itself.
 *
 * **A run may carry only one `<w:rStyle>`.** Inline code and hyperlinks both
 * want one. When they collide (```[`npm i`](url)```) the `CodeChar` style wins
 * — losing the monospace font would misrepresent the content, whereas losing
 * the *style* link for the colour and underline is recoverable, so those two
 * properties are re-applied directly.
 */

import {
  ExternalHyperlink,
  FootnoteReferenceRun,
  ImageRun,
  ImportedXmlComponent,
  InternalHyperlink,
  TextRun,
  UnderlineType,
  type IImageOptions,
  type IRunPropertiesOptions,
  type ParagraphChild,
} from "docx";

import {
  assertNever,
  findMark,
  hasMark,
  nodeText,
  type BlockNode,
  type FootnoteDefinitionNode,
  type FootnoteReferenceNode,
  type ImageNode,
  type InlineNode,
  type LinkMark,
  type Mark,
  type ResolvedImage,
} from "../model.js";
import type { RenderContext } from "./context.js";
import { STYLE_IDS } from "./styles.js";
import { safeText } from "./text.js";
import { MIN_PICTURE_PIXELS, twipsToPixels } from "./units.js";

/* -------------------------------------------------------------------------- */
/* Writing direction                                                           */
/* -------------------------------------------------------------------------- */

/**
 * The right-to-left scripts, by Unicode block.
 *
 * `U+0590`-`U+08FF` is Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan, Mandaic
 * and the Arabic supplement/extended blocks; `U+FB1D`-`U+FDFF` and
 * `U+FE70`-`U+FEFC` are the Hebrew and Arabic presentation forms (`U+FEFF`, the
 * byte-order mark, is deliberately outside the range); the two astral ranges
 * are the right-to-left planes, Cypriot through Adlam.
 *
 * Only used to *notice* right-to-left prose in a left-to-right document, so
 * being coarse is the right trade. A handful of code points inside those blocks
 * are technically neutral or weak (the Arabic comma, the Arabic-Indic digits),
 * and over-reporting one costs a single notice and changes no output; missing a
 * real Hebrew paragraph would cost the reader a mis-set document.
 */
const RTL_CHARACTERS =
  /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFC\u{10800}-\u{10FFF}\u{1E800}-\u{1EFFF}]/u;

/**
 * Run properties for the document's writing direction, and the RTL warning.
 *
 * `<w:rtl/>` does two things a paragraph's `<w:bidi/>` cannot: it tells the
 * Unicode Bidirectional Algorithm that this run's *base* level is odd, and it
 * makes Word draw the run with the `w:cs` face and size it with `w:szCs`
 * instead of falling back to the Latin face at its own 10 pt default. Both are
 * what `styles.ts:sized()` already writes `sizeComplexScript` for.
 *
 * When the document is left-to-right and the text is not, nothing can be
 * emitted — mis-setting the base direction of a Latin document to fix one
 * Arabic phrase would be worse than leaving it. So it is reported instead,
 * once, which is the whole of downword's answer to "no RTL support of any kind,
 * and no warning".
 */
function directionRun(value: string, ctx: RenderContext): IRunPropertiesOptions {
  if (ctx.options.direction === "rtl") return { rightToLeft: true };
  if (RTL_CHARACTERS.test(value)) {
    ctx.warnOnce(
      "rtl-not-enabled",
      "the document contains right-to-left text but its base direction is left-to-right, so " +
        "trailing punctuation, mixed Latin/Arabic segments and list markers will resolve on the " +
        'wrong side; pass direction: "rtl" to renderDocument',
    );
  }
  return {};
}

/* -------------------------------------------------------------------------- */
/* Marks                                                                       */
/* -------------------------------------------------------------------------- */

/** Turns a normalised mark list into run properties. */
export function runProperties(
  marks: readonly Mark[],
  ctx: RenderContext,
  /** Set when the link mark has no reachable target, so it must not be styled. */
  deadLink = false,
): IRunPropertiesOptions {
  const isCode = hasMark(marks, "inlineCode");
  const isLink = !deadLink && findMark(marks, "link") !== null;

  // `<sup><sub>x</sub></sup>` is legal markdown-it output and impossible
  // typography. `EG_RPrBase` is `maxOccurs="unbounded"`, so emitting both
  // `w:vertAlign` elements is schema-valid - but a run has one baseline, Word
  // takes the last, and the inner mark disappears without a word. Resolve it
  // here instead: superscript wins (matching what Word already did with the
  // pair), and the loss is reported rather than inferred.
  const superscript = hasMark(marks, "superscript");
  const subscript = hasMark(marks, "subscript");
  if (superscript && subscript) {
    ctx.warn(
      "conflicting-marks",
      "a run is both superscript and subscript; OOXML gives a run one baseline, so the subscript was dropped",
    );
  }

  return {
    ...(isCode ? { style: STYLE_IDS.codeChar } : isLink ? { style: STYLE_IDS.hyperlink } : {}),
    ...(hasMark(marks, "bold") ? { bold: true } : {}),
    ...(hasMark(marks, "italic") ? { italics: true } : {}),
    ...(hasMark(marks, "strikethrough") ? { strike: true } : {}),
    ...(superscript ? { superScript: true } : subscript ? { subScript: true } : {}),
    ...(hasMark(marks, "highlight") ? { highlight: ctx.theme.highlightColor } : {}),
    // Only one w:rStyle per run: CodeChar took the slot, so paint the two
    // properties that make a link recognisable back on by hand.
    ...(isCode && isLink
      ? { color: ctx.theme.colors.link, underline: { type: UnderlineType.SINGLE } }
      : {}),
  };
}

/* -------------------------------------------------------------------------- */
/* Link segmentation                                                           */
/* -------------------------------------------------------------------------- */

interface LinkSegment {
  readonly link: LinkMark | null;
  readonly nodes: readonly InlineNode[];
}

function sameLink(a: LinkMark | null, b: LinkMark | null): boolean {
  if (a === null || b === null) return a === b;
  return a.href === b.href && a.title === b.title;
}

/** Groups adjacent inline nodes that share an equal link mark. */
function segmentByLink(nodes: readonly InlineNode[]): readonly LinkSegment[] {
  const segments: { link: LinkMark | null; nodes: InlineNode[] }[] = [];
  for (const node of nodes) {
    const link = findMark(node.marks, "link");
    const last = segments[segments.length - 1];
    if (last !== undefined && sameLink(last.link, link)) {
      last.nodes.push(node);
    } else {
      segments.push({ link, nodes: [node] });
    }
  }
  return segments;
}

/* -------------------------------------------------------------------------- */
/* Images                                                                      */
/* -------------------------------------------------------------------------- */

/** Bytes for an image node: from the model first, then the prepared map. */
export function resolveImage(node: ImageNode, ctx: RenderContext): ResolvedImage | null {
  return node.resolved ?? ctx.images?.get(node) ?? null;
}

/**
 * Fits a picture inside the text column, in **both** directions.
 *
 * Word does not reflow an oversized picture, does not shrink one to fit, and
 * does not split an inline picture across a page break. Everything below the
 * first page boundary is simply not drawn. Capping the width alone left the
 * commonest real input — a tall screenshot pasted out of a chat window — as a
 * `400 x 8000` px PNG scaled to 4.17in x 83.33in in a text column 9.19in tall:
 * one strip of the top of the image, and eleven twelfths of it invisible, with
 * no diagnostic.
 *
 * So the scale is `min(1, widthLimit/width, heightLimit/height)`, which
 * preserves the aspect ratio and keeps the whole picture on the page. Being
 * scaled down for height is reported (`image-oversized`), because the reader
 * gets a smaller picture than the author drew and may want to crop it instead;
 * being scaled down for width is not, because that has always been the
 * behaviour and it is what a text column *means*.
 *
 * ### Why the result is not a whole number of pixels
 *
 * `ImageRun.transformation` is in CSS pixels, but the pixel is not the unit the
 * picture is stored in: docx writes `<wp:extent>` as `Math.round(px * 9525)`,
 * so what OOXML records is EMU. Rounding to a whole pixel *here as well* is a
 * second quantisation, 9525 times coarser than the one that has to happen, and
 * it lands entirely on the aspect ratio. This very case is the proof: 400x8000
 * scaled to the 930-pixel column is exactly `46.5 x 930` px, which rounded to
 * `47 x 930` — a ratio of 0.05054 against the picture's own 0.05, wider by
 * 1.1%. Worse shapes are worse: a `4000 x 10` banner scaled to the 601-pixel
 * column is `601 x 1.5` px, and rounding it to `601 x 2` squashes the ratio
 * from 400 to 300.
 *
 * So the pixel size stays fractional and docx's rounding is the only one. The
 * floor is one EMU rather than one pixel, for the same reason it was ever
 * there: `<wp:extent cx="0">` makes Word draw nothing at all.
 */
function fitToColumn(
  node: ImageNode,
  intrinsic: { readonly width: number; readonly height: number },
  availableWidth: number,
  availableHeight: number,
  ctx: RenderContext,
): { readonly width: number; readonly height: number } {
  const { width, height } = intrinsic;
  const widthScale = Math.min(1, twipsToPixels(availableWidth) / width);
  const heightScale = Math.min(1, twipsToPixels(availableHeight) / height);

  if (heightScale < widthScale) {
    ctx.warn(
      "image-oversized",
      `image "${node.src}" is ${width}x${height}px, taller than the text column; ` +
        `it was scaled to ${Math.round(heightScale * 100)}% so the whole picture fits on the page`,
    );
  }

  const scale = Math.min(widthScale, heightScale);
  return {
    width: Math.max(MIN_PICTURE_PIXELS, width * scale),
    height: Math.max(MIN_PICTURE_PIXELS, height * scale),
  };
}

function imageOptions(
  node: ImageNode,
  resolved: ResolvedImage,
  availableWidth: number,
  availableHeight: number,
  ctx: RenderContext,
): IImageOptions | null {
  const { width, height } = resolved;
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return null;
  }

  const transformation = fitToColumn(node, { width, height }, availableWidth, availableHeight, ctx);

  // `alt`, `title` and `src` all reach `docProps`/`wp:docPr` attributes, which
  // are XML like everything else; see `render/text.ts`.
  //
  // `id` is supplied rather than left to docx: its `DocProperties` mints one
  // from a counter it builds *per instance*, so every drawing in the document
  // would come out as `wp:docPr id="1"`. See `RenderContext.nextDrawingId`.
  const alt = safeText(node.alt, ctx);
  const altText = {
    id: String(ctx.nextDrawingId()),
    name: alt !== "" ? alt : "image",
    ...(alt !== "" ? { description: alt } : {}),
    ...(node.title !== null ? { title: safeText(node.title, ctx) } : {}),
  };

  const format = resolved.format;
  if (format === "svg") {
    // OOXML stores an SVG as the SVG *plus* a raster twin for readers that
    // cannot draw it. The model guarantees the twin; refuse the image if a
    // resolver broke that contract rather than emitting an unopenable part.
    const fallback = resolved.fallback;
    if (fallback === null) return null;
    return {
      type: "svg",
      data: resolved.data,
      fallback: { type: fallback.format, data: fallback.data },
      transformation,
      altText,
    };
  }

  return { type: format, data: resolved.data, transformation, altText };
}

function imagePlaceholder(node: ImageNode, ctx: RenderContext): TextRun {
  const label = safeText(node.alt !== "" ? node.alt : node.src, ctx);
  return new TextRun({ text: `[image: ${label}]`, style: STYLE_IDS.imagePlaceholderChar });
}

function renderImage(
  node: ImageNode,
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number,
): readonly ParagraphChild[] {
  const resolved = resolveImage(node, ctx);
  if (resolved === null) {
    ctx.warn("image-unresolved", `image "${node.src}" has no bytes; emitted a placeholder`);
    return [imagePlaceholder(node, ctx)];
  }
  const options = imageOptions(node, resolved, availableWidth, availableHeight, ctx);
  if (options === null) {
    ctx.warn(
      "image-invalid",
      `image "${node.src}" could not be embedded (bad dimensions, or an SVG with no raster fallback)`,
    );
    return [imagePlaceholder(node, ctx)];
  }
  return [new ImageRun(options)];
}

/**
 * Whether a paragraph is "just an image", and whether that image resolved.
 *
 * Drives the paragraph style: a lone image becomes a centred `Figure`, a lone
 * *broken* image becomes an `ImagePlaceholder`, and an image mixed into prose
 * stays inline in an ordinary paragraph.
 */
export function imageParagraphKind(
  children: readonly InlineNode[],
  ctx: RenderContext,
): "none" | "figure" | "placeholder" {
  let images = 0;
  let resolved = 0;
  for (const child of children) {
    if (child.type === "image") {
      images += 1;
      if (resolveImage(child, ctx) !== null) resolved += 1;
      continue;
    }
    // Whitespace around an image does not stop it being a figure.
    if (child.type === "softBreak") continue;
    if (child.type === "text" && child.value.trim() === "") continue;
    return "none";
  }
  if (images === 0) return "none";
  return resolved > 0 ? "figure" : "placeholder";
}

/* -------------------------------------------------------------------------- */
/* Raw OMML (phase 2 math)                                                     */
/* -------------------------------------------------------------------------- */

interface XmlRootHolder {
  readonly root: readonly unknown[];
}

function hasXmlRoot(value: unknown): value is XmlRootHolder {
  if (typeof value !== "object" || value === null || !("root" in value)) return false;
  return Array.isArray((value as { readonly root: unknown }).root);
}

/**
 * Injects a raw `<m:oMath>` fragment into a paragraph.
 *
 * `ImportedXmlComponent.fromXmlString` cannot be used directly: it converts the
 * xml-js *document* node, whose name is `undefined`, so the packer emits a
 * literal `<undefined>` element and corrupts the file. Unwrapping `root[0]`
 * yields the real `m:oMath` component. `root` is `protected`, hence the
 * structural check plus cast — and `ImportedXmlComponent` is not a member of
 * the `ParagraphChild` union either, though the packer accepts it.
 */
export function importOmml(omml: string): ParagraphChild | null {
  try {
    const imported: unknown = ImportedXmlComponent.fromXmlString(omml);
    if (!hasXmlRoot(imported)) return null;
    const first = imported.root[0];
    if (typeof first !== "object" || first === null) return null;
    return first as unknown as ParagraphChild;
  } catch {
    return null;
  }
}

/* -------------------------------------------------------------------------- */
/* Footnotes                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Collapses a block's plain text onto one line.
 *
 * Only used for the block kinds a spliced footnote has to flatten (a fence, a
 * thematic break, raw HTML), where the newlines are structure the parenthesis
 * cannot hold.
 */
function oneLine(value: string): string {
  return value.replace(/\s+/gu, " ").trim();
}

/**
 * One block of a footnote body, rendered as runs inside the sentence.
 *
 * Prose keeps everything — a link in a spliced footnote is still a link, bold
 * is still bold — because a paragraph's inline children are exactly what
 * {@link renderInline} already takes. Blocks that are not prose have no inline
 * form at all, so they degrade to their text: a fence keeps its monospace, and
 * quotes and lists recurse so their prose keeps its marks too.
 */
function footnoteBlockRuns(
  block: BlockNode,
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number,
): readonly ParagraphChild[] {
  switch (block.type) {
    case "paragraph":
    case "heading":
      return renderInline(block.children, ctx, availableWidth, availableHeight);

    case "blockquote":
      return footnoteBlocksRuns(block.children, ctx, availableWidth, availableHeight);

    case "list":
      return footnoteBlocksRuns(
        block.children.flatMap((item) => [...item.children]),
        ctx,
        availableWidth,
        availableHeight,
      );

    case "codeBlock": {
      const value = oneLine(block.value);
      return value === ""
        ? []
        : [new TextRun({ text: safeText(value, ctx), style: STYLE_IDS.codeChar })];
    }

    case "table":
    case "thematicBreak":
    case "htmlBlock":
    case "mathBlock": {
      const value = oneLine(nodeText(block));
      return value === "" ? [] : [new TextRun({ text: safeText(value, ctx) })];
    }

    case "footnoteDefinition":
      // A definition nested inside a definition is not a note; `renderBlocks`
      // reports the same shape as `footnote-misplaced` on the other path.
      ctx.warn(
        "footnote-misplaced",
        `footnote definition [^${block.label}] was nested inside another footnote and was dropped`,
      );
      return [];

    default:
      return assertNever(block, "block node");
  }
}

/** {@link footnoteBlockRuns} over a run of blocks, separated by single spaces. */
function footnoteBlocksRuns(
  blocks: readonly BlockNode[],
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number,
): readonly ParagraphChild[] {
  const out: ParagraphChild[] = [];

  for (const block of blocks) {
    const runs = footnoteBlockRuns(block, ctx, availableWidth, availableHeight);
    if (runs.length === 0) continue;
    if (out.length > 0) out.push(new TextRun({ text: " " }));
    out.push(...runs);
  }

  return out;
}

/**
 * Splices a footnote into the sentence that cited it, in parentheses.
 *
 * Two situations reach here, and they want the same thing. `footnotes: false`
 * asks for it outright — the note goes where a reader with no footnote pane
 * (a converter, a plain-text extraction, a slide) will still see it. And a
 * `[^y]` met *inside* footnote `x` has no other legal home: Word cannot draw a
 * footnote inside a footnote, so the inner note joins the outer one's text
 * rather than becoming a reference `word/footnotes.xml` would have to resolve
 * against itself.
 *
 * The open set is also the cycle guard. `[^a]` and `[^b]` citing each other
 * would otherwise recurse until the stack ran out; the second visit to a note
 * that is already open drops the marker and says so.
 */
function inlineFootnote(
  definition: FootnoteDefinitionNode,
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number,
): readonly ParagraphChild[] {
  if (!ctx.footnotes.open(definition.number)) {
    ctx.warn(
      "footnote-content-dropped",
      `footnote [^${definition.label}] cites itself, directly or through another note; ` +
        `the repeated marker was dropped rather than expanded forever`,
    );
    return [];
  }
  // The note's words are now somewhere a reader can see them, which is what
  // stops `buildFootnotes` reporting it as one nothing points at.
  ctx.footnotes.markInlined(definition.number);

  try {
    const body = footnoteBlocksRuns(definition.children, ctx, availableWidth, availableHeight);
    // An empty note has nothing to say, and " ()" says it worse than nothing.
    if (body.length === 0) return [];
    return [new TextRun({ text: " (" }), ...body, new TextRun({ text: ")" })];
  } finally {
    ctx.footnotes.close(definition.number);
  }
}

/**
 * A `[^x]` marker: a real Word footnote reference, or the note itself inline.
 *
 * The dangling case is the one that matters for file validity. A
 * `<w:footnoteReference w:id="7"/>` whose footnote is not in
 * `word/footnotes.xml` is a reference into nothing — so a marker with no
 * definition never becomes one. It is drawn as superscript text carrying the
 * label the author wrote, which looks the same on the page, resolves to
 * nothing, and cannot corrupt the part.
 */
function renderFootnoteReference(
  node: FootnoteReferenceNode,
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number,
  props: IRunPropertiesOptions,
): readonly ParagraphChild[] {
  const definition = ctx.footnotes.byNumber.get(node.number);

  if (definition === undefined) {
    ctx.warn(
      "footnote-unresolved",
      `footnote reference [^${node.label}] has no definition; the marker was drawn as plain ` +
        `superscript text rather than as a reference into an empty footnote`,
    );
    return [
      new TextRun({
        text: safeText(node.label, ctx),
        ...props,
        style: STYLE_IDS.footnoteReference,
      }),
    ];
  }

  if (!ctx.options.footnotes || ctx.footnotes.depth() > 0) {
    return inlineFootnote(definition, ctx, availableWidth, availableHeight);
  }

  // Emits <w:rStyle w:val="FootnoteReference"/> + <w:footnoteReference w:id="N"/>.
  // Recorded so `buildFootnotes` writes the note this now points at, and only
  // the notes something points at.
  ctx.footnotes.markReferenced(node.number);
  return [new FootnoteReferenceRun(node.number)];
}

/* -------------------------------------------------------------------------- */
/* Inline dispatch                                                             */
/* -------------------------------------------------------------------------- */

function renderInlineNode(
  node: InlineNode,
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number,
  deadLink = false,
): readonly ParagraphChild[] {
  const props = runProperties(node.marks, ctx, deadLink);

  switch (node.type) {
    case "text": {
      // Prose: sanitized, and the one place the document's writing direction is
      // decided per run. Both are no-ops for the overwhelming common case.
      const value = safeText(node.value, ctx);
      return value === ""
        ? []
        : [new TextRun({ text: value, ...props, ...directionRun(value, ctx) })];
    }

    case "hardBreak":
      // `break: 1` emits <w:br/> inside the current run: a line break, not a
      // new paragraph.
      return [new TextRun({ break: 1, ...props })];

    case "softBreak":
      switch (ctx.options.softBreak) {
        case "space":
          return [new TextRun({ text: " ", ...props })];
        case "break":
          return [new TextRun({ break: 1, ...props })];
        case "ignore":
          return [];
        default:
          return assertNever(ctx.options.softBreak, "soft break policy");
      }

    case "image":
      return renderImage(node, ctx, availableWidth, availableHeight);

    case "htmlInline": {
      ctx.warn("html-inline", `inline HTML is not converted: ${node.value}`);
      if (ctx.options.html === "drop") return [];
      // Markup, not prose: no `<w:rtl/>`, because the tag names are Latin.
      const value = safeText(node.value, ctx);
      return [new TextRun({ text: value, ...props, style: STYLE_IDS.htmlChar })];
    }

    case "footnoteReference":
      return renderFootnoteReference(node, ctx, availableWidth, availableHeight, props);

    case "mathInline": {
      if (node.omml !== null) {
        const math = importOmml(node.omml);
        if (math !== null) return [math];
      }
      ctx.warn("math-unconverted", `inline math was not converted to OMML: ${node.value}`);
      const value = safeText(node.value, ctx);
      return [new TextRun({ text: value, ...props, style: STYLE_IDS.codeChar })];
    }

    default:
      return assertNever(node, "inline node");
  }
}

/**
 * Percent-decodes an `#anchor` before it is looked up.
 *
 * Not cosmetic: markdown-it normalises every destination through `mdurl.encode`,
 * so `[jump](#café-setup)` arrives as `#caf%C3%A9-setup` while the heading's
 * slug — which `slugify` builds from `\p{L}`, deliberately keeping every script
 * — is still `café-setup`. Without this, *no* internal link to a heading
 * containing a non-ASCII letter ever resolved, and every such link silently
 * degraded to plain text with a `link-unresolved` notice blaming the author.
 *
 * A malformed escape (`%zz`, a lone `%`) makes `decodeURIComponent` throw; the
 * raw text is the better guess then, and a bookmark named after it simply will
 * not be found.
 */
function decodeAnchor(anchor: string): string {
  if (!anchor.includes("%")) return anchor;
  try {
    return decodeURIComponent(anchor);
  } catch {
    return anchor;
  }
}

/** Where a link mark actually points, once resolved against the document. */
type LinkTarget =
  | { readonly kind: "external"; readonly href: string }
  | { readonly kind: "internal"; readonly anchor: string };

/**
 * Resolves a link mark, or returns `null` if it cannot become a hyperlink.
 *
 * Done *before* the segment's runs are built so that a dead link produces plain
 * text rather than blue underlined text that does nothing when clicked.
 */
function resolveLinkTarget(link: LinkMark | null, ctx: RenderContext): LinkTarget | null {
  if (link === null) return null;

  if (link.title !== null && link.title !== "") {
    // OOXML *does* have `w:hyperlink/@w:tooltip`, but docx's `ExternalHyperlink`
    // and `InternalHyperlink` expose no way to set it, so the title parses,
    // survives in the model, and then has nowhere to go. Small, but not silent.
    ctx.warn(
      "link-title-dropped",
      `the link title ${JSON.stringify(link.title)} was not rendered: docx cannot write w:hyperlink/@w:tooltip`,
    );
  }

  const href = link.href.trim();
  // `[text]()` - an empty destination would mint a relationship pointing at
  // nothing, which some readers treat as a damaged part.
  if (href === "") return null;

  if (!href.startsWith("#")) return { kind: "external", href: link.href };

  const anchor = ctx.bookmarks.get(decodeAnchor(href.slice(1)));
  if (anchor === undefined) {
    ctx.warn(
      "link-unresolved",
      `internal link "${link.href}" does not match any heading; rendered as plain text`,
    );
    return null;
  }
  // `w:anchor` resolves against the bookmark's *name*, not its numeric id.
  return { kind: "internal", anchor: anchor.name };
}

/**
 * Renders inline content, coalescing adjacent equal links.
 *
 * `availableWidth` and `availableHeight` are the containing text column's
 * extents in twips; together they cap embedded images so a picture can spill
 * neither past the right margin nor off the bottom of the page. Inside a table
 * cell the width is the cell's; the height stays the page's, because a row
 * grows to fit its content but a page does not.
 */
export function renderInline(
  nodes: readonly InlineNode[],
  ctx: RenderContext,
  availableWidth: number,
  availableHeight: number = ctx.page.contentHeight,
): ParagraphChild[] {
  const out: ParagraphChild[] = [];

  for (const segment of segmentByLink(nodes)) {
    const target = resolveLinkTarget(segment.link, ctx);

    const children: ParagraphChild[] = [];
    for (const node of segment.nodes) {
      children.push(
        ...renderInlineNode(node, ctx, availableWidth, availableHeight, target === null),
      );
    }
    if (children.length === 0) continue;

    if (target === null) out.push(...children);
    else if (target.kind === "internal")
      out.push(new InternalHyperlink({ anchor: target.anchor, children }));
    else out.push(new ExternalHyperlink({ children, link: target.href }));
  }

  return out;
}
