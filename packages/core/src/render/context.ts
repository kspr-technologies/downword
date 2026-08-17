/**
 * Render-time state.
 *
 * Split in two on purpose:
 *
 *  - {@link RenderContext} is created once per `renderDocument()` call and is
 *    the same object everywhere. It owns the things that must be shared —
 *    the numbering registry, the bookmark table, the warning sink.
 *  - {@link BlockContext} is *positional* and immutable: it describes where in
 *    the tree the current block sits (how deep in quotes, how deep in lists,
 *    how much left indent it has inherited). Descending into a container makes
 *    a new one with `{ ...parent, … }` rather than mutating, so nothing can
 *    leak back up when a branch finishes rendering.
 */

import type { FootnoteDefinitionNode } from "../model.js";
import type { Theme } from "./theme.js";
import type { NumberingRegistry } from "./numbering.js";
import type {
  HighlightMap,
  Highlighter,
  ImageMap,
  PageSettings,
  RenderWarningCode,
  ResolvedRenderOptions,
} from "./types.js";

/**
 * One heading's anchor: the name links point at, and the numeric id that pairs
 * its `<w:bookmarkStart>` with its `<w:bookmarkEnd>`.
 *
 * The two are different things and OOXML needs both. `w:name` is what
 * `<w:hyperlink w:anchor="…">` resolves against; `w:id` is `CT_Bookmark`'s
 * *identifier*, and §17.13.6.2 requires it to be unique within the part —
 * it is the key Word's Bookmarks dialog, cross-references and `TOC` fields are
 * built on. docx's own `Bookmark` mints a fresh counter per instance and so
 * writes `w:id="1"` for every bookmark in the document; see `renderHeading`.
 */
export interface BookmarkAnchor {
  /** Legal Word bookmark name; see `toBookmarkName`. */
  readonly name: string;
  /** Unique within `word/document.xml`. Allocated by the pre-pass, in order. */
  readonly id: number;
}

/**
 * Everything the renderer needs to know about the document's footnotes.
 *
 * Built once, before anything is rendered, because a `[^x]` marker has to know
 * three things the node itself cannot tell it: whether a definition for it
 * exists at all (a dangling reference must not become a `<w:footnoteReference>`
 * pointing at a footnote that is not in the part), what that definition *says*
 * (so `footnotes: false` can splice it into the sentence), and whether a
 * footnote body is being rendered right now.
 *
 * That last one is what makes nesting safe. OOXML's content model permits a
 * `<w:footnoteReference>` inside `word/footnotes.xml`, but Word cannot author or
 * display a footnote inside a footnote, so downword never writes one: a `[^y]`
 * met while footnote `x` is open is spliced into `x`'s text in parentheses
 * instead — the same degradation `footnotes: false` applies everywhere. The
 * open set doubles as the cycle guard for `[^a]` and `[^b]` that cite each
 * other, which would otherwise recurse until the stack ran out.
 */
export interface FootnoteIndex {
  /**
   * Definition by its docx footnote id, i.e. {@link FootnoteDefinitionNode.number}.
   *
   * Only definitions at the document root are here — one nested inside another
   * block is a `footnote-misplaced` warning, not a note.
   */
  readonly byNumber: ReadonlyMap<number, FootnoteDefinitionNode>;
  /** How many footnote bodies are being rendered right now; 0 in the body. */
  readonly depth: () => number;
  /**
   * Marks `number` as open, or returns `false` if it already was — which only
   * happens when a footnote cites itself, directly or through a cycle.
   */
  readonly open: (number: number) => boolean;
  /** Undoes one successful {@link FootnoteIndex.open}. */
  readonly close: (number: number) => void;
  /**
   * Records that a real `<w:footnoteReference w:id="N"/>` was written.
   *
   * This is what decides whether the note goes into `word/footnotes.xml` at
   * all. A `<w:footnote>` nothing points at is one Word never draws and drops
   * on the next save, so downword writes the part and the body as a matched
   * pair: every id in one is an id in the other, in both directions.
   */
  readonly markReferenced: (number: number) => void;
  /** Whether {@link FootnoteIndex.markReferenced} was called for `number`. */
  readonly wasReferenced: (number: number) => boolean;
  /** Records that the note's text was spliced into the flow instead. */
  readonly markInlined: (number: number) => void;
  /**
   * Whether the note reached the reader at all, by either route.
   *
   * A definition that is neither referenced nor inlined is one whose words are
   * in no part of the output, which is the whole of `footnote-unreferenced`.
   */
  readonly wasUsed: (number: number) => boolean;
}

/**
 * Builds a {@link FootnoteIndex} over one document's root-level definitions.
 *
 * @param definitions - Root-level definitions, in document order. A duplicate
 *   number keeps the first: `Document({ footnotes })` is keyed by id, so a
 *   second definition claiming the same one could only overwrite it.
 */
export function createFootnoteIndex(definitions: readonly FootnoteDefinitionNode[]): FootnoteIndex {
  const byNumber = new Map<number, FootnoteDefinitionNode>();
  for (const definition of definitions) {
    if (!byNumber.has(definition.number)) byNumber.set(definition.number, definition);
  }

  const openNumbers = new Set<number>();
  const referenced = new Set<number>();
  const inlined = new Set<number>();

  return {
    byNumber,
    depth: () => openNumbers.size,
    open: (number: number) => {
      if (openNumbers.has(number)) return false;
      openNumbers.add(number);
      return true;
    },
    close: (number: number) => {
      openNumbers.delete(number);
    },
    markReferenced: (number: number) => {
      referenced.add(number);
    },
    wasReferenced: (number: number) => referenced.has(number),
    markInlined: (number: number) => {
      inlined.add(number);
    },
    wasUsed: (number: number) => referenced.has(number) || inlined.has(number),
  };
}

/** State shared by the whole render. */
export interface RenderContext {
  readonly theme: Theme;
  readonly options: ResolvedRenderOptions;
  readonly page: PageSettings;
  readonly numbering: NumberingRegistry;
  /**
   * Heading id (the model's slug) -> the anchor it becomes.
   *
   * Built in a pre-pass so that a `[jump](#later-heading)` appearing *before*
   * its target still resolves.
   */
  readonly bookmarks: ReadonlyMap<string, BookmarkAnchor>;
  /**
   * Allocates the next unused `w:id` for a bookmark this render invents.
   *
   * The **same** allocator the heading pre-pass drew
   * {@link RenderContext.bookmarks} from, not a second counter beside it — so a
   * Phase 2 feature that needs its own anchors (a `TOC` field, a `REF`
   * cross-reference) cannot collide with a heading's, and the seam is exercised
   * by every heading rather than only once that feature lands.
   */
  readonly nextBookmarkId: () => number;
  /**
   * Allocates the next unused `wp:docPr/@id` for a drawing.
   *
   * Exactly the {@link RenderContext.nextBookmarkId} story, in a second id
   * space. `wp:docPr/@id` is an `ST_DrawingElementId` that must be unique
   * across the document, and docx 9.7.1 builds its `DocProperties` counter
   * *inside the constructor* (`docPropertiesUniqueNumericIdGen()`), so every
   * drawing it numbers for itself comes out as `id="1"`. Two pictures in one
   * document is enough to hit it, and a `mermaid` pass that turns fences into
   * pictures makes it the common case — so the renderer supplies the id.
   */
  readonly nextDrawingId: () => number;
  /** The document's footnotes; see {@link FootnoteIndex}. */
  readonly footnotes: FootnoteIndex;
  readonly highlights: HighlightMap | null;
  readonly highlighter: Highlighter | null;
  readonly images: ImageMap | null;
  /** Reports a non-fatal problem. Never throws. */
  readonly warn: (code: RenderWarningCode, message: string) => void;
  /**
   * Reports a non-fatal problem **at most once per code** for this document.
   *
   * For degradations that are a property of the document rather than of one
   * node: a 30-level list would otherwise raise twenty identical
   * `indent-clamped` notices, and a Hebrew document one `rtl-not-enabled` per
   * text run. Never throws.
   */
  readonly warnOnce: (code: RenderWarningCode, message: string) => void;
}

/** Where the block currently being rendered sits in the tree. */
export interface BlockContext {
  /**
   * Left indent inherited from enclosing containers, in twips.
   *
   * Zero at the document root, which is the common case and the one where the
   * renderer can stay purely style-driven: a paragraph only carries an explicit
   * `<w:ind>` when something above it actually moved it.
   */
  readonly indent: number;
  /** Block-quote nesting depth; 0 outside any quote. */
  readonly quoteDepth: number;
  /** `w:ilvl` of the enclosing list item, or -1 when not inside a list. */
  readonly listLevel: number;
  /**
   * True nesting depth of the enclosing list item, or -1 outside a list.
   *
   * Equal to {@link BlockContext.listLevel} for the first nine levels and
   * larger past them: OOXML has exactly nine, so `listLevel` saturates while
   * this keeps counting. The difference is what the renderer indents by, so a
   * tenth level still looks nested even though Word calls it the ninth.
   */
  readonly listDepth: number;
  /** Numbering instance of the enclosing top-level list, or -1. */
  readonly listInstance: number;
  /**
   * The numbering *reference* the enclosing ordered list joined, or `null` when
   * no ordered list encloses this one at the immediately preceding level.
   *
   * Reusing an ancestor's instance is only correct when the ancestor sits on
   * the same abstract definition — that is what puts an `ilvl` 0 paragraph of
   * the same `numId` between two sibling sub-lists and makes Word restart the
   * inner counter. See `renderList`.
   */
  readonly listReference: string | null;
  /**
   * True while rendering the body of a footnote, so its paragraphs pick up the
   * built-in `FootnoteText` style instead of `Normal`.
   */
  readonly inFootnote: boolean;
}

/** The block context at the document root. */
export const ROOT_BLOCK_CONTEXT: BlockContext = {
  indent: 0,
  quoteDepth: 0,
  listLevel: -1,
  listDepth: -1,
  listInstance: -1,
  listReference: null,
  inFootnote: false,
};
