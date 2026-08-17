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
