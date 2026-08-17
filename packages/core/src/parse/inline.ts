import type { Token } from "markdown-it";

import {
  bold,
  footnoteReference,
  hardBreak,
  highlight,
  htmlInline,
  image,
  inlineCode,
  italic,
  link,
  mathInline,
  softBreak,
  strikethrough,
  subscript,
  superscript,
  text,
  type InlineNode,
  type Mark,
  type MarkType,
} from "../model.js";
import type { ParseContext } from "./context.js";

/**
 * Reads a `string` attribute off a token.
 *
 * markdown-it 15 types attribute values as `string | number` (an ordered list's
 * `start` really does arrive as a number), so every read has to be widened.
 *
 * @returns The value as a string, or `null` when the attribute is absent.
 */
export function attrString(token: Token, name: string): string | null {
  const value = token.attrGet(name);
  if (value === null) return null;
  return typeof value === "string" ? value : String(value);
}

/**
 * Flattens inline tokens to plain text, the way an `alt` attribute needs.
 *
 * Only the characters a reader would see survive: emphasis markers, link
 * destinations and raw HTML contribute nothing, and both break kinds collapse
 * to a single space (a newline inside an `alt` string is never wanted).
 */
function flattenInlineText(tokens: readonly Token[]): string {
  let out = "";

  for (const token of tokens) {
    switch (token.type) {
      case "text":
      case "code_inline":
      case "math_inline":
      case "math_inline_double":
        out += token.content;
        break;
      case "image":
        out += altTextOf(token);
        break;
      case "softbreak":
      case "hardbreak":
        out += " ";
        break;
      default:
        break;
    }
  }

  return out;
}

/**
 * The alt text of an `image` token.
 *
 * markdown-it always emits `alt=""` in `token.attrs` and keeps the real alt in
 * the token's parsed `children` (so that `![**bold** alt](x)` can be rendered
 * with markup stripped). `token.content` holds the raw source as a fallback for
 * the degenerate case of an image with no children.
 */
function altTextOf(token: Token): string {
  const children = token.children;
  if (children === null || children.length === 0) return token.content;
  return flattenInlineText(children);
}

/** Mark types that {@link parseInlineTokens} pushes and pops as a stack. */
const OPENERS: Readonly<Record<string, MarkType>> = {
  strong_open: "bold",
  em_open: "italic",
  s_open: "strikethrough",
  link_open: "link",
  sup_open: "superscript",
  sub_open: "subscript",
  mark_open: "highlight",
};

/** The closing counterpart of every entry in {@link OPENERS}. */
const CLOSERS: Readonly<Record<string, MarkType>> = {
  strong_close: "bold",
  em_close: "italic",
  s_close: "strikethrough",
  link_close: "link",
  sup_close: "superscript",
  sub_close: "subscript",
  mark_close: "highlight",
};

/**
 * Converts the flat children of one markdown-it `inline` token into model
 * inline nodes.
 *
 * markdown-it does **not** nest inline markup: `**[a](b)**` arrives as the flat
 * sequence `strong_open, link_open, text, link_close, strong_close`. Because
 * the model puts formatting in `marks` on every inline node rather than in
 * wrapper nodes, rebuilding it needs nothing more than a stack of the currently
 * open marks — no tree surgery, and `[**`npm i`**](url)` falls out as a single
 * `TextNode` carrying `link + bold + inlineCode`.
 *
 * Every builder call normalises the stack (dedup by type, canonical order), so
 * the same document always produces byte-identical output.
 *
 * @param tokens - `inlineToken.children`.
 * @param ctx - Slugger + warning sink.
 * @returns Inline nodes in document order. Empty text runs are dropped.
 */
export function parseInlineTokens(tokens: readonly Token[], ctx: ParseContext): InlineNode[] {
  const out: InlineNode[] = [];
  const marks: Mark[] = [];

  const closeMark = (type: MarkType, token: Token): void => {
    for (let index = marks.length - 1; index >= 0; index -= 1) {
      if (marks[index]?.type === type) {
        marks.splice(index, 1);
        return;
      }
    }
    ctx.warn("unexpected-token", `unbalanced "${token.type}" with no matching opener`, token);
  };

  for (const token of tokens) {
    const opener = OPENERS[token.type];
    if (opener !== undefined) {
      marks.push(openMark(opener, token));
      continue;
    }

    const closer = CLOSERS[token.type];
    if (closer !== undefined) {
      closeMark(closer, token);
      continue;
    }

    switch (token.type) {
      case "text":
        // `***x***` yields zero-length text tokens around the nested emphasis;
        // an empty run has no representation in OOXML, so drop it.
        if (token.content.length > 0) out.push(text(token.content, marks));
        break;

      case "code_inline":
        out.push(text(token.content, [...marks, inlineCode()]));
        break;

      // markdown-it's `breaks` option only changes its HTML renderer, which
      // this parser never runs; applying it here is what makes it observable.
      case "softbreak":
        out.push(ctx.hardBreaks ? hardBreak(marks) : softBreak(marks));
        break;

      case "hardbreak":
        out.push(hardBreak(marks));
        break;

      case "image":
        out.push(
          image(attrString(token, "src") ?? "", {
            alt: altTextOf(token),
            title: attrString(token, "title"),
            marks,
          }),
        );
        break;

      case "html_inline":
        ctx.warn("raw-html", `raw inline HTML: ${summarise(token.content)}`, token);
        out.push(htmlInline(token.content, marks));
        break;

      case "footnote_ref": {
        const reference = readFootnoteIdentity(token);
        if (reference === null) {
          ctx.warn("unexpected-token", "footnote reference without a usable id", token);
          break;
        }
        out.push(
          footnoteReference(reference.identifier, reference.number, {
            label: reference.label,
            marks,
          }),
        );
        break;
      }

      // Emitted by markdown-it-footnote at the end of a definition to link back
      // to the reference. It is presentation, not content.
      case "footnote_anchor":
        break;

      case "math_inline":
      case "math_inline_double":
        out.push(mathInline(token.content, { marks }));
        break;

      default:
        ctx.warn("unsupported-token", `unsupported inline token "${token.type}"`, token);
        break;
    }
  }

  if (marks.length > 0) {
    ctx.warn(
      "unexpected-token",
      `inline content ended with ${String(marks.length)} unclosed mark(s)`,
    );
  }

  return out;
}

/** Builds the mark for an opening token, reading a link's href/title from it. */
function openMark(type: MarkType, token: Token): Mark {
  switch (type) {
    case "link":
      return link(attrString(token, "href") ?? "", attrString(token, "title"));
    case "bold":
      return bold();
    case "italic":
      return italic();
    case "strikethrough":
      return strikethrough();
    case "superscript":
      return superscript();
    case "subscript":
      return subscript();
    case "highlight":
      return highlight();
    case "inlineCode":
      return inlineCode();
  }
}

/** Identity of a footnote as carried by a markdown-it-footnote token. */
export interface FootnoteIdentity {
  readonly identifier: string;
  readonly label: string;
  /** 1-based; markdown-it-footnote's `meta.id` is 0-based, docx ids must be >= 1. */
  readonly number: number;
}

/**
 * Reads `{ id, label }` off a `footnote_ref` / `footnote_open` token.
 *
 * `meta` is typed `Record<string, unknown>`, so every field needs a runtime
 * check. Inline footnotes (`^[note]`) carry **no** `label` at all — they were
 * never labelled in the source — so the rendered number stands in for one.
 *
 * @returns The identity, or `null` if the token carries no usable id.
 */
export function readFootnoteIdentity(token: Token): FootnoteIdentity | null {
  const meta = token.meta;
  if (meta === null) return null;

  const id = meta["id"];
  if (typeof id !== "number" || !Number.isInteger(id) || id < 0) return null;

  const rawLabel = meta["label"];
  const label = typeof rawLabel === "string" && rawLabel.length > 0 ? rawLabel : String(id + 1);

  return { identifier: label, label, number: id + 1 };
}

/** Truncates a raw source fragment so warnings stay one line long. */
export function summarise(value: string): string {
  const collapsed = value.replace(/\s+/g, " ").trim();
  return collapsed.length > 60 ? `${collapsed.slice(0, 57)}...` : collapsed;
}
