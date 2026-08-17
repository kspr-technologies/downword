/**
 * The one vocabulary every stage's diagnostics share.
 *
 * downword has four places that can discover a problem — the parser, the image
 * pass, the two async prepare passes and the renderer — and exactly one place a
 * host wants to read them: {@link import("./options.js").ConvertOptions.onWarning}.
 * For that to be usable rather than merely populated, every warning any stage
 * raises has to answer the same three questions in the same shape:
 *
 *  - **What kind of problem is it?** a `code`, from a small closed set per
 *    stage, so a host can filter, group and count without matching prose.
 *  - **How much does it matter?** a {@link WarningSeverity}, so a host can say
 *    "2 problems, 5 notes" instead of listing eighteen equal-looking lines.
 *  - **What do I tell the user?** a `message`, already naming the offending
 *    construct.
 *
 * Stages then add whatever *they* alone know: a parse warning carries a source
 * line, an image diagnostic carries the `src` that failed. Nothing else varies.
 *
 * Severity is a property of the code, not of the moment — every stage keeps a
 * total `Record<Code, WarningSeverity>` table rather than deciding at the call
 * site, so the answer cannot drift between two places that raise the same code.
 */

/**
 * How much a warning matters.
 *
 * The line between the two is *whether the document lost anything*:
 *
 * - `"error"` — content the source contained is not in the `.docx`, or was
 *   replaced by a placeholder. Something is missing; a host should say so.
 * - `"notice"` — the content is all there, but downword had to change how it is
 *   expressed (a construct OOXML cannot hold, a limit it had to clamp to, an
 *   attribute Word has no field for). Worth logging, not worth alarming anyone.
 *
 * A conversion never fails because of either: see `DownwordError` for the
 * things that actually stop it.
 */
export type WarningSeverity = "error" | "notice";

/** The shape every stage's warning starts from. See {@link WarningSeverity}. */
export interface BaseWarning {
  /** Machine-readable kind, from the raising stage's own closed set. */
  readonly code: string;
  /** How much it matters. Derived from `code`, never decided per occurrence. */
  readonly severity: WarningSeverity;
  /** Human-readable, already naming the offending construct. Safe to show a user. */
  readonly message: string;
}
