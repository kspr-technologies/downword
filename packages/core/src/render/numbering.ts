/**
 * `word/numbering.xml` generation.
 *
 * ## Why not `bullet: { level }`
 *
 * docx ships a convenience `bullet` option, and it is a trap for a markdown
 * renderer. It hard-codes `numId 1` — a single concrete numbering instance for
 * the entire document — with fixed glyphs, fixed indents and no way to restart.
 * Everything below exists because a markdown document needs the opposite of all
 * four.
 *
 * ## The model
 *
 * OOXML numbering is two-layered:
 *
 *  - `<w:abstractNum>` — the *shape* of a list: nine levels, each with a format
 *    (`decimal`, `bullet`, …), a level text (`%2.`), an indent and a start
 *    value. docx calls this a `reference`.
 *  - `<w:num>` — a concrete *instance* pointing at an abstract, optionally with
 *    `<w:lvlOverride w:ilvl="0"><w:startOverride>`. docx derives one per
 *    distinct `${reference}-${instance}` pair it sees on a paragraph.
 *
 * Three consequences drive the design here, all verified against docx 9.7.1:
 *
 *  1. **Sibling lists must not continue each other.** Two adjacent `1. 2. 3.`
 *     lists that share a `numId` render as `1..3` then `4..6`. Every *top-level*
 *     list therefore takes a fresh `instance`, which docx turns into its own
 *     `<w:num>` with `startOverride`. Nested lists deliberately reuse their
 *     ancestor's instance, since that is what makes level 1 restart under each
 *     level-0 item.
 *  2. **`startOverride` is only ever written for `ilvl="0"`, and its value
 *     comes from the abstract's level 0.** So `3. 4. 5.` (a markdown `start`)
 *     cannot be expressed by an instance override — it needs a *separate
 *     abstract* whose level 0 declares `start: 3`. A nested list with a custom
 *     start needs one whose level *N* declares it.
 *  3. Unreferenced entries in `numbering.config` still emit an `<w:abstractNum>`,
 *     so references are created lazily and only when a paragraph asks for one.
 *
 * Every reference defines all nine levels with the same indent ladder
 * (`720 * (level + 1)` twips, `360` hanging), which is what makes a bullet list
 * nested inside an ordered list line up with its parent instead of drifting.
 */

import { AlignmentType, LevelFormat, type ILevelsOptions, type INumberingOptions } from "docx";

import type { BulletLevel, LevelFormatValue, Theme } from "./theme.js";

/** OOXML defines exactly nine list levels, `w:ilvl` 0-8. */
export const LIST_LEVEL_COUNT = 9;

/** The deepest level a list may nest to. */
export const MAX_LIST_LEVEL = LIST_LEVEL_COUNT - 1;

/** Reference names for the fixed (non start-dependent) list shapes. */
export const NUMBERING_REFERENCES = {
  bullet: "downword-bullet",
  taskChecked: "downword-task-checked",
  taskUnchecked: "downword-task-unchecked",
  taskPlain: "downword-task-plain",
  /** Ordered lists that start at 1 — the overwhelming majority. */
  ordered: "downword-ordered",
} as const;

/** Builds the reference name for an ordered list with a non-default start. */
function orderedReferenceName(start: number, level: number): string {
  return `${NUMBERING_REFERENCES.ordered}-s${start}-l${level}`;
}

/** Collects the numbering definitions a single render actually used. */
export interface NumberingRegistry {
  /**
   * Allocates a fresh list instance. Called once per *top-level* list that
   * contains an ordered list anywhere in it; nested lists inherit their
   * ancestor's value.
   */
  readonly nextInstance: () => number;
  /**
   * The one instance shared by every top-level list with no ordered list
   * anywhere inside it.
   *
   * A fresh instance exists to restart a *counter*, and bullets and task
   * checkboxes have none: two adjacent bullet lists sharing a `numId` are
   * indistinguishable from two that do not. Handing them all the same instance
   * is therefore invisible in the output — and it matters, because docx's
   * packer resolves numbering placeholders by running one whole-document
   * `String.replace` **per concrete numbering** (`NumberingReplacer`). One
   * instance per list makes packing quadratic in document size: a 1 MB
   * bullet-heavy paste minted ~26,000 of them and spent 66 s in that loop.
   *
   * Allocated lazily so a document without any such list does not reserve one.
   */
  readonly sharedInstance: () => number;
  /** Reference for a bulleted list. */
  readonly bullet: () => string;
  /** Reference for one task-list item, chosen by its checkbox state. */
  readonly task: (checked: boolean | null) => string;
  /**
   * Reference for an ordered list starting at `start`, used at depth `level`.
   *
   * `level` matters because `startOverride` only ever applies to level 0: a
   * nested list that starts at 5 needs an abstract whose *level 5* declares it.
   */
  readonly ordered: (start: number, level: number) => string;
  /**
   * The `numbering.config` payload, or `null` when the document contains no
   * lists at all (in which case no `numbering.xml` should be emitted).
   */
  readonly build: () => INumberingOptions | null;
}

/** The `style` payload of a numbering level; non-optional so it can be spread. */
type LevelStyle = NonNullable<ILevelsOptions["style"]>;

function levelIndent(theme: Theme, level: number): LevelStyle {
  return {
    paragraph: {
      indent: {
        left: theme.spacing.listIndent * (level + 1),
        hanging: theme.spacing.listHanging,
      },
    },
  };
}

function levels(count: number): readonly number[] {
  return Array.from({ length: count }, (_unused, index) => index);
}

function bulletLevels(theme: Theme, glyph: BulletLevel | null): readonly ILevelsOptions[] {
  return levels(LIST_LEVEL_COUNT).map((level) => {
    const rung = glyph ??
      theme.bulletLevels[level % theme.bulletLevels.length] ?? { text: "•", font: null };
    const indent = levelIndent(theme, level);
    return {
      level,
      format: LevelFormat.BULLET,
      text: rung.text,
      alignment: AlignmentType.LEFT,
      style: rung.font === null ? indent : { ...indent, run: { font: rung.font } },
    };
  });
}

function orderedLevels(theme: Theme, startLevel: number, start: number): readonly ILevelsOptions[] {
  return levels(LIST_LEVEL_COUNT).map((level) => {
    const format: LevelFormatValue =
      theme.orderedFormats[level % theme.orderedFormats.length] ?? LevelFormat.DECIMAL;
    // Word right-aligns roman numerals so `viii.` and `i.` share a right edge.
    const roman = format === LevelFormat.LOWER_ROMAN || format === LevelFormat.UPPER_ROMAN;
    return {
      level,
      format,
      // `%N` interpolates level N-1, so a level-2 list numbers itself with %3.
      text: `%${level + 1}.`,
      alignment: roman ? AlignmentType.RIGHT : AlignmentType.LEFT,
      start: level === startLevel ? start : 1,
      style: levelIndent(theme, level),
    };
  });
}

/** Creates an empty registry bound to a theme. */
export function createNumberingRegistry(theme: Theme): NumberingRegistry {
  const configs = new Map<string, readonly ILevelsOptions[]>();
  let instance = 0;
  let shared: number | null = null;

  const ensure = (reference: string, create: () => readonly ILevelsOptions[]): string => {
    if (!configs.has(reference)) configs.set(reference, create());
    return reference;
  };

  const nextInstance = (): number => {
    instance += 1;
    return instance;
  };

  return {
    nextInstance,

    sharedInstance: () => {
      shared ??= nextInstance();
      return shared;
    },

    bullet: () => ensure(NUMBERING_REFERENCES.bullet, () => bulletLevels(theme, null)),

    task: (checked) => {
      const glyph =
        checked === true
          ? { text: theme.taskGlyphs.checked, font: theme.fonts.symbol }
          : checked === false
            ? { text: theme.taskGlyphs.unchecked, font: theme.fonts.symbol }
            : { text: theme.taskGlyphs.plain, font: theme.fonts.symbol };
      const reference =
        checked === true
          ? NUMBERING_REFERENCES.taskChecked
          : checked === false
            ? NUMBERING_REFERENCES.taskUnchecked
            : NUMBERING_REFERENCES.taskPlain;
      return ensure(reference, () => bulletLevels(theme, glyph));
    },

    ordered: (start, level) => {
      const safeStart = Number.isFinite(start) ? Math.max(0, Math.trunc(start)) : 1;
      const safeLevel = Math.min(Math.max(0, Math.trunc(level)), MAX_LIST_LEVEL);
      if (safeStart === 1) {
        return ensure(NUMBERING_REFERENCES.ordered, () => orderedLevels(theme, 0, 1));
      }
      return ensure(orderedReferenceName(safeStart, safeLevel), () =>
        orderedLevels(theme, safeLevel, safeStart),
      );
    },

    build: () => {
      if (configs.size === 0) return null;
      // Sorted so the abstract ids in numbering.xml depend only on *which*
      // references were used, never on the order the document happened to use
      // them in. Golden snapshots stay stable when a block is moved.
      const config = [...configs.entries()]
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
        .map(([reference, levelOptions]) => ({ reference, levels: levelOptions }));
      return { config };
    },
  };
}
