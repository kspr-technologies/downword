import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync } from "node:fs";
import { writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import ts from "typescript";
import { beforeAll, describe, expect, it } from "vitest";

/**
 * Documentation doctest.
 *
 * Every `ts`/`js` code sample in the READMEs is extracted, written to disk and
 * type-checked **against the built `.d.ts` files** — the same declarations a
 * consumer installs, resolved through the package's real `exports` map. A
 * sample that drifts from the API (a renamed option, a changed return type, an
 * export that moved to another subpath) stops compiling, and CI goes red.
 *
 * This is the only test that can catch that class of bug. Unit tests import
 * `src/`, so they keep passing while the published types say something else,
 * and prose is invisible to every other kind of check.
 *
 * ### Two deliberate relaxations
 *
 * `noUnusedLocals` and `noUnusedParameters` are **off**. A README sample ends
 * at the interesting line — `const bytes = await convert(md);` — and padding
 * every one with a `console.log` to satisfy a style rule would make the docs
 * worse to read without making the API any more correct. Everything that can
 * catch drift (`strict`, `exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`)
 * stays on.
 */

const PACKAGE_DIR = fileURLToPath(new URL("..", import.meta.url));
const REPO_ROOT = join(PACKAGE_DIR, "..", "..");
const SRC_DIR = join(PACKAGE_DIR, "src");
const DIST_ENTRY = join(PACKAGE_DIR, "dist", "index.d.ts");

/** Scratch directory for the extracted samples. Inside node_modules, so nothing ignores it twice. */
const WORK_DIR = join(PACKAGE_DIR, "node_modules", ".readme-doctest");

/** The READMEs under test: the published package one, and the repository front page. */
const READMES = [
  { label: "packages/core/README.md", path: join(PACKAGE_DIR, "README.md") },
  { label: "README.md", path: join(REPO_ROOT, "README.md") },
] as const;

/** Fence info strings that mean "this is code that must compile". */
const CHECKED_LANGUAGES = new Map<string, "ts" | "js">([
  ["ts", "ts"],
  ["tsx", "ts"],
  ["typescript", "ts"],
  ["js", "js"],
  ["jsx", "js"],
  ["javascript", "js"],
]);

interface Sample {
  readonly readme: string;
  /** 1-based line of the opening fence, so a failure points at the README. */
  readonly line: number;
  readonly kind: "ts" | "js";
  readonly code: string;
  /** Filename the sample is written to. */
  readonly file: string;
}

/**
 * Extracts fenced code blocks.
 *
 * Line-based rather than one big regular expression, because a fence can only
 * be closed by a run of at least as many backticks as opened it — which is how
 * a markdown sample containing a nested fence stays one block.
 */
function extractSamples(readme: string, markdown: string): Sample[] {
  const lines = markdown.split("\n");
  const samples: Sample[] = [];
  let index = 0;

  while (index < lines.length) {
    const opening = /^(\s*)(`{3,})\s*([^\s`]*)/.exec(lines[index] ?? "");
    if (opening === null) {
      index += 1;
      continue;
    }

    const [, indent = "", fence = "```", info = ""] = opening;
    const closing = new RegExp(`^\\s*\`{${fence.length},}\\s*$`);
    const startLine = index + 1;
    const body: string[] = [];

    index += 1;
    while (index < lines.length && !closing.test(lines[index] ?? "")) {
      body.push(
        (lines[index] ?? "").startsWith(indent)
          ? (lines[index] ?? "").slice(indent.length)
          : (lines[index] ?? ""),
      );
      index += 1;
    }
    index += 1;

    const kind = CHECKED_LANGUAGES.get(info.toLowerCase());
    if (kind === undefined) continue;

    samples.push({
      readme,
      line: startLine,
      kind,
      code: `${body.join("\n")}\n`,
      file: `${readme.replace(/[^a-z0-9]+/gi, "-")}-${String(startLine).padStart(4, "0")}.${kind}`,
    });
  }

  return samples;
}

/** Newest mtime under a directory tree. */
function newestMtime(dir: string): number {
  let newest = 0;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    newest = Math.max(newest, entry.isDirectory() ? newestMtime(full) : statSync(full).mtimeMs);
  }
  return newest;
}

interface PackageExports {
  readonly exports: Readonly<Record<string, { readonly import?: { readonly types?: string } }>>;
}

/**
 * Maps every subpath in the package's `exports` to its built `.d.ts`.
 *
 * Derived from `package.json` rather than hard-coded, so a new subpath export
 * is checkable the moment it is declared — and a subpath whose `types` path is
 * a lie fails here rather than in a user's editor.
 */
function declarationPaths(): Record<string, string[]> {
  const pkg = JSON.parse(readFileSync(join(PACKAGE_DIR, "package.json"), "utf8")) as PackageExports;

  const paths: Record<string, string[]> = {};
  for (const [subpath, conditions] of Object.entries(pkg.exports)) {
    if (subpath === "./package.json") continue;
    const types = conditions.import?.types;
    if (types === undefined) throw new Error(`exports["${subpath}"] declares no import types`);
    const specifier = subpath === "." ? "downword" : `downword/${subpath.slice(2)}`;
    paths[specifier] = [join(PACKAGE_DIR, types)];
  }
  return paths;
}

let samples: Sample[] = [];
let diagnostics: readonly ts.Diagnostic[] = [];

beforeAll(() => {
  // CI runs `test` before `build`, so the declarations this test checks
  // against may not exist yet. Build on demand, and rebuild when src is newer.
  if (!existsSync(DIST_ENTRY) || statSync(DIST_ENTRY).mtimeMs < newestMtime(SRC_DIR)) {
    execFileSync("pnpm", ["run", "build"], { cwd: PACKAGE_DIR, stdio: "inherit" });
  }

  samples = READMES.flatMap(({ label, path }) => extractSamples(label, readFileSync(path, "utf8")));

  rmSync(WORK_DIR, { recursive: true, force: true });
  mkdirSync(WORK_DIR, { recursive: true });
  for (const sample of samples) {
    writeFileSync(join(WORK_DIR, sample.file), sample.code);
  }

  const program = ts.createProgram({
    rootNames: samples.map((sample) => join(WORK_DIR, sample.file)),
    options: {
      target: ts.ScriptTarget.ES2022,
      lib: ["lib.es2022.d.ts", "lib.dom.d.ts", "lib.dom.iterable.d.ts"],
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
      moduleDetection: ts.ModuleDetectionKind.Force,
      strict: true,
      noUncheckedIndexedAccess: true,
      exactOptionalPropertyTypes: true,
      noImplicitOverride: true,
      noImplicitReturns: true,
      noFallthroughCasesInSwitch: true,
      useUnknownInCatchVariables: true,
      esModuleInterop: true,
      allowSyntheticDefaultImports: true,
      resolveJsonModule: true,
      allowJs: true,
      checkJs: true,
      skipLibCheck: true,
      noEmit: true,
      types: ["node"],
      typeRoots: [join(PACKAGE_DIR, "node_modules", "@types")],
      baseUrl: PACKAGE_DIR,
      paths: declarationPaths(),
    },
  });

  diagnostics = ts.getPreEmitDiagnostics(program);
}, 300_000);

/** Turns a diagnostic into `README.md:42 (sample line 3): TS2551: …`. */
function describeDiagnostic(diagnostic: ts.Diagnostic): string {
  const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, " ");
  const file = diagnostic.file;
  if (file === undefined) return `TS${diagnostic.code}: ${message}`;

  const name = relative(WORK_DIR, file.fileName);
  const sample = samples.find((candidate) => candidate.file === name);
  const offset = diagnostic.start ?? 0;
  const { line, character } = file.getLineAndCharacterOfPosition(offset);
  const where =
    sample === undefined
      ? name
      : `${sample.readme}:${sample.line + line + 1} (sample line ${line + 1}, col ${character + 1})`;

  return `${where}: TS${diagnostic.code}: ${message}`;
}

describe("README samples", () => {
  it("finds samples in every README", () => {
    for (const { label } of READMES) {
      expect(
        samples.filter((sample) => sample.readme === label).length,
        `${label} has no ts/js code samples - did the extractor break?`,
      ).toBeGreaterThan(0);
    }
    expect(samples.length).toBeGreaterThanOrEqual(10);
  });

  it("every subpath in the exports map has declarations on disk", () => {
    for (const [specifier, [declaration]] of Object.entries(declarationPaths())) {
      expect(declaration, `${specifier} has no declaration path`).toBeDefined();
      expect(existsSync(declaration as string), `${specifier} -> ${declaration} is missing`).toBe(
        true,
      );
    }
  });

  it("type-check against the built .d.ts without a single error", () => {
    expect(diagnostics.map(describeDiagnostic)).toEqual([]);
  });

  it("resolve `downword` to the built declarations, not to src", () => {
    // A `paths` entry pointing at src would make this test pass while the
    // published types were broken, so assert the mapping itself.
    for (const targets of Object.values(declarationPaths())) {
      expect(targets[0]).toMatch(/[/\\]dist[/\\].*\.d\.ts$/);
    }
  });
});

/**
 * Relative links in the READMEs point at files that exist.
 *
 * A README is a landing page, and a 404 on it costs more credibility than a
 * typo in the prose. Nothing else in the suite reads the prose, so a renamed
 * `docs/` file or a moved package would otherwise go unnoticed until a reader
 * clicked it.
 *
 * Scope is deliberately narrow: inline `[text](target)` links to paths on disk.
 * Absolute URLs are somebody else's uptime, and `#anchor` fragments would need
 * this test to reimplement GitHub's heading-slug rules — a second, subtly
 * different implementation of which is a worse bug than the one it catches.
 */
describe("README links", () => {
  /** `[text](target)`, ignoring the `!` of an image and any `"title"` suffix. */
  const INLINE_LINK = /\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g;

  /**
   * Removes everything that *looks* like a link but is not one.
   *
   * Three sources of false positives, all of them present in these files: an
   * HTML comment (the hero-GIF TODO names a path that does not exist yet, on
   * purpose), a fenced block (sample code contains markdown), and an inline
   * code span (the conversion table documents `[text](url)` as syntax).
   *
   * Code spans are stripped **per line**, not across the whole file. Pairing
   * backtick runs document-wide drifts the moment one line contains an odd
   * number of them — ``` | `` `code` ``, ` ```lang ` | ``` does — after which a
   * single "span" swallows the rest of the file and every real link with it.
   * Fenced blocks are already gone by then, so nothing legitimate spans lines.
   */
  function stripNonLinks(markdown: string): string {
    return markdown
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/^(\s*)(`{3,})[\s\S]*?^\s*\2\s*$/gm, "")
      .split("\n")
      .map((line) => line.replace(/(`+)[^`]*?\1/g, ""))
      .join("\n");
  }

  for (const { label, path } of READMES) {
    it(`${label} has no broken relative links`, () => {
      const markdown = stripNonLinks(readFileSync(path, "utf8"));
      const base = dirname(path);
      const broken: string[] = [];

      for (const [, target] of markdown.matchAll(INLINE_LINK)) {
        if (target === undefined) continue;
        if (/^(?:[a-z][a-z0-9+.-]*:|\/\/|#)/i.test(target)) continue;

        const [pathPart = ""] = target.split("#");
        if (pathPart === "") continue;

        const resolved = resolve(base, decodeURIComponent(pathPart));
        if (!existsSync(resolved)) broken.push(`${target} -> ${resolved}`);
      }

      expect(broken, `${label} links to files that do not exist`).toEqual([]);
    });
  }
});
