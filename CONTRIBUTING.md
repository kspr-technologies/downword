# Contributing to downword

Thanks for helping out. This document describes the commands and layout that
actually exist in this repository today — if something here is wrong, that's a
bug worth reporting.

By participating you agree to the [Code of Conduct](CODE_OF_CONDUCT.md).

---

## 1. Development setup

Requirements:

| Tool        | Version                 | Why                                           |
| ----------- | ----------------------- | --------------------------------------------- |
| Node.js     | `>=20` (22 recommended) | CI tests 20 and 22                            |
| pnpm        | `10.33.0` (pinned)      | `packageManager` field; use `corepack enable` |
| LibreOffice | optional, CI-only       | the `.docx` validity gate (`soffice`)         |

```sh
git clone https://github.com/kspr-technologies/downword.git
cd downword
corepack enable          # picks up pnpm@10.33.0 from package.json
pnpm install
```

> `pnpm size` needs Node **>= 22.18** because `size-limit@13` declares
> `engines.node: ^22.18.0 || ^24 || >=26`. Everything else runs on Node 20. CI
> only runs the size check on the Node 22 matrix leg.

## 2. Commands

All of these run from the repository root.

| Command                 | What it does                                                        |
| ----------------------- | ------------------------------------------------------------------- |
| `pnpm build`            | `tsup` build of every package in `packages/*` (ESM + CJS + `.d.ts`) |
| `pnpm test`             | `vitest run` in every package                                       |
| `pnpm test:watch`       | `vitest` in watch mode, all packages in parallel                    |
| `pnpm typecheck`        | `tsc --noEmit` per package                                          |
| `pnpm lint`             | ESLint flat config over the whole repo                              |
| `pnpm lint:fix`         | …with `--fix`                                                       |
| `pnpm format`           | Prettier write                                                      |
| `pnpm format:check`     | Prettier check (what CI would run)                                  |
| `pnpm size`             | `size-limit` budget for `packages/core`                             |
| `pnpm examples`         | Runs both examples; the dual ESM/CJS smoke test                     |
| `pnpm docx:validity`    | The LibreOffice `.docx` gate (`scripts/docx-validity.mjs`)          |
| `pnpm changeset`        | Records a changeset for the next release                            |
| `pnpm version-packages` | `changeset version` — applies pending changesets (CI only)          |
| `pnpm release`          | Build + `changeset publish` (CI only)                               |

That is every script in the root `package.json`. Two more helpers are run
directly by CI rather than through a script:

```sh
node scripts/docx-validity.mjs --require-soffice   # `pnpm docx:validity`, but strict
node scripts/pack-smoke.mjs                        # pack the tarball, import it as ESM + CJS
```

Scoping to one package:

```sh
pnpm --filter downword test
pnpm --filter downword-cli build
```

## 3. Repository layout

```
packages/core/            # the `downword` package
  src/
    index.ts              # public API barrel
    convert.ts            # convert() — the conversion entry point
    types.ts              # public types (ConvertOptions, DownwordPlugin, …)
    plugins/math.ts       # `downword/plugins/math`    (stub)
    plugins/mermaid.ts    # `downword/plugins/mermaid` -> src/mermaid/**
  tests/
    convert.test.ts       # unit tests
    golden.test.ts        # golden/snapshot tests over word/document.xml
    helpers/normalize.ts  # strips nondeterminism before snapshotting
    helpers/docx.ts       # unzip helpers + fixture writer
    __fixtures__/out/     # generated .docx (gitignored, read by the CI gate)
    __fixtures__/corrupt/ # negative control for the CI gate (committed)
  tsup.config.ts          # build: entries, dual format, dts
  vitest.config.ts
  .size-limit.json        # bundle budget

packages/cli/             # the `downword-cli` package (bin: downword)
  src/index.ts            # node:util parseArgs, no CLI framework dependency
  tests/cli.test.ts       # spawns the built binary

examples/basic-node/      # ESM consumer  (`import`)
examples/basic-cjs/       # CJS consumer  (`require`)

scripts/docx-validity.mjs # LibreOffice validity gate (T0.4)
scripts/pack-smoke.mjs    # packs the tarball and consumes it as ESM + CJS
.github/workflows/ci.yml  # lint → typecheck → test → build → examples → size,
                          #   plus the pack smoke and docx gate jobs
```

## 4. Adding a renderer node type

The markdown renderer is not implemented yet (this is the T0 scaffold). Today
all conversion logic lives in **`packages/core/src/convert.ts`**, which emits one
plain paragraph per blank-line-separated block. Until the renderer lands, "add a
node type" means: add its handling in `convert.ts` and lock the OOXML output in
with a golden test.

The workflow is the same either way:

1. **Write the golden test first.** Add a case to
   `packages/core/tests/golden.test.ts` with a small markdown sample that
   exercises only your node type. Use `writeFixture()` from
   `packages/core/tests/helpers/docx.ts` so the file lands in
   `tests/__fixtures__/out/` — CI's LibreOffice gate converts everything in that
   directory to PDF, so your new output is checked for real-world openability
   automatically.

2. **Implement it** in `packages/core/src/convert.ts`.

3. **Run the tests and review the snapshot diff.**

   ```sh
   pnpm --filter downword test
   ```

   The first run writes the snapshot into
   `packages/core/tests/__snapshots__/golden.test.ts.snap`. **Read the generated
   OOXML before committing it** — a golden test only helps if the golden output
   was correct on the day it was recorded. Update with:

   ```sh
   pnpm --filter downword test -- -u
   ```

4. **Keep the output deterministic.** Snapshots must not churn between runs. If
   `docx` emits something nondeterministic (ids, timestamps, hashes), add a
   scrubber to `SCRUBBERS` in `packages/core/tests/helpers/normalize.ts` with a
   `what:` description — and only for values that are genuinely nondeterministic.
   Verify with two consecutive runs:

   ```sh
   pnpm test && pnpm test    # second run must report zero written snapshots
   ```

5. **Check the size budget** if you added a dependency:

   ```sh
   pnpm build && pnpm size
   ```

   The budget for `downword` is 150 kB min+gzip **excluding** the `docx` writer
   (which is a ~103 kB gzip pre-bundled blob and is listed under `ignore` in
   `packages/core/.size-limit.json`).

6. **Add a changeset** (see below).

### Adding a plugin entry point

Plugins are separate subpath exports so that heavy optional dependencies stay
out of the main bundle. To add `downword/plugins/<name>`:

1. Create `packages/core/src/plugins/<name>.ts` exporting a factory that returns
   a `DownwordPlugin` (see `packages/core/src/types.ts`).
2. Add the entry to `entry` in `packages/core/tsup.config.ts`.
3. Add the `"./plugins/<name>"` block to `exports` in
   `packages/core/package.json` (both `import` and `require` conditions, each
   with its own `types`).
4. Add a test that imports the built subpath so a broken `exports` map fails CI.

## 5. Useful facts about the `docx` writer

Hard-won constraints — please don't rediscover these the hard way:

- `docx@9.7.1` is a **pre-bundled ~1 MB / ~103 kB gzip** artifact with Node
  polyfills baked in. It is browser-safe (no `fs`/`path`/`stream` imports) but
  effectively not tree-shakeable.
- `ImportedXmlComponent.fromXmlString()` **cannot be inserted directly** — it
  produces an `<undefined>` root element and corrupts the package. Unwrap
  `.root[0]`.
- `ImageRun` string `data` is interpreted as **base64 only**; raw SVG markup
  throws from `atob`. Pass `new TextEncoder().encode(svg)`.
- Ordered lists with `start=N` need a **separate numbering `reference`** per
  distinct start value; `instance` only restarts at the config's level-0 `start`.
- The built-in `bullet: { level }` shares one global numbering instance and is
  unconfigurable — define your own `md-ul` reference instead.
- `Quote` is not a `styles.default` key; declare it in `paragraphStyles`.
- Prefer `Packer.toArrayBuffer()` / `Packer.toBlob()`. `toBuffer()`/`toStream()`
  work in the browser only via bundled polyfills.

## 6. The `.docx` validity gate

Well-formed XML is not the same thing as "Word opens it". Every generated
fixture is converted to PDF by headless LibreOffice in CI:

```sh
pnpm test            # writes fixtures to packages/core/tests/__fixtures__/out/
pnpm docx:validity   # runs scripts/docx-validity.mjs
```

Locally, without LibreOffice installed, the script still runs its structural
checks and then reports the conversion step as skipped:

```
[ skip ] libreoffice - soffice not found. Install LibreOffice, or set SOFFICE=…
```

In CI (`$CI` is set, and `--require-soffice` is passed) a missing `soffice` is a
hard failure. The gate also runs a **negative control** against
`packages/core/tests/__fixtures__/corrupt/corrupt.docx`; if that file ever
converts successfully, the gate reports itself as broken and CI fails.

The negative control forces LibreOffice's `MS Word 2007 XML` import filter,
because otherwise LibreOffice sniffs the content, falls back to its plain text
filter and "converts" the fixture anyway. A **positive control** runs that same
forced filter against a known-good `.docx` first: if the filter is ever renamed
or missing, the negative control would "refuse" everything for the wrong reason,
so the gate fails loudly instead of quietly proving nothing.

Every conversion writes to its own temporary directory. LibreOffice names its
output after the input's basename and sometimes exits `0` while producing
nothing, so sharing one directory would let a leftover PDF from a same-named
file turn a failure into a pass.

## 7. Commit, changesets and releases

- Commit messages: [Conventional Commits](https://www.conventionalcommits.org/)
  (`feat:`, `fix:`, `docs:`, `chore:`, `refactor:`, `test:`, `ci:`).
- Every user-visible change needs a changeset:

  ```sh
  pnpm changeset
  ```

  Pick the affected packages and a bump type, and write the changelog entry in
  the voice of someone reading release notes.

- Releases are automated: merging the "Version Packages" PR opened by the
  Changesets bot triggers `.github/workflows/release.yml`, which publishes to
  npm with **provenance** (`id-token: write` + `npm publish --provenance`).
  Provenance is deliberately _not_ set in `.npmrc`, because that would break
  local `pnpm pack` outside a CI OIDC context.

- **Nothing publishes while the packages are on version `0.0.0`.**
  `changesets/action` runs its publish script on any push to `main` with no
  pending changesets, and `changeset publish` has no special case for `0.0.0` —
  it only skips `private` packages and versions already on the registry. The
  `Decide whether publishing is allowed` step in `release.yml` therefore blanks
  out the publish script until the first real version bump. Remove nothing: the
  guard disarms itself as soon as `packages/core/package.json#version` moves off
  `0.0.0`.

## 8. Pull requests

Before opening a PR:

```sh
pnpm lint && pnpm typecheck && pnpm test && pnpm build && pnpm size
```

Then fill in the [pull request template](.github/PULL_REQUEST_TEMPLATE.md). PRs
that change conversion output must include the updated golden snapshot and say
why the new output is correct.
