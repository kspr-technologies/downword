<!--
Thanks for contributing to downword.
Keep the PR focused; unrelated changes are much harder to review.
-->

## What does this change?

<!-- One or two sentences. What is different after this PR? -->

## Why?

<!-- Link the issue: "Closes #123". If there is no issue, explain the motivation. -->

Closes #

## How was it verified?

<!-- Delete what does not apply. -->

- [ ] `pnpm lint`
- [ ] `pnpm typecheck`
- [ ] `pnpm test`
- [ ] `pnpm build`
- [ ] `pnpm size` (needed if dependencies or bundled code changed)
- [ ] Opened the generated `.docx` in a real Word version — state which:

## Conversion output

<!--
If this PR changes what ends up in the .docx, answer these. Otherwise write N/A.
-->

- Golden snapshots updated: <!-- yes / no / N/A -->
- Why the new OOXML is correct:
- New fixture written to `packages/core/tests/__fixtures__/out/`: <!-- yes / no -->

## Changeset

- [ ] I ran `pnpm changeset` (required for anything user-visible)
- [ ] Not needed — this change is internal only (docs, CI, tests, refactor)

## Breaking changes

<!-- Describe the break and the migration path, or write "None". -->

None

## Checklist

- [ ] I read [CONTRIBUTING.md](../CONTRIBUTING.md)
- [ ] Commits follow Conventional Commits
- [ ] New behaviour has tests
- [ ] Public API changes are reflected in the types and docs
- [ ] No network requests were added to the library (downword runs offline, by design)
