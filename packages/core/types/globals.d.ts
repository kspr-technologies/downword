/**
 * Build-time constant injected by tsup (`define`) and by vitest (`define`).
 * Its value is always `packages/core/package.json#version`.
 *
 * Keep the two definitions in sync:
 *  - packages/core/tsup.config.ts
 *  - packages/core/vitest.config.ts
 */
declare const __DOWNWORD_VERSION__: string;
