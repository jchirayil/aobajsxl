// vitest.config.ts
/// <reference types="vitest" />
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // Your test files location
    include: ['test/**/excel-base.test.{ts,js}'],
    // Enable globals (like describe, it, expect) without explicit imports
    globals: true,
    // Environment for testing (e.g., 'node' for backend, 'jsdom' for DOM manipulation)
    environment: 'node', // Or 'jsdom' if your library has DOM dependencies
    // Options for type checking
    typecheck: {
      tsconfig: './tsconfig.test.json', // Point to your test tsconfig
    },
    testTimeout: 0
    // For ESM module resolution and path aliases
    // If you have "paths" in tsconfig.json, Vitest can often pick them up.
    // server: {
    //   deps: {
    //     inline: [
    //       // If you have specific packages that need to be inlined for ESM
    //       // e.g., if a dependency is published as CJS but imports ESM internally
    //     ],
    //   },
    // },
  },
});