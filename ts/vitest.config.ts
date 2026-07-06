import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    // The differential tests initialize the C# WASM engine once.
    testTimeout: 30_000,
    hookTimeout: 60_000,
  },
});
