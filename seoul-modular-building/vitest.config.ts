// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: unit tests for the pure layout/validation layer
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'node',
  },
});
