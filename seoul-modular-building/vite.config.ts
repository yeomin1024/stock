// VERSION: v1.0.0 — 2026-11-02 — Initial implementation
import { defineConfig } from 'vite';

export default defineConfig({
  // Relative base so the built dist/ works from any static path (or a sub-folder).
  base: './',
  build: {
    target: 'es2022',
    // three.js core alone is ~500 kB minified; this is expected for a single-page 3D viewer.
    chunkSizeWarningLimit: 800,
  },
});
