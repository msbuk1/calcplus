import { defineConfig } from 'vitest/config';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig({
  esbuild: {
    jsx: 'automatic',
  },
  // PDFKit is imported on demand by File → Export PDF. Pre-bundling it stops
  // the dev server from discovering it mid-click and reloading the page, which
  // otherwise loses the export (the dependency-reload flake).
  optimizeDeps: {
    include: ['pdfkit', 'hyperformula'],
  },
  resolve: {
    alias: {
      '@model': fileURLToPath(new URL('./src/model', import.meta.url)),
      '@ooxml': fileURLToPath(new URL('./src/ooxml', import.meta.url)),
      '@commands': fileURLToPath(new URL('./src/commands', import.meta.url)),
      '@engine': fileURLToPath(new URL('./src/engine', import.meta.url)),
      '@view': fileURLToPath(new URL('./src/view', import.meta.url)),
      '@grid': fileURLToPath(new URL('./src/grid', import.meta.url)),
      '@chrome': fileURLToPath(new URL('./src/chrome', import.meta.url)),
      '@io': fileURLToPath(new URL('./src/io', import.meta.url)),
      '@testing': fileURLToPath(new URL('./src/testing', import.meta.url)),
    },
  },
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    testTimeout: 30_000,
    hookTimeout: 30_000,
    coverage: {
      provider: 'v8',
      include: ['src/model/**', 'src/ooxml/**', 'src/commands/**', 'src/view/**'],
    },
  },
});
