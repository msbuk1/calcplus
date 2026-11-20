import { fileURLToPath, URL } from 'node:url';
import { defineConfig } from 'vitest/config';
import baseConfig from '../../vite.config';

/**
 * The Node microbenchmark suite (performance §6, layer 1). It is not part of
 * `npm test`: the corpus includes a million-cell shape and a pathological
 * fan-out, so it runs on demand with `npm run bench:model`.
 *
 * The root config's aliases are reused unchanged; only the test selection is
 * replaced. A single fork keeps the heap measurement meaningful and
 * `--expose-gc` lets the harness force garbage collection before it records the
 * steady heap.
 */
export default defineConfig({
  root: fileURLToPath(new URL('../..', import.meta.url)),
  esbuild: baseConfig.esbuild,
  resolve: baseConfig.resolve,
  test: {
    include: ['bench/model/**/*.bench.ts'],
    environment: 'node',
    testTimeout: 600_000,
    hookTimeout: 600_000,
    fileParallelism: false,
    pool: 'forks',
    execArgv: ['--expose-gc'],
  },
});
