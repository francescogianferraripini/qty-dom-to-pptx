import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    name: 'fidelity',
    environment: 'node',
    globals: false,
    include: ['tests/fidelity/**/*.test.js'],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // Cases share a single Playwright browser instance held in module state;
    // running them in parallel would multiply browser launches.
    pool: 'forks',
    fileParallelism: false,
  },
});
