import { describe, it, beforeAll, afterAll, expect } from 'vitest';
import { readdir, mkdir, readFile, writeFile, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { runCase, closeBrowser, SLIDE_W, SLIDE_H } from './lib/browser.js';
import { pptxBufferToPng } from './lib/rasterize.js';
import { diffPngs } from './lib/diff.js';
import { writeReport } from './lib/report.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CASES_DIR = path.join(__dirname, 'cases');
const OUTPUT_DIR = path.join(__dirname, 'output');
const REPORT_DIR = path.join(__dirname, 'report');

// Per-case threshold for the foreground-aware *content* delta — block-level
// edge-density + color-shift mismatch (see diff.js). Raw pixel-percent is
// still computed and reported, but the budget gates on contentPercent so
// background-dominated slides can't mask foreground regressions. Default is
// loose enough that current state passes; tighten as fixes land.
const PERCENT_BUDGET = Number(process.env.FIDELITY_BUDGET ?? 85);

const results = [];

// Each case is { name, file, slideIndex? }.
//   - Top-level *.html files become a single case.
//   - A subdirectory containing manifest.json becomes N cases, one per slide.
//     The harness loads the page once with `?print-pdf` (reveal.js print
//     mode), reuses the context across slides, and exports the nth
//     `.pdf-page` element for slideIndex N.
//   - Subdirectories without a manifest fall through to per-file cases.
async function listCases() {
  const entries = await readdir(CASES_DIR, { withFileTypes: true });
  const cases = [];
  for (const entry of entries) {
    if (entry.isFile() && entry.name.endsWith('.html')) {
      cases.push({ name: entry.name.replace(/\.html$/, ''), file: entry.name });
      continue;
    }
    if (!entry.isDirectory()) continue;

    const manifestPath = path.join(CASES_DIR, entry.name, 'manifest.json');
    let manifest = null;
    try {
      await stat(manifestPath);
      manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
    } catch {
      // no manifest — fall through
    }

    if (manifest && manifest.entry && Number.isInteger(manifest.slides)) {
      const file = `${entry.name}/${manifest.entry}`;
      const pad = String(manifest.slides).length;
      const budget = typeof manifest.budget === 'number' ? manifest.budget : undefined;
      for (let i = 0; i < manifest.slides; i++) {
        const idx = String(i + 1).padStart(pad, '0');
        cases.push({ name: `${entry.name}-${idx}`, file, slideIndex: i, budget });
      }
    } else {
      const sub = await readdir(path.join(CASES_DIR, entry.name));
      for (const f of sub) {
        if (f.endsWith('.html')) {
          cases.push({
            name: `${entry.name}-${f.replace(/\.html$/, '')}`,
            file: `${entry.name}/${f}`,
          });
        }
      }
    }
  }
  return cases.sort((a, b) => a.name.localeCompare(b.name));
}

beforeAll(async () => {
  await rm(OUTPUT_DIR, { recursive: true, force: true });
  await mkdir(OUTPUT_DIR, { recursive: true });
  await mkdir(REPORT_DIR, { recursive: true });
});

afterAll(async () => {
  await closeBrowser();
  // Copy result PNGs into the report folder so relative <img> srcs work.
  const reportRows = results.map((r) => {
    const base = r.name;
    return {
      name: base,
      sourcePng: r.sourcePngPath ? path.relative(REPORT_DIR, r.sourcePngPath) : '',
      pptxPng: r.pptxPngPath ? path.relative(REPORT_DIR, r.pptxPngPath) : '',
      diffPng: r.diffPngPath ? path.relative(REPORT_DIR, r.diffPngPath) : '',
      percent: r.percent ?? 0,
      contentPercent: r.contentPercent ?? 0,
      edgePercent: r.edgePercent ?? 0,
      colorPercent: r.colorPercent ?? 0,
      mismatched: r.mismatched ?? 0,
      total: r.total ?? 0,
      budget: r.budget ?? 0,
      error: r.error,
    };
  });
  await writeReport(path.join(REPORT_DIR, 'index.html'), reportRows);
});

describe('fidelity harness', async () => {
  const caseEntries = await listCases();

  for (const entry of caseEntries) {
    const { name, file, slideIndex, budget } = entry;
    const caseBudget = budget ?? PERCENT_BUDGET;
    it(name, async () => {
      const result = { name };
      results.push(result);

      try {
        const { sourcePng, pptxBuffer } = await runCase(
          path.join(CASES_DIR, file),
          { slideIndex },
        );

        const pptxPath = path.join(OUTPUT_DIR, `${name}.pptx`);
        await writeFile(pptxPath, pptxBuffer);

        const sourcePngPath = path.join(OUTPUT_DIR, `${name}.source.png`);
        await writeFile(sourcePngPath, sourcePng);
        result.sourcePngPath = sourcePngPath;

        const pptxPng = await pptxBufferToPng(pptxBuffer, OUTPUT_DIR, name);
        const pptxPngPath = path.join(OUTPUT_DIR, `${name}.pptx.png`);
        await writeFile(pptxPngPath, pptxPng);
        result.pptxPngPath = pptxPngPath;

        const {
          diffPng,
          mismatched,
          total,
          percent,
          contentPercent,
          edgePercent,
          colorPercent,
        } = await diffPngs(sourcePng, pptxPng, SLIDE_W, SLIDE_H);
        const diffPngPath = path.join(OUTPUT_DIR, `${name}.diff.png`);
        await writeFile(diffPngPath, diffPng);
        result.diffPngPath = diffPngPath;
        result.mismatched = mismatched;
        result.total = total;
        result.percent = percent;
        result.contentPercent = contentPercent;
        result.edgePercent = edgePercent;
        result.colorPercent = colorPercent;
        result.budget = caseBudget;

        expect(contentPercent).toBeLessThanOrEqual(caseBudget);
      } catch (err) {
        result.error = err && err.stack ? err.stack : String(err);
        throw err;
      }
    }, 60_000);
  }
});
