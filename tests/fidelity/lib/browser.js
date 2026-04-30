import { chromium } from 'playwright';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const BUNDLE_PATH = path.resolve(__dirname, '../../../dist/dom-to-pptx.bundle.js');

// Slide pixel dims for the default 16x9 layout (10in x 5.625in @ 96 DPI).
export const SLIDE_W = 960;
export const SLIDE_H = 540;

let _browser = null;
let _bundleSource = null;
// Multi-slide reveal cases share one browser context+page per case file:
// the page is loaded once with `?print-pdf`, every `.pdf-page` exported in
// turn. Single-slide cases skip the cache and tear down per call.
const _pageCache = new Map();

export async function getBrowser() {
  if (!_browser) {
    _browser = await chromium.launch({ headless: true });
  }
  return _browser;
}

export async function closeBrowser() {
  for (const { context } of _pageCache.values()) {
    try {
      await context.close();
    } catch {
      // best-effort teardown
    }
  }
  _pageCache.clear();
  if (_browser) {
    await _browser.close();
    _browser = null;
  }
}

async function loadBundle() {
  if (!_bundleSource) {
    _bundleSource = await readFile(BUNDLE_PATH, 'utf8');
  }
  return _bundleSource;
}

async function setupPage(caseHtmlPath, { printMode }) {
  const browser = await getBrowser();
  const context = await browser.newContext({
    viewport: { width: SLIDE_W, height: SLIDE_H },
    deviceScaleFactor: 1,
  });
  const page = await context.newPage();
  const fileUrl =
    'file://' + path.resolve(caseHtmlPath) + (printMode ? '?print-pdf' : '');
  await page.goto(fileUrl, { waitUntil: 'load' });
  await page.evaluate(async () => {
    if (window.__revealReady && typeof window.__revealReady.then === 'function') {
      await window.__revealReady;
    }
  });
  if (printMode) {
    await page.waitForFunction(
      () => document.querySelectorAll('.pdf-page').length > 0,
      null,
      { timeout: 30_000 },
    );
    // Yield two animation frames so reveal's print layout (transforms, sizing)
    // settles before we read computed styles or rects.
    await page.evaluate(
      () =>
        new Promise((r) =>
          requestAnimationFrame(() => requestAnimationFrame(r)),
        ),
    );
  }
  const bundle = await loadBundle();
  await page.addScriptTag({ content: bundle });
  await page.evaluate(() => document.fonts && document.fonts.ready);
  return { context, page };
}

/**
 * Run a fidelity case in a (possibly cached) browser context.
 * Returns { sourcePng: Buffer, pptxBuffer: Buffer }.
 *
 * - Single-slide cases (slideIndex == null) export `#target` from a fresh
 *   context that's torn down on return.
 * - Multi-slide cases (slideIndex != null) load the page once with
 *   `?print-pdf` and export the nth `.pdf-page`. The page+context are kept
 *   in `_pageCache` so subsequent slides reuse them; closeBrowser() tears
 *   them down at the end of the run.
 */
export async function runCase(caseHtmlPath, options = {}) {
  const { slideIndex = null } = options;
  const printMode = slideIndex != null;

  let entry;
  let owned = false;
  if (printMode) {
    const key = path.resolve(caseHtmlPath);
    entry = _pageCache.get(key);
    if (!entry) {
      entry = await setupPage(caseHtmlPath, { printMode: true });
      _pageCache.set(key, entry);
    }
  } else {
    entry = await setupPage(caseHtmlPath, { printMode: false });
    owned = true;
  }

  const { context, page } = entry;
  try {
    let target;
    if (printMode) {
      const handles = await page.$$('.pdf-page');
      if (slideIndex >= handles.length) {
        throw new Error(
          `pdf-page index ${slideIndex} out of range (have ${handles.length})`,
        );
      }
      target = handles[slideIndex];
    } else {
      target = await page.$('#target');
      if (!target) throw new Error(`Case missing #target: ${caseHtmlPath}`);
    }

    const sourcePng = await target.screenshot({ type: 'png' });

    const base64 = await page.evaluate(
      async ({ idx }) => {
        const el =
          idx == null
            ? document.getElementById('target')
            : document.querySelectorAll('.pdf-page')[idx];
        const blob = await window.domToPptx.exportToPptx(el, {
          skipDownload: true,
          layout: 'LAYOUT_16x9',
        });
        const buf = await blob.arrayBuffer();
        let binary = '';
        const bytes = new Uint8Array(buf);
        const chunk = 0x8000;
        for (let i = 0; i < bytes.length; i += chunk) {
          binary += String.fromCharCode.apply(
            null,
            bytes.subarray(i, i + chunk),
          );
        }
        return btoa(binary);
      },
      { idx: printMode ? slideIndex : null },
    );

    const pptxBuffer = Buffer.from(base64, 'base64');
    return { sourcePng, pptxBuffer };
  } finally {
    if (owned) {
      await context.close();
    }
  }
}
