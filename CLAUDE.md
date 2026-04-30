# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

- `npm test` / `npm run test:watch` — Vitest under jsdom (`vitest.config.js`). No test files currently exist in the repo; new specs should live alongside `src/` as `*.test.js` (or under `tests/`) and rely on the jsdom environment.
- `npm run build` — Rollup builds two artifacts from `src/index.js` (`rollup.config.js`):
  - `dist/dom-to-pptx.mjs` + `.cjs` — library builds. All runtime deps (`pptxgenjs`, `html2canvas`, `jszip`, `fonteditor-core`, `opentype.js`, `pako`) are marked **external** and must be installed by the consumer.
  - `dist/dom-to-pptx.bundle.js` — UMD `domToPptx` global for `<script>` use. Bundles **everything** plus Node polyfills (`rollup-plugin-polyfill-node`, plus `buffer`/`stream-browserify`/`process`/`util`/`events` shims) so the file can run standalone in a browser.
- `npm run lint` / `npm run format` — ESLint flat config + Prettier. ESLint ignores `dist/**`, downgrades `no-unused-vars` and `no-undef` to warnings, and assumes browser globals.
- Package manager: lockfile is `pnpm-lock.yaml` but `CONTRIBUTING.md` documents `npm install` / `npm test`. Either works; don't regenerate the other lockfile.

## Architecture

The library is a **DOM measurement + style translation engine**, not a screenshot tool. Everything is driven by `getBoundingClientRect()` and `getComputedStyle()` from a live browser; output is native, editable PPTX shapes/text/images.

### Pipeline (`src/index.js`)

`exportToPptx(target, options)` is the only public entry point. Per call:

1. **Layout setup.** Picks slide dimensions from `options.width/height` (custom layout) → `options.layout` (`LAYOUT_4x3`/`16x10`/`WIDE`) → default `LAYOUT_16x9` (10 × 5.625 in). Stashes the chosen size on `extendedOptions._slideWidth/_slideHeight` for downstream code.
2. **Per-slide processing** (`processSlide`):
   - Computes a `layoutConfig` with `rootX/rootY/scale/offX/offY`. Scale is `min(slideW/contentW, slideH/contentH)` so the source element fits the slide; offsets center it. **Children are positioned absolutely against the root**, so Flexbox/Grid never need to be "understood" — only their final laid-out coordinates matter.
   - **Two-phase traversal**: a synchronous `collect()` walk pushes lightweight render items onto `renderQueue`; any heavy work (image loading, html2canvas snapshots, SVG-to-PNG) is deferred to a closure pushed onto `asyncTasks`. After the walk, all async tasks run in parallel via `Promise.all`. This keeps the hot DOM-reading phase tight and avoids reflow churn.
   - The render queue is filtered (drop `skip`/empty-image items), then sorted by `(zIndex, domOrder)` to preserve stacking. Final items are dispatched to `slide.addShape` / `addImage` / `addText` / `addTable`.
3. **Font embedding** (only if `autoEmbedFonts` or `options.fonts` are non-empty). The PPTX is generated once via `pptx.write({ outputType: 'blob' })`, then `PPTXEmbedFonts` (`src/font-embedder.js`) re-opens the zip with `JSZip` and:
   - Adds a `Default Extension="fntdata"` entry to `[Content_Types].xml`.
   - Sets `saveSubsetFonts` / `embedTrueTypeFonts` on `p:presentation` and inserts `p:embeddedFont` entries.
   - Adds `Relationship` entries to `ppt/_rels/presentation.xml.rels` pointing at `ppt/fonts/<rid>.fntdata`.
   - Converts each font buffer to EOT-style `fntdata` via `fontToEot` in `src/font-utils.js` (uses `fonteditor-core` + `pako`) — PowerPoint's font embedding format is EOT, not raw TTF/WOFF.
   - Auto-detection (`getUsedFontFamilies` + `getAutoDetectedFonts` in `utils.js`) scans the DOM for used font families and tries to find their `url(...)` in stylesheets — so CORS-friendly fonts (Google Fonts with `crossorigin="anonymous"`) embed automatically.
4. **Output.** Always returns the final `Blob`; downloads via a transient `<a>` tag unless `options.skipDownload` is true.

### Module roles

- `src/index.js` — entry point, slide pipeline, render-queue dispatcher. ~1200 lines including `prepareRenderItem` which decides whether a node becomes a shape, text, image, table, or html2canvas raster fallback.
- `src/utils.js` — the bulk of the style-translation logic. Lives close to the metal: `parseColor` (uses a hidden canvas for color normalization), `getTextStyle`, `getVisibleShadow` (CSS Cartesian → PPTX polar shadows), `generateGradientSVG` (CSS `linear-gradient` parser → SVG vector for gradient fills), `generateBlurredSVG`, `generateCompositeBorderSVG`, `generateCustomShapeSVG`, writing-mode helpers, `extractTableData`, `collectTextParts` (mixed-style rich text), font-detection helpers.
- `src/image-processor.js` — `getProcessedImage` draws the source image to an offscreen canvas at 2× resolution, builds a rounded-rect path with per-corner radius clamping, applies `globalCompositeOperation = 'source-in'` to mask without halos, and respects `object-fit` (`fill`/`contain`/`cover`/`none`/`scale-down`) and `object-position`. Returns a PNG data URL. Requires CORS-accessible images.
- `src/font-embedder.js` / `src/font-utils.js` — post-process the generated PPTX zip to embed fonts.

### Key invariants when editing

- **All coordinates are absolute against the slide root.** Don't try to preserve CSS layout intent; preserve the rendered rectangle. New element types should call `getBoundingClientRect()` and convert to inches via the `scale` and `offX/offY` from `layoutConfig`.
- **Unit conversions.** `PPI = 96` and `PX_TO_INCH = 1/96` in `index.js`. Font sizes use `px → pt` ≈ `× 0.75`. Don't round font sizes destructively — fractional points (e.g. 11.3pt) are preserved deliberately (per v1.1.7 changelog).
- **Z-ordering** is `(zIndex, domOrder)`, not raw DOM order — `collect()` propagates the parent z-index unless the node has its own non-`auto` value.
- **html2canvas is a fallback**, used for things the engine can't translate to native shapes (e.g. `backdrop-filter`). Prefer extending the native translators in `utils.js` over rasterizing.
- **Rollup `external` list** in `rollup.config.js` must be kept in sync with `package.json` `dependencies`. Adding a new runtime dep means adding it to `external` for the library build (or it gets duplicated into consumer bundles), while leaving the bundle build to swallow it.
- **Browser-only APIs.** The library uses `document`, `window.getComputedStyle`, `Image`, `Canvas`, `URL.createObjectURL`. Tests run under `jsdom` (`vitest.config.js`); anything that needs real canvas/image decoding can't be unit-tested there without mocks.
