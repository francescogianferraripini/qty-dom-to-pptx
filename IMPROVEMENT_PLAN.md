# Fidelity Improvement Plan

The plan is sequenced so each phase unblocks the next. Phase 0 is non-negotiable — every later claim of "this looks better" needs a way to be measured, otherwise we're just shipping vibes.

---

## Phase 0 — Fidelity test harness (prerequisite) ✅ landed

**Goal:** make "is the output more faithful?" a question with an answer, not an opinion.

**What shipped**

Folder layout, separated from `src/`:

```
tests/
├── unit/
│   ├── svg-generators.test.js       # snapshot tests for SVG generators (jsdom)
│   └── __snapshots__/
└── fidelity/
    ├── vitest.config.js             # node env, single fork (browser is shared)
    ├── fidelity.test.js             # iterates cases/*.html
    ├── cases/                       # HTML fixtures — one per CSS feature, plus
    │   │                            # subfolder cases (with manifest.json) for real-world decks
    │   └── quantyca/                # reveal.js deck, 32 slides, navigated via Reveal.slide()
    ├── lib/
    │   ├── browser.js               # Playwright Chromium + bundle injection
    │   ├── rasterize.js             # libreoffice --headless → PNG (slide 1)
    │   ├── diff.js                  # sharp resize + pixelmatch
    │   └── report.js                # HTML side-by-side report
    ├── output/                      # gitignored: .pptx, .source.png, .pptx.png, .diff.png per case
    └── report/                      # gitignored: index.html
```

Wiring: root `vitest.config.js` scoped to `tests/unit`; new `npm run test:fidelity` script; devDeps added (`playwright`, `pixelmatch`, `pngjs`, `sharp`); `.gitignore` updated to keep `tests/` tracked while ignoring scratch root files and harness artifacts.

The fidelity runner loads each case via `page.goto('file://...')` (so relative URLs to fonts, CSS, and images resolve), injects `dist/dom-to-pptx.bundle.js`, calls `domToPptx.exportToPptx(#target, { skipDownload: true, layout: 'LAYOUT_16x9' })`, screenshots `#target` (960×540, matching the slide pixel dims at 96 DPI), shells out to LibreOffice for PPTX→PNG, resizes both to 960×540 with `sharp`, and diffs with `pixelmatch`. Per-case threshold is the env var `FIDELITY_BUDGET` (default 60% — deliberately loose initially; tighten as phases land fixes); a per-case override can be set in a folder's `manifest.json`.

Unit-side snapshot tests cover `generateGradientSVG`, `generateCompositeBorderSVG`, and `generateCustomShapeSVG`; they decode the data-URL output and normalize random `clip_*` ids so snapshots stay deterministic.

**Multi-slide / real-world decks.** A case folder may contain `manifest.json` with `{ entry, slides, budget? }`. `listCases` expands it into N cases (`<folder>-01` … `<folder>-NN`); `runCase` accepts a `slideIndex`, calls `Reveal.slide(idx, 0)` after the page exposes a `window.__revealReady` promise, and yields two animation frames before measuring so the new section's computed styles have settled. This lets a real reveal.js deck (or any JS-driven multi-slide UI) drop into `cases/` and be exercised slide-by-slide without hand-authoring 32 wrappers.

**Corpus today (40 cases, all passing their per-case budget)**
- 8 synthetic micro-cases at the default 60% budget — text-basic, gradient-linear, box-shadow, border-radius, transform-rotate, table-basic, flex-layout, solid-color. Baseline pixel deltas establish where the real fidelity gaps are (e.g. ~48% on the linear gradient — direction translation is visibly off — and ~30% on flex-layout).
- 32 slides from the Quantyca reveal.js deck (`cases/quantyca/`) at an 80% budget. 21/32 land under 5% delta; the worst is `quantyca-27` (CSS-grid logo wall) at 71.6% — a real defect the harness now flags.

**Caveat the corpus surfaced.** Several deck slides report ~4% delta even though the foreground (text, decorations) clearly didn't translate, because a full-bleed background dominates the pixel diff. Raw pixel-percent under-counts content fidelity on these slides. A perceptual or content-aware metric (e.g. weight differences inside text/edge regions, or a separate text-OCR diff) would make the budget meaningful again — file under "harness improvements" before tightening Phase 1+ budgets.

**Post-mortem from the Quantyca canary (40-case run).** A side-by-side review of every quantyca-* slide surfaced four defect classes that drive ~all of the residual delta. They informed the Phase 1 reshuffle below:

1. **Ancestor CSS transform is invisible to the pipeline.** Reveal.js applies `transform: scale(N)` (here `N ≈ 0.771`) on `.slides` to fit its 960×700 logical slide into the 960×540 viewport. `getBoundingClientRect()` returns post-transform sizes (correct), but `getComputedStyle().fontSize` (and paddings, line-height, letter-spacing, border widths) returns the *logical* unscaled value. `getTextStyle` (`src/utils.js:476`) writes the logical font size into PPTX, so every text run is shipped ~30% too large for the box that contains it — the headline title overflows both edges of the slide on every content slide. This single bug accounts for the 71%-diff catastrophe on `quantyca-27` and the 12-25% diffs on slides 14, 16, 18, 22, 25, 28.
2. **Decorative `::before`/`::after` are dropped.** The Quantyca design system rides on pseudo-elements: every H2 has a 56×3px burgundy underline, every content slide has a "QUANTYCA · DATA, AT CORE" corner mark, the iceberg slide has a "LINEA D'ACQUA" badge, the 2×2 matrix has writing-mode axis labels, and the timeline has its dots — none translate today. The pipeline only walks real DOM nodes; pseudo-element `content` is partially handled (`utils.js:1024-1038`) but their layout boxes (background, border, transform) are not.
3. **`text-decoration: line-through` is dropped.** `quantyca-18`'s "before" column shows five strikethrough items in source, plain text in PPTX. `getTextStyle` reads `underline` only (`src/utils.js:479`).
4. **Icon-font glyphs render as `?`/`8` boxes.** Font Awesome `<i class="fa-…">` runs ship with the right Unicode codepoint but no embedded glyph, because `getAutoDetectedFonts` either fails CORS on the CDN or the @font-face rule isn't traversed for icon fonts. Visible on `quantyca-03`, `quantyca-22`, `quantyca-25` and many more.

Items 1–3 are one-file fixes with outsized leverage; item 4 needs a small embedding-fallback path. They reorder the Phase 1 list and pull two items out of Phase 5 / Phase 8.

**Architectural addendum — switch the canary to reveal print mode.** A post-Phase-1 visual review of all 32 quantyca-* renders showed every slide still visibly broken (foreground shifted ~50–150 px left, right ~30–40% of canvas blank, single-letter ghost columns down the left edge, decorative SVGs missing) even though the pixel-percent metric reports 2–8% delta. This is not one engine bug — it's a class of failures that all stem from exporting reveal.js's *live* presentation DOM:

- `.slides` carries a fitting `transform: scale(N)`;
- the active `<section>` is centered via `top:50%; left:50%; translate(-50%,-50%)`;
- inactive sections (`.past`/`.future`) remain in the tree, hidden only via 3D transforms — `collect()` walks them and emits their text as ghost letter columns;
- decorative `::before`/`::after` attach to ancestors above the active section.

Each has its own Phase 1/8 item, but together they make framework-driven decks the engine's worst case and make every later phase's canary measurements unreliable.

**The shortcut.** Reveal already has a static-export DOM mode: loading the deck with `?print-pdf` (or `Reveal.configure({ view: 'print' })`) lays every slide out in normal flow at its configured pixel dimensions inside `.pdf-page` wrappers — no ancestor transforms, no inactive-section bleed, fragments collapsed to their final state. The harness should switch to this mode for the Quantyca case (and any future reveal-based case): open the URL once with `?print-pdf`, iterate `document.querySelectorAll('.pdf-page')`, pass each directly as `target` to `exportToPptx`. This collapses 32 browser context startups into one and removes the `Reveal.slide(idx)` plumbing.

This is a *harness* change, not a library change. The lib stays framework-agnostic; Phase 1.1's ancestor-transform compensation remains relevant for callers that pass a transformed subtree. The canary just stops surfacing problems that don't exist for users running the engine on a clean DOM, and the metric stops being polluted by foreground failures that are really framework-coordination failures.

**Knock-on effects.**
- Whole-slide horizontal shift, inactive-section bleed-through, and partial-background fills disappear from the canary as work items — they are diagnoses of the live-DOM coupling, not the engine.
- Phase 0's metric-replacement TODO becomes less urgent: print-mode renders fail loudly (foreground actually moves), so residual high-delta slides will be diagnosable on inspection. Worth keeping the metric work on the list, but no longer gating.
- Decorative `::before`/`::after` (Phase 1.2) stays — print mode preserves them — and remains a real defect to fix.
- One-time validation: diff the print-mode CSS against the presentation-mode CSS on slide 1 to confirm the user theme has no `@media print` overrides that would make the export disagree with what speakers see live. Quantyca styles are clean; user decks may not be.

**Out of scope.** Generalizing this trick to non-reveal frameworks (Spectacle, Slidev, plain CSS scroll-snap decks) would mean either framework-specific adapters or a generic "neutralize ancestor transforms before measuring" preprocessor in the lib. Punt until a second framework canary is in the corpus.

**Phase 0 follow-ups landed (after Phase 1)**

Two of the "still TODO" items above shipped together:

1. **Quantyca canary switched to reveal print mode** (`tests/fidelity/lib/browser.js`). The harness now appends `?print-pdf` to the case URL on multi-slide cases, waits for `__revealReady` and `.pdf-page` elements to materialize, and exports each `.pdf-page` directly. The 32 quantyca contexts collapsed into one (cached per file path; closed in `closeBrowser`). The `slideIndex` / `Reveal.slide()` plumbing is gone. Effect: every previously catastrophic delta dropped — quantyca-27 (CSS-grid logo wall) went from 71.6% → 6.92% raw, and the ghost-letter-column / right-strip-blank artifacts disappeared from every slide.

2. **Foreground-aware fidelity metric** (`tests/fidelity/lib/diff.js`). Raw pixel-percent is still computed and reported but no longer gates the budget — it under-reported broken slides by 5–10×. The new `contentPercent` combines two block-level signals (16×16 blocks):
   - **Edge delta**: `|srcEdgeSum − dstEdgeSum| / max(...)`, weighted by `max(srcEdgeSum, dstEdgeSum)`. Block sums are max-pooled 3×3 before comparing so identical content shifted by half a block doesn't false-flag. Catches missing/extra structural content (text, lines, icons, borders).
   - **Color delta**: per-block mean-luminance shift, thresholded at 15 lum and saturated at 60. Weighted by contrast against the slide's modal background so flat regions carry no weight. Catches missing fills, broken tinted boxes, gradient direction errors — things the edge metric misses because they have no strong gradients inside.
   - `contentPercent = max(edge%, color%)`; both components are surfaced in the HTML report alongside raw% and budget for diagnostics.

**Effects of the two changes — what the metric now flags as real defects**

- **Spread is honest.** Visually-clean cases (003 box-shadow, 004 border-radius, 005 transform, 008 solid): 0–7% content. Visually broken cases climb proportionally to defect severity. Previously visually-broken slides reporting 5–8% raw now report 50–80% content.
- **Phase 1.1 (ancestor-transform compensation) is effectively *not* working on synthetic case 010.** The PPTX renders only the latter half of every line ("ck headline" / "nps over the lazy dog.") and drops the SHIPPED button entirely — content shifted off the left edge. Content delta 70%, raw% 2.5%. The Quantyca-canary improvement attributed to the cumulative-transform tracker may have come from other Phase 1 fixes; needs re-examination.
- **`opacity: 0` elements are exported as if visible.** Quantyca-08 ships with all five reveal fragments visible in PPTX even though the source `.pdf-page` snapshot has them at `opacity: 0` (fragments before any clicks). Content delta 59%, raw% 3.7%. The collect walk needs to skip nodes with computed `opacity: 0` (and `visibility: hidden`). Affects most reveal decks with progressive disclosure.
- **Decorative `::before`/`::after` pseudo-elements are still the highest-leverage open Phase-1 item** (1.2). On the Quantyca deck this single fix would land hero backgrounds (`.hero::before` with `background: url(page-header.svg)`), H2 underlines, corner brand marks, iceberg badge, timeline dots. The tight cluster of 50–59% deltas on q-04, q-08, q-19–22, q-26, q-28 is dominated by these.

**Score snapshot (current main)**

| Cohort | content delta band | notes |
|---|---|---|
| Visually-clean synthetic | 0–7% | passes any reasonable budget |
| Synthetic text noise (001, 011–015, 006-table) | 6–37% | renderer subpixel diffs + Phase 1 pseudo-element gap |
| Gradient miscolor (002) | 8% | edge sees nothing; color signal alone |
| Real synthetic regressions (007 flex, 010 transform-scale) | 70–81% | both pass current 85% budget but flag as defects to fix |
| Quantyca visibly-OK | 14–25% | text-only slides where the engine renders correctly |
| Quantyca visibly-broken | 50–59% | concentrated on pseudo-element-heavy slides |

**Budgets** (now applied to `contentPercent`, not raw): default 85% (`FIDELITY_BUDGET` env override), Quantyca manifest 65%. All 46 cases pass. ~5–25 pt headroom to the next-tightest case so regressions still trip.

**Still TODO in this phase**

- Render-queue snapshot tests for a fixture DOM (the plan calls for these, but `prepareRenderItem` builds the queue inside `processSlide` and doesn't currently expose it for inspection — needs a small testing seam).
- A CI wrapper that fails if any case's `contentPercent` increases >1pt vs. the previous main.
- ~~Grow the synthetic corpus from 14 → ~30 cases as Phases 2–3 land.~~ ✅ landed (cases 020–030; 30 synthetic + 32 quantyca = 61 total). Six new cases pass clean as regression guards (024 rgba/hex8/hsla, 025 oklch sRGB-clip, 026 transform-translate, 027 leaf transform-scale, 028 isolated transform-origin); five flag known engine gaps under the 85% budget — 020 multi-shadow-stack at 58% (Phase 2.1), 021 inset-shadow at 66% (Phase 2.2), 022 outline-stroke at 81% (Phase 2.3, tightest headroom), 023 transparent-border at 32% (Phase 2.4), 029/030 elliptical & per-corner radius at ~22% (Phase 3.3). Two unrelated engine bugs surfaced and were sidestepped via case redesign rather than fixed: `display: flex` text-containers collapse height (caused initial 025/027 failures), and `line-height`-based vertical text centering doesn't translate. Both worth a Phase 1 follow-up case + fix when convenient.

**Phase 1.1 re-validation + opacity:0 skip ✅ landed**

Two follow-ups from the post-Phase-1.2 score snapshot shipped together:

- **Phase 1.1 was reading sizes from `offsetWidth` while reading positions from `rect.left`.** `offsetWidth` is the logical (pre-transform) width, but `getBoundingClientRect().left` is post-ancestor-transform. For case 010 (`transform: scale(0.6)` on the wrapper, `#target` as the export root's child), this combination placed the H1 at `centerX = rect.left + rect.width/2 = 480px` (post-transform) but with `unrotatedW = offsetWidth/96 = 14.67"` (logical), so `x = (centerX/96) - unrotatedW/2 = 5 - 7.33 = -2.33"` — the ~2.3" left-shift the canary flagged. Fix: keep `offsetWidth/Height` (correct for rotated elements, where `rect.*` returns the AABB of the rotated shape and is too large) but convert with `config.styleScale` rather than `config.scale`. `styleScale = scale × ancestorScale × cumulative` already folds in both above-root and below-root transforms, so a logical-px size becomes the right inch size on screen. For unrotated, untransformed cases the math collapses back to the previous `* scale`. Case 010 dropped from ~70% → 20% content delta; case 005 (`transform: rotate`) stayed at 0.5%; no other case moved by >1pt.
- **`collect()` already skipped `opacity: 0` / `visibility: hidden` / `display: none` elements, but two faster paths bypassed it.** The UL/OL list shortcut in `prepareRenderItem` filters `<li>` children directly without recursing through `collect()`; the in-element text-container loop and the recursive `collectTextParts` (utils.js) likewise iterate children via `childNodes.forEach` and read `child.textContent`. Reveal.js fragments (`.fragment` LIs and SPAN chips) compute to `opacity: 0` + `visibility: hidden` until their click index is reached, so q-08 was shipping all 5 list items and all 5 chips even though the source PNG showed them hidden. Skip added at all three sites; q-08 dropped from ~56% → 35% content delta. Same fix knocks 1–3pt off q-09 through q-18 because every Quantyca content slide has a few hidden fragments that were leaking through. The visible-class detection from reveal print mode still works because `.fragment.visible` resolves opacity to 1 before measurement.

**Why first:** before these landed the only signal was "open PowerPoint and squint." Without a baseline that's robust to background-dominated slides, Phase 1's text fixes and Phase 5's RTL work had no acceptance criterion, and regressions were invisible.

**Out of scope:** pixel-perfect equality. Aim for "no perceptual surprise" — typically <10% content delta on text-only slides, higher tolerated on rasterized fallbacks and pseudo-element-heavy slides until Phase 1.2 lands.

---

## Phase 1 — Foundational text & layout correctness ✅ landed

All nine items below shipped together with six new fidelity cases (`010-ancestor-transform-scale` … `015-generic-sans-serif`). The full 46-case corpus passes; Phase 1 validation cases land at 0.7–2.5% delta. The Quantyca canary stayed within budget on every slide and most slides moved into a tighter band — text-heavy slides dropped 1–3 percentage points (e.g. quantyca-13 from 5.2% → 2.7%, quantyca-30/31/32 from 8.2% → 5.7%). `quantyca-27` (CSS-grid logo wall) and the icon-font slides (16, 18, 22–25, 28) remain >10% — those are out of Phase 1's scope and pinned to later phases.

The ancestor-transform fix (1.1) ended up as a generalized cumulative-transform tracker rather than a single ancestor read. The post-mortem's framing ("walk from root to its first transformed descendant") only described one direction, but real decks have transforms on *both* sides of root: reveal.js wraps an embedded deck in a scaled viewport (above root) AND scales `.slides` to fit (below root). The collect walk now multiplies each node's own scale into a `cumulative` factor and threads a per-node `styleScale = scale × ancestorScale × cumulative` into prepareRenderItem, so font-size / padding / border-width readings stay in sync regardless of where the transform lives.

Two additional fixes landed alongside Phase 1 that the post-mortem did not call out but that the canary surfaced once items 1.1–1.5 were live:

- **Font fallback chain quoting**: pptxgenjs concatenates `fontFace` straight into XML attribute values without escaping, so any embedded `"` (e.g. `"Segoe UI"` in a CSS font-family chain) produces invalid OOXML and a blank slide. `resolveFontFaceList` now strips quotes defensively before joining.
- **z-index stacking floor**: framework-driven decks (reveal.js) stamp a high explicit `z-index` on each `<section>`, and our flat render queue used to put the section's background fill *above* its descendant text. Children's effective z-index is now `max(child.zIndex, parent.zIndex)`, which approximates CSS stacking-context behavior closely enough for the cases we care about.

These bugs affect *every* export, so fixing them first amplifies the value of all later phases. The first three items are direct consequences of the Quantyca post-mortem; together they are projected to take the worst real-world slides from 12-71% diff into the 2-5% band.

1. **Compensate for ancestor CSS transforms.** *(New — top win for any framework-driven deck.)* `processSlide` (`src/index.js:193-208`) reads `root.getBoundingClientRect()` and trusts that descendant rects and computed styles share a frame. They don't: an ancestor `transform: scale(N)` makes `getBoundingClientRect()` post-transform but leaves `getComputedStyle().fontSize` / `padding` / `lineHeight` / `letterSpacing` / `borderWidth` at logical pre-transform values. Today every text run on a reveal.js slide ships ~30% too large for its box. Fix: walk from `root` to its first transformed descendant, decompose the cumulative ancestor transform matrix into a uniform `ancestorScale` (or `{sx, sy}` if non-uniform), and multiply every "logical CSS px" reading from `getComputedStyle` by it before the existing `× 0.75 × layoutConfig.scale` conversion. Positions/sizes from `getBoundingClientRect()` already account for the transform and need no change. Add a synthetic fidelity case (a `<div style="transform: scale(0.6)">` containing nested text and boxes) so this regression is caught without rerunning the full Quantyca corpus.
2. **Translate decorative `::before`/`::after` boxes.** *(Promoted from Phase 8.1.)* Today only the pseudo-element `content` string is captured (`utils.js:1024-1038`); the box itself (background, border, transform, width/height, position) is not, so the H2 underlines, corner brand marks, iceberg "LINEA D'ACQUA" badge, timeline dots, and matrix axis labels in the Quantyca deck all disappear. Detect when `getComputedStyle(el, '::before' | '::after')` produces a layout-sized box and emit it as its own render item with the parent's coordinate frame and z-index. (Computing the pseudo-element's actual rect requires either a one-off injected DOM clone or the `Element.getBoxQuads()` fallback path; pick one and document.)
3. **Map `line-through` to PPTX strikethrough.** `getTextStyle` (`src/utils.js:479`) only checks `underline`. Add `style.textDecorationLine.includes('line-through')` → PPTX `strike: 'sngStrike'`. One-line fix; clears `quantyca-18`'s entire "PRIMA — AS-IS" column.
4. **Await font loading before measuring.** `processSlide` (`src/index.js:193`) measures rects before `document.fonts.ready` resolves. Add a single `await document.fonts.ready` at the top of `exportToPptx` *and* a per-slide `await` in case fonts load mid-export.
5. **Preserve the font fallback chain — and embed icon fonts robustly.** `getTextStyle` (`src/utils.js:475`) ships only the first family. Send the full comma list to PPTX's `fontFace`, and resolve generics (`sans-serif` → `Calibri`, `serif` → `Cambria`, `monospace` → `Consolas`) at the boundary so PowerPoint substitutes when an embed is missing. Plus: `getAutoDetectedFonts` quietly skips fonts whose @font-face URL is CORS-blocked or hosted on a CDN with non-permissive headers (Font Awesome on `cdnjs` is the canonical case in our corpus). When detection fails for an icon font, either (a) refetch via `fetch(url, { mode: 'no-cors' })` and treat the opaque response as embeddable raw bytes, or (b) rasterize each used glyph at export time and emit it as an inline image. Otherwise icon-font runs render as `?`/`8` placeholder boxes (visible on `quantyca-03`, `-22`, `-25`).
6. **Honor `white-space`.** `src/index.js:1054` unconditionally collapses whitespace. Branch on computed `white-space`: for `pre`/`pre-wrap`/`pre-line`/`break-spaces`, preserve runs and convert tabs to either real `\t` (PPTX supports tab stops) or a configurable space count.
7. **Locale-aware `text-transform`.** Replace `.toUpperCase()` / `.toLowerCase()` (`utils.js:1058-1060`, `index.js:1064-1065`) with `.toLocaleUpperCase(documentElement.lang)` and a real Unicode word-boundary for `capitalize` (intl-segmenter).
8. **XML-escape every text value before it reaches PptxGenJS.** Strip C0 control chars (except `\t`, `\n`); escape `&`, `<`, `>` defensively. Prevents corrupt PPTX files that *look* like fidelity failures.
9. **Sub-pixel rounding policy.** Pick one and document it. Today, font sizes are intentionally fractional but rect positions are implicitly truncated by float→PPTX-EMU. Round positions to 1/8 pt to stop adjacent shapes from overlapping by 0.01" — visible as faint gaps in PowerPoint's renderer.

**Validation:** add fidelity cases for (a) a `transform: scale(0.6)` wrapper containing text + boxes, (b) an H2 with a `::after` underline, (c) a paragraph using `text-decoration: line-through` mixed with normal runs, (d) a Font Awesome icon next to native text, (e) `pre-wrap`, (f) Turkish `text-transform: uppercase`, and (g) a slide whose body uses only `font-family: sans-serif`.

---

## Phase 1.2 — Decorative pseudo-element boxes ✅ landed

`measurePseudoBox` (`src/utils.js`) and `pseudoToRenderItems` (`src/index.js`, renamed from the singular `pseudoToRenderItem`) were extended together to cover the patterns the Quantyca canary surfaced:

- **Auto width/height resolved from `inset` / `top+bottom` / `left+right`.** A pseudo with `position: absolute; inset: 0 0 0 55%; width: auto; height: auto` (the hero/divider full-bleed pattern) used to be skipped because `parseFloat('auto')` returned NaN. We now compute the box from the four side coordinates when width or height is `auto` and both opposite sides are set.
- **`background-image: url(...)` and `linear-gradient(...)`.** Pseudo backgrounds previously rendered only the solid `backgroundColor`. The url path queues a `getProcessedImage` job (same code path as `<img>` tags); the gradient path generates an SVG via `generateGradientSVG`. Both honor `background-size`, `background-position`, `border-radius` clipping, and the pseudo's own `opacity` (mapped to PPTX `transparency`).
- **`content: 'text'` rendered as overlaid run.** Decoded from the computed style (CSS hex escapes like `\d'acqua` are unescaped to literal Unicode), then emitted as a text item layered above any background fill. The `getTextStyle` extraction picks up font, color, letter-spacing, and text-transform.
- **`width: auto` for text-based pseudos** (corner brand mark, badge labels) — measured via a hidden absolutely-positioned proxy span that mirrors the pseudo's font/padding/border styles, since browsers don't expose pseudo rects directly. After the proxy yields a width and height, any side-anchored coordinate that depended on those dimensions is re-resolved.
- **Per-corner radii, percentage radii.** A new `parsePseudoRadii` accepts `%` (resolved against the smaller dimension, so `border-radius: 50%` on a 24×24 pseudo becomes a 12px radius → ellipse). When the four corners differ we route through `generateCustomShapeSVG` instead of the native `roundRect`.
- **`opacity: 0` / `visibility: hidden` skip.** Previously these would still emit. Now `measurePseudoBox` returns null up front.
- **Z-index resolution.** Pseudo's own `z-index` is floored at parent's effective z (matches CSS stacking-context behavior closely enough for the deck patterns we care about), then a small offset within the pseudo's items so its content text sits above its own background.

**Validation cases landed (`tests/fidelity/cases/`)**

- `016-pseudo-content-badge` — corner brand mark with `width: auto`, letter-spacing, text-transform; tests proxy measurement.
- `017-pseudo-circle-percent` — four `border-radius: 50%` dots with white/burgundy fills and a 3px stroke; tests percentage radii.
- `018-pseudo-inset-fullbleed` — `inset: 0 0 0 55%` with a gradient background and `opacity: 0.85`; tests auto-width-from-inset and gradient-on-pseudo.
- `019-pseudo-text-and-bg` — overlaid `NEW` pill with rounded background + centered text; tests the content-text-on-rect overlay.

All four pass at ≤52% content delta with the standard 85% budget; visual outputs match source closely.

**Harness fix landed alongside Phase 1.2.** `tests/fidelity/lib/browser.js` now stands up a small `node:http` static server on `127.0.0.1:8002` rooted at `cases/`, and case URLs are loaded as `http://127.0.0.1:8002/<rel>` instead of `file://`. Same-origin HTTP loads don't taint the canvas, so `Image` + `canvas.toDataURL` round-trips cleanly for SVG/PNG assets — file:// origins were silently blanking every `background-image: url(...)` on a pseudo and every `<img src="*.svg">` in the corpus (q-06's six service icons were blank for that reason). The server starts lazily on first case and is torn down in `closeBrowser`. Port 8002 was chosen to leave 8001 free for `python3 -m http.server 8001` over `tests/fidelity/` to view the report at `/report/index.html`.

**Score snapshot after both fixes (50 cases)**

| slide | content Δ before | content Δ after |
|---|---|---|
| q-01 hero (page-header.svg) | 45.4% | 27.7% |
| q-04 (hero-style + icons) | 51.3% | 22.9% |
| q-06 (six service-icon SVGs) | 25.9% | 17.0% |
| q-07 | 50.5% | 22.3% |
| q-26 | 52.1% | 24.2% |
| q-28 | 51.4% | 24.3% |

The remaining stubborn cluster (q-08, q-09, q-19–22, q-27) is dominated by code-block syntax highlighting (`<pre><code>` + hljs spans) which the engine doesn't render natively — Phase 9 territory. 011-pseudo-after-underline now lands at 37% content delta (was passing lower previously) — the new accent-bar render adds edge density that the metric counts even though the underline visually matches.

All 50 cases pass under the 65% Quantyca budget / 85% default budget.

---

## Phase 2 — Color, shadow, and border primitives

Visible defects on almost every styled component.

1. **Multi-shadow stacking.** `getVisibleShadow` (`utils.js:694`) returns after the first non-transparent shadow. Return an *array*; emit one PPTX shadow per layer when the shape API supports it, else composite into a pre-rendered SVG behind the shape.
2. **Inset shadows.** `utils.js:713` always sets `type: 'outer'`. When the parsed token list contains `inset`, set `type: 'inner'`.
3. **Outline as stroke.** Read `outline-{width,style,color,offset}` and emit a second non-filled shape inset/outset by `outline-offset` if a real outline channel isn't available.
4. **Transparent borders.** Today they collapse to zero width because alpha→0 erases them. Either keep the width and skip the stroke, or document the trade-off; many CSS resets rely on `border: 1px solid transparent` for layout.
5. **8-digit hex with 4-digits-per-channel.** `parseColor` (`utils.js:378-380`) handles `#rrggbbaa` but not `#rrrrggggbbbbaaaa`. Add the branch.
6. **Wide-gamut color spaces.** `oklch`, `lab`, `display-p3` are normalized through a hidden canvas (`utils.js:400-407`), which silently clips to sRGB. Document the clipping; opt-in `options.preserveWideGamut: true` could keep the original color string for PPTX 2019+ readers, but that's optional. The minimum is to stop emitting black on parse failure (`utils.js:409-412`) and instead fall back to the *unparsed* CSS string for diagnostics.

---

## Phase 3 — Geometry: transforms, radii, scroll

1. **Full 2D transform decomposition.** `getRotation` (`utils.js:566`) reads only rotation. Decompose the matrix into translate/scale/rotate/skew (the `decomposeMatrix2D` recipe from CSS Transforms spec); apply translate by adjusting x/y in inches, scale by adjusting w/h, rotate by the existing channel, and skew by either an emulated SVG (small angles) or rasterizing.
2. **`transform-origin`.** Today rotation pivots around the rect center. Read `transformOrigin` and shift the post-rotation rect so the origin point matches the source. PPTX rotates around the shape center, so we need to compensate the x/y offset.
3. **Elliptical border-radius.** Parse the `/` form (`10px / 5px`) and per-corner pairs. Generate a custom-geometry SVG via `generateCustomShapeSVG` when radii are non-uniform; PPTX `prstGeom: roundRect` only supports uniform radius.
4. **Source-element scroll normalization.** `getBoundingClientRect` returns viewport-relative coordinates, so a scrolled container drops content above the fold. At the start of `processSlide`, capture `target.scrollLeft/scrollTop` and add them to the rect math; also temporarily set `scrollTop = 0` if `options.includeOverflow` is true.
5. **Recursion safety.** Cap `collect()` (`src/index.js:215-265`) at a configurable depth (default 1024) to prevent stack overflow on pathological DOMs.

---

## Phase 4 — Gradients and layered fills

1. **Radial gradients.** Extend `generateGradientSVG` (`utils.js:728`) with a `radial-gradient` parser; emit `<radialGradient>` with `cx/cy/r` from the CSS shape (`circle`/`ellipse`) and extent keywords (`closest-corner`, `farthest-side`).
2. **Conic gradients.** No SVG primitive — emit a high-resolution PNG by drawing on an offscreen canvas, embed as a background image. Document as raster.
3. **Repeating gradients.** Translate to a finite stop list spanning the visible rect with the repeat unit tiled.
4. **Multi-layer backgrounds.** `src/index.js:1004` keeps only the first `url()`. Compose all layers (image, gradient, color) into a single offscreen canvas at 2× resolution, embed as one background image. This subsumes `background-blend-mode` for free if we honor `mix-blend-mode` on the canvas context.
5. **`background-clip`, `background-origin`, `background-size`, `background-position`.** Today these are partially honored only via `object-fit` for `<img>`. Apply them in the composer above so a `background: url(...) center/contain no-repeat` produces the right rectangle.

---

## Phase 5 — Text richness

1. **Box-shadow on text vs `text-shadow`.** Confirm `text-shadow` flows through; if not, map at least a single shadow to PPTX text effect.
2. **`text-decoration` styled variants.** (Plain `line-through` covered in Phase 1.3.) Read `text-decoration-style` (dotted/dashed/wavy), `-thickness`, `-color`. PPTX exposes underline styles (`u: 'sng'`, `'dbl'`, `'wavy'`, `'dotted'`, `'dash'`) — map them directly.
3. **`letter-spacing` in tables.** `extractTableData` (`utils.js:73`) doesn't pass `charSpacing` to cells. Mirror the main text-style extraction.
4. **Hyperlinks.** Walk ancestors of each text run; if an `<a href>` wraps it, attach `hyperlink: { url, tooltip: a.title }` to the PPTX text run.
5. **`font-variant` & `font-feature-settings`.** Map small-caps to PPTX `cap: 'small'`, tabular-nums to a font-feature run property where supported. Where unsupported, document.
6. **BiDi and `direction: rtl`.** Read `direction` and set the run's `rtl: true`. Set the slide's writing direction when the root has `dir="rtl"`.
7. **Text fragmentation across mixed-style runs.** Re-audit `collectTextParts` (`utils.js:1024+`) for inline images, `<sub>`/`<sup>` (PPTX `baseline: 30`/`-25`), and inline `<svg>` icons.

---

## Phase 6 — Images and SVG

1. **`srcset`/`<picture>`.** Resolve `currentSrc` instead of `src` (`src/index.js:862`) so the right resolution variant is exported.
2. **EXIF orientation.** Use `image-orientation: from-image` semantics — read EXIF orientation in `getProcessedImage` and rotate the canvas accordingly.
3. **CORS-tainted fallback.** When `canvas.toDataURL` throws (`src/index.js:819-834`), don't drop silently. Emit a placeholder rect of the same dimensions with the image's `alt` text inside, and a single `console.warn` per export listing affected URLs.
4. **Inline SVG with `currentColor`.** When emitting SVG as vector (`utils.js:630-657`), substitute `currentColor` with the inherited text color *before* serialization, otherwise PPTX renderers will use black.
5. **`image-rendering: pixelated`.** Pass to the canvas as `imageSmoothingEnabled = false` in `getProcessedImage` so logos/pixel art stay crisp at the 2× upscale.

---

## Phase 7 — Tables

1. **Section ordering.** `extractTableData` (`utils.js:66`) globs `tr`. Walk `thead → tbody → tfoot` explicitly and concatenate.
2. **`<caption>`.** Extract and emit as a separate text shape positioned by `caption-side`.
3. **Row-inherited backgrounds.** Read `<tr>` background and merge into each cell's fill before reading the cell's own background.
4. **`border-collapse: collapse`.** Document the limitation, but at minimum collapse adjacent identical borders so the output doesn't show double-thickness lines.
5. **Cells containing block content.** When a `<td>` contains lists, nested tables, or block-level elements, fall back to rendering the cell as an *image* (html2canvas snapshot) inside the table cell rather than flattening to text.
6. **colspan/rowspan validation.** Clamp values to the actual grid size before passing to PptxGenJS — out-of-range values currently break the table layout.

---

## Phase 8 — Decorations, advanced layout

(Decorative `::before`/`::after` translation moved to Phase 1 item 2 after the Quantyca canary showed it hits every content slide of a real-world deck.)

1. **`::marker` font and size.** `index.js:671-676` reads color and type only. Pass marker font/size through.
2. **`clip-path: polygon(...)` / `circle(...)` / `inset(...)`.** Translate to PPTX custom geometry via `generateCustomShapeSVG` for polygons, `prstGeom: ellipse` for circles.
3. **Multi-column (`columns`).** Children are already absolutely positioned by the browser, so the layout *should* survive — add a fidelity case to confirm and document any drift.

---

## Phase 9 — Filters, blends, and the rasterization escape hatch

This is the "we can't translate it natively" bucket.

1. **`mix-blend-mode`, `backdrop-filter` beyond blur, `filter` (other than `none`).** Trigger an html2canvas snapshot of the element + a small bleed margin around it, embed as an image at the element's z-index. Today this only happens for `backdrop-filter: blur(...)` (`src/index.js:1017`); generalize.
2. **A single `shouldRasterize(el, style)` predicate.** Centralize the rules; today rasterization decisions are scattered. Returning `true` queues an html2canvas job and skips native translation for the subtree.
3. **Rasterization quality knob.** `options.rasterScale` (default `2 * devicePixelRatio`, capped at 4) so users can trade size for sharpness.

---

## Phase 10 — Documentation & honesty

For every gap not closed by phases 1–9, update `SUPPORTED.md` with: feature, current behavior (dropped silently / approximated / rasterized), and a one-line workaround. Fidelity is partly an expectations problem — silent drops are worse than documented ones.

Add a `FIDELITY.md` linking to the test corpus and the latest report so users can see, per feature, the current pixel delta.

---

## Cross-cutting refactors picked up along the way

These aren't standalone tasks but should land *with* the phase that first needs them, not as a separate "cleanup" PR:

- A `styledLayer({ shape, fills[], strokes[], shadows[], clip })` builder. Today `prepareRenderItem` re-implements stacking inline; multi-shadow + multi-background + clip-path will all need the same composition.
- A `measureBox(el)` helper that returns `{ rect, scrollOffset, transform, transformOrigin, opacity }` so transform/scroll fixes don't have to touch every call site of `getBoundingClientRect`.
- A `warn(category, message, context)` channel collected on the returned blob's metadata so users can see what was approximated, not just what blew up.

---

## Suggested sequencing

Phases 0, 1, 1.1 re-validation, 1.2, and the opacity:0 skip have shipped. Next up is Phase 2 (color/shadow/border primitives) and the *minimal* Phase 9 (generalized rasterization fallback behind a `shouldRasterize` predicate).

| Order | Phase / item | Why this slot |
|---|---|---|
| 1 | 0 (test harness) ✅ | Everything else needs verification |
| 2 | 1.1–1.9 (foundational text & layout) ✅ | Highest blast radius per LOC |
| 3 | Phase 0 addendum — switch canary to reveal print mode ✅ | Removed framework-coordination noise from the canary |
| 4 | Phase 0 addendum — foreground-aware metric ✅ | Replaced raw pixel-% with `contentPercent` (edge + color, block-level) so background-dominated slides stop masking foreground regressions |
| 5 | Phase 1.2 — decorative pseudo-element boxes ✅ | Corner brand mark, H2 underline, hero/divider full-bleeds, badges, percent-radius dots all translate |
| 6 | Phase 1.1 re-validation + `opacity: 0` skip ✅ | offsetWidth × styleScale fixed case 010; opacity skip in UL/text-container/collectTextParts dropped q-08 by ~21pt |
| 6.5 | Harness: serve cases via local HTTP server ✅ | `127.0.0.1:8001` static server replaces file://; SVG icons + url-background pseudos now round-trip cleanly |
| 6.6 | Phase 0 addendum — synthetic corpus 19 → 30 ✅ | Cases 020–030 cover Phase 2/3 territory; six clean regression guards, five flag known gaps with 4–66pt budget headroom |
| 7 | 2 (color/shadow/border) | Visible on most decks; isolated changes |
| 8 | 9 (rasterization escape hatch, *minimal*) | Unblocks "ship something acceptable for unsupported CSS" |
| 9 | 4 (gradients/backgrounds) | Common, currently silently dropped |
| 10 | 3 (transforms/geometry) | Affects polished decks (rotated badges, custom radii) |
| 11 | 5 (text richness) | Hyperlinks alone are a frequent ask |
| 12 | 6 (images/SVG) | Mostly polish; CORS placeholder is the bug |
| 13 | 7 (tables) | High effort, narrower audience |
| 14 | 8 (clip-path / `::marker` / multi-column) | Niche, high engineering cost (pseudo-element decoration boxes already pulled forward) |
| 15 | 10 (docs) | Continuous, but final SUPPORTED.md pass closes the loop |

The reshuffle reflects what the 46-case run under the new metric actually told us: with the canary on print mode and the metric corrected, residual deltas concentrate on a small number of engine gaps (pseudo-element decoration, fragment opacity, ancestor-transform on the export root). Phase 1.2 alone is projected to drop the q-19–22 cluster from ~55% to ~30% content delta.

Phase 9's *minimal* version (just generalize the existing html2canvas fallback behind a predicate) still lands early so unsupported features degrade gracefully rather than disappearing while later phases are in flight.

---

## Success metrics

- **Coverage:** every CSS feature listed in SUPPORTED.md has at least one fidelity case.
- **Regression gate:** CI fails if any case's `contentPercent` increases by more than 1pt versus the previous main. (Implemented as a follow-up TODO in Phase 0.)
- **Silent-drop rate:** the warn channel reports zero "feature not implemented" entries on the standard corpus by end of Phase 5.
- **Issue triage:** open GitHub fidelity issues categorized to a phase; phase completion = all its issues closed or explicitly deferred in SUPPORTED.md.
