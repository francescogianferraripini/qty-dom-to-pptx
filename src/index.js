// src/index.js
import * as PptxGenJSImport from 'pptxgenjs';
import html2canvas from 'html2canvas';
import { PPTXEmbedFonts } from './font-embedder.js';
import JSZip from 'jszip';

// Normalize import
const PptxGenJS = PptxGenJSImport?.default ?? PptxGenJSImport;

import {
  parseColor,
  getTextStyle,
  isTextContainer,
  parseShadowList,
  shadowLayerToPptx,
  composeOuterShadows,
  composeInsetShadows,
  getCornerRadii,
  getCornerRadiiXY,
  isElliptical,
  isPerCorner,
  generateGradientSVG,
  getRotation,
  hasSkew,
  getWritingModeVert,
  svgToPng,
  svgToSvg,
  getPadding,
  paddingToPptxMargin,
  getSoftEdges,
  generateBlurredSVG,
  getBorderInfo,
  generateCompositeBorderSVG,
  isClippedByParent,
  generateCustomShapeSVG,
  getUsedFontFamilies,
  getAutoDetectedFonts,
  extractTableData,
  collectTextParts,
  applyTextTransform,
  sanitizeText,
  processWhitespace,
  getAncestorScale,
  nodeOwnScale,
  measurePseudoBox,
  resolveContainerAlignment,
} from './utils.js';
import { getProcessedImage } from './image-processor.js';

const PPI = 96;
const PX_TO_INCH = 1 / PPI;

// Per-element skew warnings. PPTX shapes can't represent CSS skew transforms,
// so a skewed element renders as its un-skewed rectangle. The warning helps
// users diagnose visible distortion. WeakSet persists across exports so we
// don't repeat warnings for the same element on re-runs.
const _warnedSkew = new WeakSet();
function _warnSkewOnce(node, transformStr) {
  if (_warnedSkew.has(node)) return;
  _warnedSkew.add(node);
  console.warn(
    '[dom-to-pptx] CSS skew detected (transform:',
    transformStr +
      ') — PPTX shapes cannot represent skew, element will render un-skewed. Rasterize manually if exact match is required.',
    node
  );
}

/**
 * Phase 2 multi-shadow planner. Reads box-shadow into a list of layers and
 * decides how to render them in PPTX:
 *   - 0 layers           → no shadow.
 *   - 1 outer (no inset) → native pptxgenjs `shadow:` option (existing path).
 *   - 1 inner (no outer) → native `shadow:` with type 'inner'.
 *   - everything else    → composite outer halo behind shape and/or inner
 *                          halo above shape, both as PNG images. The native
 *                          shape carries no shadow option in this branch.
 *
 * Returns `{ primaryShadow, outerImage, innerImage }`. `outerImage` and
 * `innerImage` are `{ dataUrl, paddingPx } | null`. paddingPx is in CSS px.
 */
function planShadows(shadowStr, widthPx, heightPx, radii, scale) {
  const list = parseShadowList(shadowStr);
  if (!list.length) return { primaryShadow: null, outerImage: null, innerImage: null };
  const outer = list.filter((s) => !s.inset);
  const inner = list.filter((s) => s.inset);
  let primary = null;
  let outerToComp = outer;
  // Inset shadows always go through the canvas-composite path. PowerPoint
  // accepts <a:innerShdw> in the OOXML, but LibreOffice's PPTX renderer drops
  // it silently — we'd ship a slide that looks correct in PowerPoint and
  // wrong in every other viewer. Compositing is portable.
  const innerToComp = inner;
  if (outer.length === 1 && inner.length === 0) {
    primary = shadowLayerToPptx(outer[0], scale);
    outerToComp = [];
  }
  const outerImage = outerToComp.length
    ? composeOuterShadows(widthPx, heightPx, radii, outerToComp)
    : null;
  const innerImage = innerToComp.length
    ? composeInsetShadows(widthPx, heightPx, radii, innerToComp)
    : null;
  return { primaryShadow: primary, outerImage, innerImage };
}

/**
 * Main export function.
 * @param {HTMLElement | string | Array<HTMLElement | string>} target
 * @param {Object} options
 * @param {string} [options.fileName]
 * @param {boolean} [options.skipDownload=false] - If true, prevents automatic download
 * @param {Object} [options.listConfig] - Config for bullets
 * @param {boolean} [options.svgAsVector=false] - If true, keeps SVG as vector (for Convert to Shape in PowerPoint)
 * @returns {Promise<Blob>} - Returns the generated PPTX Blob
 */
export async function exportToPptx(target, options = {}) {
  const resolvePptxConstructor = (pkg) => {
    if (!pkg) return null;
    if (typeof pkg === 'function') return pkg;
    if (pkg && typeof pkg.default === 'function') return pkg.default;
    if (pkg && typeof pkg.PptxGenJS === 'function') return pkg.PptxGenJS;
    if (pkg && pkg.PptxGenJS && typeof pkg.PptxGenJS.default === 'function')
      return pkg.PptxGenJS.default;
    return null;
  };

  const PptxConstructor = resolvePptxConstructor(PptxGenJS);
  if (!PptxConstructor) throw new Error('PptxGenJS constructor not found.');
  const pptx = new PptxConstructor();

  // Make sure web fonts have actually loaded before we measure anything.
  // getComputedStyle returns the *requested* font-family even when the font
  // hasn't downloaded yet, which leaves text measured against the fallback
  // metrics and produces shifted layouts in the export.
  if (typeof document !== 'undefined' && document.fonts && document.fonts.ready) {
    try {
      await document.fonts.ready;
    } catch {
      /* swallow — fonts.ready can reject in odd environments */
    }
  }

  // 1. Layout Handling
  let finalWidth = 10; // default 16:9
  let finalHeight = 5.625;

  if (options.width && options.height) {
    pptx.defineLayout({ name: 'CUSTOM', width: options.width, height: options.height });
    pptx.layout = 'CUSTOM';
    finalWidth = options.width;
    finalHeight = options.height;
  } else if (options.layout) {
    pptx.layout = options.layout;
    // Map standard layouts for internal scale calculation if possible,
    // though PptxGenJS defaults to 16:9 if unknown.
    if (options.layout === 'LAYOUT_4x3') {
      finalWidth = 10;
      finalHeight = 7.5;
    } else if (options.layout === 'LAYOUT_16x10') {
      finalWidth = 10;
      finalHeight = 6.25;
    } else if (options.layout === 'LAYOUT_WIDE') {
      finalWidth = 13.3;
      finalHeight = 7.5;
    }
  } else {
    pptx.layout = 'LAYOUT_16x9';
  }

  // Pass these dimensions to options so processSlide can use them
  const extendedOptions = {
    ...options,
    _slideWidth: finalWidth,
    _slideHeight: finalHeight,
  };

  const elements = Array.isArray(target) ? target : [target];

  for (const el of elements) {
    const root = typeof el === 'string' ? document.querySelector(el) : el;
    if (!root) {
      console.warn('Element not found, skipping slide:', el);
      continue;
    }
    const slide = pptx.addSlide();
    await processSlide(root, slide, pptx, extendedOptions);
  }

  // 3. Font Embedding Logic
  let finalBlob;
  let fontsToEmbed = options.fonts || [];

  if (options.autoEmbedFonts) {
    // A. Scan DOM for used font families
    const usedFamilies = getUsedFontFamilies(elements);

    // B. Scan CSS for URLs matches
    const detectedFonts = await getAutoDetectedFonts(usedFamilies);

    // C. Merge (Avoid duplicates)
    const explicitNames = new Set(fontsToEmbed.map((f) => f.name));
    for (const autoFont of detectedFonts) {
      if (!explicitNames.has(autoFont.name)) {
        fontsToEmbed.push(autoFont);
      }
    }

    if (detectedFonts.length > 0) {
      console.log(
        'Auto-detected fonts:',
        detectedFonts.map((f) => f.name)
      );
    }
  }

  if (fontsToEmbed.length > 0) {
    // Generate initial PPTX
    const initialBlob = await pptx.write({ outputType: 'blob' });

    // Load into Embedder
    const zip = await JSZip.loadAsync(initialBlob);
    const embedder = new PPTXEmbedFonts();
    await embedder.loadZip(zip);

    // Fetch and Embed
    for (const fontCfg of fontsToEmbed) {
      try {
        const response = await fetch(fontCfg.url);
        if (!response.ok) throw new Error(`Failed to fetch ${fontCfg.url}`);
        const buffer = await response.arrayBuffer();

        // Infer type
        const ext = fontCfg.url.split('.').pop().split(/[?#]/)[0].toLowerCase();
        let type = 'ttf';
        if (['woff', 'otf'].includes(ext)) type = ext;

        await embedder.addFont(fontCfg.name, buffer, type);
      } catch (e) {
        console.warn(`Failed to embed font: ${fontCfg.name} (${fontCfg.url})`, e);
      }
    }

    await embedder.updateFiles();
    finalBlob = await embedder.generateBlob();
  } else {
    // No fonts to embed
    finalBlob = await pptx.write({ outputType: 'blob' });
  }

  // 4. Output Handling
  // If skipDownload is NOT true, proceed with browser download
  if (!options.skipDownload) {
    const fileName = options.fileName || 'export.pptx';
    const url = URL.createObjectURL(finalBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Always return the blob so the caller can use it (e.g. upload to server)
  return finalBlob;
}

/**
 * Worker function to process a single DOM element into a single PPTX slide.
 * @param {HTMLElement} root - The root element for this slide.
 * @param {PptxGenJS.Slide} slide - The PPTX slide object to add content to.
 * @param {PptxGenJS} pptx - The main PPTX instance.
 */
async function processSlide(root, slide, pptx, globalOptions = {}) {
  // Per-slide await in case fonts loaded mid-export (e.g. a deck navigates
  // to a new slide whose CSS pulls a fresh @font-face).
  if (typeof document !== 'undefined' && document.fonts && document.fonts.ready) {
    try {
      await document.fonts.ready;
    } catch {
      /* ignore */
    }
  }
  const rootRect = root.getBoundingClientRect();
  const PPTX_WIDTH_IN = globalOptions._slideWidth || 10;
  const PPTX_HEIGHT_IN = globalOptions._slideHeight || 5.625;

  // Phase 3.4 — scroll normalization. `getBoundingClientRect()` is viewport-
  // relative, so a scrolled root puts above-fold children at negative rect.top
  // (clipped off-slide) and leaves below-fold children at viewport bottom.
  // Default behavior preserves "export what's visible" semantics: the slide
  // shows the current scrolled viewport. `options.includeOverflow=true`
  // captures the FULL scrollable content area: content size becomes
  // root.scrollWidth/Height and the origin is back-shifted by the current
  // scroll offsets so above-fold content lands inside the slide.
  const rootScrollLeft = root.scrollLeft || 0;
  const rootScrollTop = root.scrollTop || 0;
  const includeOverflow = !!globalOptions.includeOverflow;
  const contentWidthPx = includeOverflow
    ? Math.max(rootRect.width, root.scrollWidth || 0)
    : rootRect.width;
  const contentHeightPx = includeOverflow
    ? Math.max(rootRect.height, root.scrollHeight || 0)
    : rootRect.height;

  const contentWidthIn = contentWidthPx * PX_TO_INCH;
  const contentHeightIn = contentHeightPx * PX_TO_INCH;
  const scale = Math.min(PPTX_WIDTH_IN / contentWidthIn, PPTX_HEIGHT_IN / contentHeightIn);

  // Compensate for CSS transforms — both ABOVE root (reveal.js wraps the deck
  // in a `transform: scale(N)` viewport) and BELOW root inside the subtree
  // (reveal.js applies a fitting transform to `.slides`). `getBoundingClientRect`
  // is post-transform but `getComputedStyle` returns logical (pre-transform)
  // px, so without correction every font-size/padding/border-width is read
  // ~1/N too large for the on-screen rectangle. `scale` continues to be used
  // for rect-derived positions and sizes (already post-transform); `styleScale`
  // is the multiplier we apply to every CSS-px reading before the existing
  // px→pt conversion. Strict-ancestor transforms produce a uniform multiplier;
  // descendant transforms compound with each step into the subtree, tracked
  // during traversal as `cumulative`.
  const ancestorScale = getAncestorScale(root);
  // styleScale at root (no descendant transform applied yet). Helpers that
  // measure root itself can use this; the per-node styleScale is computed
  // inside `collect`.
  const rootStyleScale = scale * ancestorScale;

  const layoutConfig = {
    // When `includeOverflow` is set, treat the root's content origin (i.e.
    // where children sit at scrollLeft/scrollTop = 0) as the slide origin so
    // scrolled-out content maps inside the slide.
    rootX: includeOverflow ? rootRect.x - rootScrollLeft : rootRect.x,
    rootY: includeOverflow ? rootRect.y - rootScrollTop : rootRect.y,
    scale: scale,
    styleScale: rootStyleScale,
    ancestorScale: ancestorScale,
    offX: (PPTX_WIDTH_IN - contentWidthIn * scale) / 2,
    offY: (PPTX_HEIGHT_IN - contentHeightIn * scale) / 2,
  };

  const renderQueue = [];
  const asyncTasks = []; // Queue for heavy operations (Images, Canvas)
  let domOrderCounter = 0;

  // Cap recursion to defend against pathological DOMs (deeply-nested
  // user-generated markup, runaway templates). Each call to `collect`
  // pushes a stack frame; without a cap, JS engines blow up around 10k.
  // 1024 is well below the engine limit but high enough that no real
  // document hits it. Override via `options.maxDomDepth` for testing.
  const MAX_DOM_DEPTH = globalOptions.maxDomDepth || 1024;
  let depthOverflowReported = false;

  // Sync Traversal Function
  function collect(node, parentZIndex, parentCumulative, depth = 0) {
    if (depth >= MAX_DOM_DEPTH) {
      if (!depthOverflowReported) {
        console.warn(
          `[dom-to-pptx] DOM depth exceeded ${MAX_DOM_DEPTH} — bailing out of recursion. ` +
            `Some content may be missing from the export.`
        );
        depthOverflowReported = true;
      }
      return;
    }
    const order = domOrderCounter++;

    let currentZ = parentZIndex;
    let nodeStyle = null;
    let cumulative = parentCumulative;
    const nodeType = node.nodeType;

    if (nodeType === 1) {
      nodeStyle = window.getComputedStyle(node);
      // Optimization: Skip completely hidden elements immediately
      if (
        nodeStyle.display === 'none' ||
        nodeStyle.visibility === 'hidden' ||
        nodeStyle.opacity === '0'
      ) {
        return;
      }
      if (nodeStyle.zIndex !== 'auto') {
        // Floor at parent. CSS treats a child's z-index as local to its
        // parent's stacking context, so a parent.zIndex=11 with child.zIndex=2
        // still renders the parent BEHIND the child within that context. Our
        // pipeline flattens everything into a single queue, so we clamp the
        // child's effective z so its bg/text never slides below an ancestor's
        // fill. This matters for framework-driven decks (reveal.js stamps a
        // high z-index on each <section>).
        const ownZ = parseInt(nodeStyle.zIndex);
        if (!isNaN(ownZ)) currentZ = Math.max(ownZ, parentZIndex);
      }
      // Pick up this node's own scale transform. Compounds with the parent
      // chain so descendants inside e.g. `.slides { transform: scale(0.75) }`
      // get their logical font-size multiplied by 0.75 before px→pt.
      const ownScale = nodeOwnScale(nodeStyle.transform);
      if (ownScale !== 1) cumulative *= ownScale;
    }

    const nodeStyleScale = scale * ancestorScale * cumulative;
    const nodeConfig = {
      ...layoutConfig,
      root,
      styleScale: nodeStyleScale,
      cumulativeTransform: cumulative,
    };

    // Prepare the item. If it needs async work, it returns a 'job'
    const result = prepareRenderItem(
      node,
      nodeConfig,
      order,
      pptx,
      currentZ,
      nodeStyle,
      globalOptions
    );

    if (result) {
      if (result.items) {
        // Push items immediately to queue (data might be missing but filled later)
        renderQueue.push(...result.items);
      }
      if (result.job) {
        // Push the promise-returning function to the task list
        asyncTasks.push(result.job);
      }
      if (result.stopRecursion) {
        // Even when a node owns its rendering (e.g. text container with inline
        // styling) decorative pseudo-element boxes still belong to the slide.
        emitPseudoItems(node, currentZ, order, cumulative);
        return;
      }
    }

    if (nodeType === 1) emitPseudoItems(node, currentZ, order, cumulative);

    // Recurse children synchronously, propagating cumulative transform.
    const childNodes = node.childNodes;
    for (let i = 0; i < childNodes.length; i++) {
      collect(childNodes[i], currentZ, cumulative, depth + 1);
    }
  }

  function emitPseudoItems(parent, parentZ, parentOrder, parentCumulative) {
    if (!parent || parent.nodeType !== 1) return;
    if (globalOptions.disablePseudoElements) return;
    const pseudoStyleScale = scale * ancestorScale * parentCumulative;
    for (const which of ['before', 'after']) {
      const measured = measurePseudoBox(parent, which);
      if (!measured) continue;
      const result = pseudoToRenderItems(
        measured,
        { ...layoutConfig, styleScale: pseudoStyleScale },
        pptx,
        parentZ,
        parentOrder
      );
      if (result?.items?.length) renderQueue.push(...result.items);
      if (result?.job) asyncTasks.push(result.job);
    }
  }

  // 1. Traverse and build the structure (Fast). The third arg is the
  // cumulative transform from above-root *down to but not including* this
  // node — root starts with `1` and descendants compound their own scale
  // transforms into it.
  collect(root, 0, 1);

  // 2. Execute all heavy tasks in parallel (Fast)
  if (asyncTasks.length > 0) {
    await Promise.all(asyncTasks.map((task) => task()));
  }

  // 3. Cleanup and Sort
  // Remove items that failed to generate data (marked with skip)
  const finalQueue = renderQueue.filter(
    (item) => !item.skip && (item.type !== 'image' || item.options.data)
  );

  finalQueue.sort((a, b) => {
    if (a.zIndex !== b.zIndex) return a.zIndex - b.zIndex;
    return a.domOrder - b.domOrder;
  });

  // 4. Add to Slide
  for (const item of finalQueue) {
    snapItemDimensions(item);
    if (item.type === 'shape') slide.addShape(item.shapeType, item.options);
    if (item.type === 'image') slide.addImage(item.options);
    if (item.type === 'text') slide.addText(item.textParts, item.options);
    if (item.type === 'table') {
      slide.addTable(item.tableData.rows, {
        x: item.options.x,
        y: item.options.y,
        w: item.options.w,
        h: item.options.h,
        colW: item.tableData.colWidths, // Essential for correct layout
        autoPage: false,
        // Remove default table styles so our extracted CSS applies cleanly
        border: { type: 'none' },
        fill: { color: 'FFFFFF', transparency: 100 },
      });
    }
  }
}

// Reduce an elliptical-radii object (`{tl: {x, y}, ...}`) to scalar form
// (single px per corner) for code paths that don't render elliptical arcs —
// shadow-composite halos, native PPTX `roundRect`, and `tracePath`. Picks
// the larger of rx/ry per corner so the visible curve isn't undershot.
function radiiXYToScalar(r) {
  return {
    tl: Math.max(r.tl.x, r.tl.y),
    tr: Math.max(r.tr.x, r.tr.y),
    br: Math.max(r.br.x, r.br.y),
    bl: Math.max(r.bl.x, r.bl.y),
  };
}

// Turn a measurePseudoBox() result into one or more render items. A
// decorative pseudo can produce: a background image (url or gradient), a
// fill/stroke shape, and an overlaid text run for `content: 'text'`. We
// return them in z-order; an async job is returned when a url-based
// background needs to be fetched and processed.
function pseudoToRenderItems(measured, layoutConfig, pptx, parentZ, domOrder) {
  const { rect, style, contentText } = measured;
  const widthPx = rect.width;
  const heightPx = rect.height;
  if (widthPx < 0.5 || heightPx < 0.5) return null;

  const w = widthPx * PX_TO_INCH * layoutConfig.scale;
  const h = heightPx * PX_TO_INCH * layoutConfig.scale;
  const x =
    layoutConfig.offX + (rect.left - layoutConfig.rootX) * PX_TO_INCH * layoutConfig.scale;
  const y =
    layoutConfig.offY + (rect.top - layoutConfig.rootY) * PX_TO_INCH * layoutConfig.scale;

  // Effective z-index: by default a pseudo sits at the parent's stacking
  // level, so floor at parentZ. An explicit z-index on the pseudo can lift
  // it above siblings (reveal's hero/divider use this to layer the
  // page-header.svg behind the headline).
  const ownZRaw = style.zIndex;
  const ownZ = ownZRaw && ownZRaw !== 'auto' ? parseInt(ownZRaw) : NaN;
  const baseZ = isNaN(ownZ) ? parentZ : Math.max(ownZ, parentZ);

  const opacityNum = parseFloat(style.opacity);
  const safeOpacity = isNaN(opacityNum) ? 1 : opacityNum;

  const bg = parseColor(style.backgroundColor);
  const hasBg = bg.hex && bg.opacity > 0;
  const borderInfo = getBorderInfo(style, layoutConfig.styleScale);
  const hasUniformBorder = borderInfo.type === 'uniform';
  const rotation = getRotation(style.transform);

  const bgImg = style.backgroundImage;
  const hasBgImage = bgImg && bgImg !== 'none';
  const hasGradient = hasBgImage && bgImg.includes('gradient(');
  const urlMatch = hasBgImage && !hasGradient ? bgImg.match(/url\(['"]?(.*?)['"]?\)/) : null;
  const hasBgUrl = !!urlMatch;

  const radiiXY = getCornerRadiiXY(style, widthPx, heightPx);
  const radiiScalar = radiiXYToScalar(radiiXY);
  const needsCustomShape = isElliptical(radiiXY) || isPerCorner(radiiXY);
  const minDim = Math.min(widthPx, heightPx);

  // Shadow halos approximate elliptical corners with the larger axis — see
  // `radiiXYToScalar`. Phase 3.3's elliptical work is intentionally scoped
  // out of the shadow composite (per the IMPROVEMENT_PLAN out-of-scope note).
  const shadowPlan = planShadows(
    style.boxShadow,
    widthPx,
    heightPx,
    radiiScalar,
    layoutConfig.styleScale
  );
  const shadow = shadowPlan.primaryShadow;
  const hasShadowVisual = !!(shadow || shadowPlan.outerImage || shadowPlan.innerImage);

  const items = [];
  let job = null;
  let zCursor = baseZ;

  // 1a. Outer multi-shadow halo behind the shape.
  if (shadowPlan.outerImage) {
    const padIn = shadowPlan.outerImage.paddingPx * PX_TO_INCH * layoutConfig.styleScale;
    items.push({
      type: 'image',
      zIndex: zCursor,
      domOrder,
      options: {
        data: shadowPlan.outerImage.dataUrl,
        x: x - padIn,
        y: y - padIn,
        w: w + padIn * 2,
        h: h + padIn * 2,
        rotate: rotation,
      },
    });
    zCursor++;
  }

  // 1. Background image (raster URL) — async fetch + clip to radii.
  if (hasBgUrl) {
    const bgItem = {
      type: 'image',
      zIndex: zCursor,
      domOrder,
      options: {
        x, y, w, h,
        rotate: rotation,
        data: null,
        transparency: (1 - safeOpacity) * 100,
      },
    };
    items.push(bgItem);
    job = async () => {
      const processed = await getProcessedImage(
        urlMatch[1],
        widthPx,
        heightPx,
        radiiScalar,
        style.backgroundSize || 'cover',
        style.backgroundPosition || '50% 50%'
      );
      if (processed) bgItem.options.data = processed;
      else bgItem.skip = true;
    };
    zCursor++;
  } else if (hasGradient) {
    // 2. Gradient background → SVG.
    const borderForSvg = hasUniformBorder
      ? { color: borderInfo.options.color, width: parseFloat(style.borderWidth) || 0 }
      : null;
    const radiusArg = needsCustomShape ? radiiXY : radiiScalar.tl;
    const svgData = generateGradientSVG(widthPx, heightPx, bgImg, radiusArg, borderForSvg);
    if (svgData) {
      items.push({
        type: 'image',
        zIndex: zCursor,
        domOrder,
        options: {
          data: svgData,
          x, y, w, h,
          rotate: rotation,
          transparency: (1 - safeOpacity) * 100,
        },
      });
      zCursor++;
    }
  } else if (hasBg && needsCustomShape) {
    // 3. Solid fill with non-uniform / elliptical corner radii → SVG path.
    const shapeSvg = generateCustomShapeSVG(
      widthPx,
      heightPx,
      bg.hex,
      bg.opacity * safeOpacity,
      radiiXY
    );
    items.push({
      type: 'image',
      zIndex: zCursor,
      domOrder,
      options: { data: shapeSvg, x, y, w, h, rotate: rotation },
    });
    zCursor++;
  } else if (hasBg || hasUniformBorder || hasShadowVisual) {
    // 4. Native rect / roundRect / ellipse.
    let shapeType = pptx.ShapeType.rect;
    let rectRadius;
    const r = radiiScalar.tl;
    if (r > 0) {
      const isFullyRound = r >= minDim / 2;
      const isSquare = Math.abs(widthPx - heightPx) < 1;
      if (isFullyRound && isSquare) {
        shapeType = pptx.ShapeType.ellipse;
      } else {
        shapeType = pptx.ShapeType.roundRect;
        const cappedR = Math.min(r, minDim / 2);
        rectRadius = cappedR * PX_TO_INCH * layoutConfig.styleScale;
      }
    }
    const finalAlpha = safeOpacity * bg.opacity;
    const opts = {
      x, y, w, h,
      rotate: rotation,
      fill: hasBg
        ? { color: bg.hex, transparency: (1 - finalAlpha) * 100 }
        : { type: 'none' },
      line: hasUniformBorder ? borderInfo.options : null,
    };
    if (shadow) opts.shadow = shadow;
    if (typeof rectRadius === 'number') opts.rectRadius = rectRadius;
    items.push({
      type: 'shape',
      zIndex: zCursor,
      domOrder,
      shapeType,
      options: opts,
    });
    zCursor++;
  }

  // 4b. Inner multi-shadow halo above the shape (clipped to shape interior).
  if (shadowPlan.innerImage) {
    items.push({
      type: 'image',
      zIndex: zCursor,
      domOrder,
      options: {
        data: shadowPlan.innerImage.dataUrl,
        x, y, w, h,
        rotate: rotation,
      },
    });
    zCursor++;
  }

  // 5. Inline content text overlay (corner brand mark, "LINEA D'ACQUA"
  // badge, etc.). Rendered above any background fill so it reads on top.
  if (contentText) {
    const textOpts = getTextStyle(style, layoutConfig.styleScale);
    if (!isNaN(safeOpacity) && safeOpacity < 1) {
      // Approximate text opacity by transparency on the run — pptxgenjs
      // doesn't expose a per-run alpha, but transparency on the color
      // approximation is good enough for low-saturation badges.
      textOpts.transparency = (1 - safeOpacity) * 100;
    }
    let align = style.textAlign || 'left';
    if (align === 'start') align = 'left';
    if (align === 'end') align = 'right';
    const padding = getPadding(style, layoutConfig.styleScale);
    items.push({
      type: 'text',
      zIndex: zCursor,
      domOrder,
      textParts: [{ text: sanitizeText(contentText), options: textOpts }],
      options: {
        x, y, w, h,
        align,
        valign: 'middle',
        margin: paddingToPptxMargin(padding),
        wrap: false,
        autoFit: true,
        rotate: rotation,
      },
    });
  }

  if (items.length === 0) return null;
  return { items, job };
}

// 1/8pt = 1/(8*72) inch ≈ 0.001736 in. Snapping shape positions to this grid
// stops PowerPoint from rendering hairline gaps between adjacent shapes when
// floats land 0.001in apart. We deliberately do NOT round font sizes — the
// pipeline preserves fractional points (e.g. 11.3pt) as a feature.
const POSITION_SNAP_PER_INCH = 8 * 72; // eighths of a point per inch
function snapInch(v) {
  if (typeof v !== 'number' || !isFinite(v)) return v;
  return Math.round(v * POSITION_SNAP_PER_INCH) / POSITION_SNAP_PER_INCH;
}
function snapItemDimensions(item) {
  const o = item && item.options;
  if (!o) return;
  if (typeof o.x === 'number') o.x = snapInch(o.x);
  if (typeof o.y === 'number') o.y = snapInch(o.y);
  if (typeof o.w === 'number') o.w = snapInch(o.w);
  if (typeof o.h === 'number') o.h = snapInch(o.h);
  if (typeof o.rectRadius === 'number') o.rectRadius = snapInch(o.rectRadius);
}

/**
 * Optimized html2canvas wrapper
 * Includes fix for cropped icons by adjusting styles in the cloned document.
 */
async function elementToCanvasImage(node, widthPx, heightPx) {
  return new Promise((resolve) => {
    // 1. Assign a temp ID to locate the node inside the cloned document
    const originalId = node.id;
    const tempId = 'pptx-capture-' + Math.random().toString(36).substr(2, 9);
    node.id = tempId;

    const width = Math.max(Math.ceil(widthPx), 1);
    const height = Math.max(Math.ceil(heightPx), 1);
    const style = window.getComputedStyle(node);

    // Add padding to the clone to capture spilling content (like extensive font glyphs)
    const padding = 10;

    html2canvas(node, {
      backgroundColor: null,
      logging: false,
      scale: 3, // Higher scale for sharper icons
      useCORS: true, // critical for external fonts/images
      width: width + padding * 2, // Capture a larger area
      height: height + padding * 2,
      x: -padding, // Offset capture to include the padding
      y: -padding,
      onclone: (clonedDoc) => {
        const clonedNode = clonedDoc.getElementById(tempId);
        if (clonedNode) {
          // --- FIX: CLIP & FONT ISSUES ---
          // Apply styles DIRECTLY to elements to ensure html2canvas picks them up
          // This avoids issues where <style> tags in onclone are ignored or delayed

          // 1. Force FontAwesome Family on Icons
          const icons = clonedNode.querySelectorAll('.fa, .fas, .far, .fab');
          icons.forEach((icon) => {
            icon.style.setProperty('font-family', 'FontAwesome', 'important');
          });

          // 2. Fix Image Display
          const images = clonedNode.querySelectorAll('img');
          images.forEach((img) => {
            img.style.setProperty('display', 'inline-block', 'important');
          });

          // 3. Force overflow visible on the container so glyphs bleeding out aren't cut
          clonedNode.style.overflow = 'visible';

          // 4. Adjust alignment for Icons to prevent baseline clipping
          // (Applies to <i>, <span>, or standard icon classes)
          const tag = clonedNode.tagName;
          if (tag === 'I' || tag === 'SPAN' || clonedNode.className.includes('fa-')) {
            // Flex center helps align the glyph exactly in the middle of the box
            // preventing top/bottom cropping due to line-height mismatches.
            clonedNode.style.display = 'inline-flex';
            clonedNode.style.justifyContent = 'center';
            clonedNode.style.alignItems = 'center';
            clonedNode.style.setProperty('font-family', 'FontAwesome', 'important'); // Ensure root icon gets it too

            // Remove margins that might offset the capture
            clonedNode.style.margin = '0';

            // Ensure the font fits
            clonedNode.style.lineHeight = '1';
            clonedNode.style.verticalAlign = 'middle';
          }
        }
      },
    })
      .then((canvas) => {
        // Restore the original ID
        if (originalId) node.id = originalId;
        else node.removeAttribute('id');

        const destCanvas = document.createElement('canvas');
        destCanvas.width = width;
        destCanvas.height = height;
        const ctx = destCanvas.getContext('2d');

        // Draw captured canvas (which is padded) back to the original size
        // We need to draw the CENTER of the source canvas to the destination
        // The source canvas is (width + 2*padding) * scale
        // We want to draw the crop starting at padding*scale
        const scale = 3;
        const sX = padding * scale;
        const sY = padding * scale;
        const sW = width * scale;
        const sH = height * scale;

        ctx.drawImage(canvas, sX, sY, sW, sH, 0, 0, width, height);

        // --- Border Radius Clipping (Existing Logic) ---
        let tl = parseFloat(style.borderTopLeftRadius) || 0;
        let tr = parseFloat(style.borderTopRightRadius) || 0;
        let br = parseFloat(style.borderBottomRightRadius) || 0;
        let bl = parseFloat(style.borderBottomLeftRadius) || 0;

        const f = Math.min(
          width / (tl + tr) || Infinity,
          height / (tr + br) || Infinity,
          width / (br + bl) || Infinity,
          height / (bl + tl) || Infinity
        );

        if (f < 1) {
          tl *= f;
          tr *= f;
          br *= f;
          bl *= f;
        }

        if (tl + tr + br + bl > 0) {
          ctx.globalCompositeOperation = 'destination-in';
          ctx.beginPath();
          ctx.moveTo(tl, 0);
          ctx.lineTo(width - tr, 0);
          ctx.arcTo(width, 0, width, tr, tr);
          ctx.lineTo(width, height - br);
          ctx.arcTo(width, height, width - br, height, br);
          ctx.lineTo(bl, height);
          ctx.arcTo(0, height, 0, height - bl, bl);
          ctx.lineTo(0, tl);
          ctx.arcTo(0, 0, tl, 0, tl);
          ctx.closePath();
          ctx.fill();
        }

        resolve(destCanvas.toDataURL('image/png'));
      })
      .catch((e) => {
        if (originalId) node.id = originalId;
        else node.removeAttribute('id');
        console.warn('Canvas capture failed for node', node, e);
        resolve(null);
      });
  });
}

/**
 * Helper to identify elements that should be rendered as icons (Images).
 * Detects Custom Elements AND generic tags (<i>, <span>) with icon classes/pseudo-elements.
 */
function isIconElement(node) {
  // 1. Custom Elements (hyphenated tags) or Explicit Library Tags
  const tag = node.tagName.toUpperCase();
  if (
    tag.includes('-') ||
    [
      'MATERIAL-ICON',
      'ICONIFY-ICON',
      'REMIX-ICON',
      'ION-ICON',
      'EVA-ICON',
      'BOX-ICON',
      'FA-ICON',
    ].includes(tag)
  ) {
    return true;
  }

  // 2. Class-based Icons (FontAwesome, Bootstrap, Material symbols) on <i> or <span>
  if (tag === 'I' || tag === 'SPAN') {
    const cls = node.getAttribute('class') || '';
    if (
      typeof cls === 'string' &&
      (cls.includes('fa-') ||
        cls.includes('fas') ||
        cls.includes('far') ||
        cls.includes('fab') ||
        cls.includes('bi-') ||
        cls.includes('material-icons') ||
        cls.includes('icon'))
    ) {
      // Double-check: Must have pseudo-element content to be a CSS icon
      const before = window.getComputedStyle(node, '::before').content;
      const after = window.getComputedStyle(node, '::after').content;
      const hasContent = (c) => c && c !== 'none' && c !== 'normal' && c !== '""';

      if (hasContent(before) || hasContent(after)) return true;
    }
  }

  return false;
}

/**
 * Replaces createRenderItem.
 * Returns { items: [], job: () => Promise, stopRecursion: boolean }
 */
function prepareRenderItem(
  node,
  config,
  domOrder,
  pptx,
  effectiveZIndex,
  computedStyle,
  globalOptions = {}
) {
  // 1. Text Node Handling
  if (node.nodeType === 3) {
    const textContent = node.nodeValue.trim();
    if (!textContent) return null;

    const parent = node.parentElement;
    if (!parent) return null;

    if (isTextContainer(parent)) return null; // Parent handles it

    const range = document.createRange();
    range.selectNode(node);
    const rect = range.getBoundingClientRect();
    range.detach();

    const style = window.getComputedStyle(parent);
    const widthPx = rect.width;
    const heightPx = rect.height;
    const unrotatedW = widthPx * PX_TO_INCH * config.scale;
    const unrotatedH = heightPx * PX_TO_INCH * config.scale;

    const x = config.offX + (rect.left - config.rootX) * PX_TO_INCH * config.scale;
    const y = config.offY + (rect.top - config.rootY) * PX_TO_INCH * config.scale;

    return {
      items: [
        {
          type: 'text',
          zIndex: effectiveZIndex,
          domOrder,
          textParts: [
            {
              text: sanitizeText(textContent),
              options: getTextStyle(style, config.styleScale),
            },
          ],
          options: { x, y, w: unrotatedW, h: unrotatedH, margin: 0, autoFit: true },
        },
      ],
      stopRecursion: false,
    };
  }

  if (node.nodeType !== 1) return null;
  const style = computedStyle; // Use pre-computed style

  const rect = node.getBoundingClientRect();
  if (rect.width < 0.5 || rect.height < 0.5) return null;

  const zIndex = effectiveZIndex;
  const rotation = getRotation(style.transform);
  if (style.transform && style.transform !== 'none' && hasSkew(style.transform)) {
    _warnSkewOnce(node, style.transform);
  }
  const writingModeVert = getWritingModeVert(style.writingMode, style.textOrientation);
  const elementOpacity = parseFloat(style.opacity);
  const safeOpacity = isNaN(elementOpacity) ? 1 : elementOpacity;

  // Size from `offsetWidth/Height` (logical pre-transform px). For unrotated
  // elements this matches `rect.width`, but for elements with their own
  // `transform: rotate(...)` rect.* returns the axis-aligned bounding box of
  // the rotated shape — too large for the pre-rotation rectangle we need
  // before applying `rotate` separately. Convert with `styleScale` (which
  // folds in ancestor + descendant scale transforms), not the raw fit-scale,
  // so a `transform: scale(0.6)` ancestor still produces the right visible
  // width. Position uses rect.* with `scale` (already post-transform).
  const widthPx = node.offsetWidth || rect.width;
  const heightPx = node.offsetHeight || rect.height;
  const unrotatedW = widthPx * PX_TO_INCH * config.styleScale;
  const unrotatedH = heightPx * PX_TO_INCH * config.styleScale;
  const centerX = rect.left + rect.width / 2;
  const centerY = rect.top + rect.height / 2;

  let x = config.offX + (centerX - config.rootX) * PX_TO_INCH * config.scale - unrotatedW / 2;
  let y = config.offY + (centerY - config.rootY) * PX_TO_INCH * config.scale - unrotatedH / 2;
  let w = unrotatedW;
  let h = unrotatedH;

  const items = [];

  if (node.tagName === 'TABLE') {
    const tableData = extractTableData(node, config.scale, config.styleScale);
    const tableItems = [
      {
        type: 'table',
        zIndex: effectiveZIndex,
        domOrder,
        tableData: tableData,
        options: { x, y, w: unrotatedW, h: unrotatedH },
      },
    ];

    // 1. Check for Background / Shadow / Radius on the table itself
    const shadowStr = style.boxShadow;
    const hasShadow = shadowStr && shadowStr !== 'none';
    const borderRadius = parseFloat(style.borderRadius) || 0;
    const bgColor = parseColor(style.backgroundColor);
    const hasBg = bgColor.hex && bgColor.opacity > 0;

    if (hasShadow || borderRadius > 0 || hasBg) {
      const transparency = (1 - bgColor.opacity) * 100;
      const tableRadii = getCornerRadii(style, widthPx, heightPx);
      const shadowPlan = hasShadow
        ? planShadows(shadowStr, widthPx, heightPx, tableRadii, config.styleScale)
        : { primaryShadow: null, outerImage: null, innerImage: null };
      const shadow = shadowPlan.primaryShadow;
      let shapeType = pptx.ShapeType.rect;
      let rectRadius = 0;

      if (borderRadius > 0) {
        shapeType = pptx.ShapeType.roundRect;
        let cappedRadiusPx = Math.min(borderRadius, Math.min(widthPx, heightPx) / 2);
        rectRadius = cappedRadiusPx * PX_TO_INCH * config.styleScale;
      }

      // Outer multi-shadow halo behind the backing shape.
      if (shadowPlan.outerImage) {
        const padIn = shadowPlan.outerImage.paddingPx * PX_TO_INCH * config.styleScale;
        tableItems.unshift({
          type: 'image',
          zIndex: effectiveZIndex,
          domOrder,
          options: {
            data: shadowPlan.outerImage.dataUrl,
            x: x - padIn,
            y: y - padIn,
            w: unrotatedW + padIn * 2,
            h: unrotatedH + padIn * 2,
          },
        });
      }

      // Add a backing shape item before the table
      tableItems.unshift({
        type: 'shape',
        zIndex: effectiveZIndex,
        domOrder, // Same domOrder ensures it renders before the table (queue order)
        shapeType,
        options: {
          x,
          y,
          w: unrotatedW,
          h: unrotatedH,
          fill: hasBg ? { color: bgColor.hex, transparency } : { type: 'none' },
          shadow,
          rectRadius,
        },
      });

      // Inner multi-shadow halo above the backing shape (clipped to shape).
      if (shadowPlan.innerImage) {
        // Insert just after the backing shape, before the table cells, so it
        // overlays the fill but stays under cell text.
        tableItems.splice(1, 0, {
          type: 'image',
          zIndex: effectiveZIndex,
          domOrder,
          options: {
            data: shadowPlan.innerImage.dataUrl,
            x,
            y,
            w: unrotatedW,
            h: unrotatedH,
          },
        });
      }
    }

    return {
      items: tableItems,
      stopRecursion: true,
    };
  }

  if ((node.tagName === 'UL' || node.tagName === 'OL') && !isComplexHierarchy(node)) {
    const listItems = [];
    const liChildren = Array.from(node.children).filter((c) => {
      if (c.tagName !== 'LI') return false;
      // Skip LIs the browser is rendering as fully invisible — reveal.js
      // fragments live in the DOM with opacity:0 / visibility:hidden until
      // their click index is reached.
      const cs = window.getComputedStyle(c);
      return !(cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0');
    });

    liChildren.forEach((child, index) => {
      const liStyle = window.getComputedStyle(child);
      const liRect = child.getBoundingClientRect();
      const parentRect = node.getBoundingClientRect(); // node is UL/OL

      // 1. Determine Bullet Config
      let bullet = { type: 'bullet' };
      const listStyleType = liStyle.listStyleType || 'disc';

      if (node.tagName === 'OL' || listStyleType === 'decimal') {
        bullet = { type: 'number' };
      } else if (listStyleType === 'none') {
        bullet = false;
      } else {
        let code = '2022'; // disc
        if (listStyleType === 'circle') code = '25CB';
        if (listStyleType === 'square') code = '25A0';

        // --- CHANGE: Color & Size Logic (Option > ::marker > CSS color) ---
        let finalHex = '000000';
        let markerFontSize = null;

        // A. Check Global Option override
        if (globalOptions?.listConfig?.color) {
          finalHex = parseColor(globalOptions.listConfig.color).hex || '000000';
        }
        // B. Check ::marker pseudo element (supported in modern browsers)
        else {
          const markerStyle = window.getComputedStyle(child, '::marker');
          const markerColor = parseColor(markerStyle.color);
          if (markerColor.hex) {
            finalHex = markerColor.hex;
          } else {
            // C. Fallback to LI text color
            const colorObj = parseColor(liStyle.color);
            if (colorObj.hex) finalHex = colorObj.hex;
          }

          // Check ::marker font-size
          const markerFs = parseFloat(markerStyle.fontSize);
          if (!isNaN(markerFs) && markerFs > 0) {
            // Convert px->pt for PPTX (logical px → uses styleScale)
            markerFontSize = markerFs * 0.75 * config.styleScale;
          }
        }

        bullet = { code, color: finalHex };
        if (markerFontSize) {
          bullet.fontSize = markerFontSize;
        }
      }

      // 2. Calculate Dynamic Indent (Respects padding-left)
      // Visual Indent = Distance from UL left edge to LI Content left edge.
      // PptxGenJS 'indent' = Space between bullet and text?
      // Actually PptxGenJS 'indent' allows setting the hanging indent.
      // We calculate the TOTAL visual offset from the parent container.
      // 1 px = 0.75 pt (approx, standard DTP).
      // We must scale it by config.scale.
      const visualIndentPx = liRect.left - parentRect.left;
      /*
         Standard indent in PPT is ~27pt.
         If visualIndentPx is small (e.g. 10px padding), we want small indent.
         If visualIndentPx is large (40px padding), we want large indent.
         We treat 'indent' as the value to pass to PptxGenJS.
      */
      const computedIndentPt = visualIndentPx * 0.75 * config.scale;

      if (bullet && computedIndentPt > 0) {
        bullet.indent = computedIndentPt;
        // Also support custom margin between bullet and text if provided in listConfig?
        // For now, computedIndentPt covers the visual placement.
      }

      // 3. Extract Text Parts (font-size etc. are read from computed style → styleScale)
      const parts = collectTextParts(child, liStyle, config.styleScale);

      if (parts.length > 0) {
        parts.forEach((p) => {
          if (!p.options) p.options = {};
        });

        // A. Apply Bullet
        // Workaround: pptxgenjs bullets inherit the style of the text run they are attached to.
        // To support ::marker styles (color, size) that differ from the text, we create
        // a "dummy" text run at the start of the list item that carries the bullet configuration.
        if (bullet) {
          const firstPartInfo = parts[0].options;

          // Create a dummy run. We use a Zero Width Space to ensure it's rendered but invisible.
          // This "run" will hold the bullet and its specific color/size.
          const bulletRun = {
            text: '\u200B',
            options: {
              ...firstPartInfo, // Inherit base props (fontFace, etc.)
              color: bullet.color || firstPartInfo.color,
              fontSize: bullet.fontSize || firstPartInfo.fontSize,
              bullet: bullet,
            },
          };

          // Don't duplicate transparent or empty color from firstPart if bullet has one
          if (bullet.color) bulletRun.options.color = bullet.color;
          if (bullet.fontSize) bulletRun.options.fontSize = bullet.fontSize;

          // Prepend
          parts.unshift(bulletRun);
        }

        // B. Apply Spacing
        let ptBefore = 0;
        let ptAfter = 0;

        // A. Check Global Options (Expected in Points)
        if (globalOptions.listConfig?.spacing) {
          if (typeof globalOptions.listConfig.spacing.before === 'number') {
            ptBefore = globalOptions.listConfig.spacing.before;
          }
          if (typeof globalOptions.listConfig.spacing.after === 'number') {
            ptAfter = globalOptions.listConfig.spacing.after;
          }
        }
        // B. Fallback to CSS Margins (Convert px -> pt — logical, styleScale)
        else {
          const mt = parseFloat(liStyle.marginTop) || 0;
          const mb = parseFloat(liStyle.marginBottom) || 0;
          if (mt > 0) ptBefore = mt * 0.75 * config.styleScale;
          if (mb > 0) ptAfter = mb * 0.75 * config.styleScale;
        }

        if (ptBefore > 0) parts[0].options.paraSpaceBefore = ptBefore;
        if (ptAfter > 0) parts[0].options.paraSpaceAfter = ptAfter;

        if (index < liChildren.length - 1) {
          parts[parts.length - 1].options.breakLine = true;
        }

        listItems.push(...parts);
      }
    });

    if (listItems.length > 0) {
      // Add background if exists
      const bgColorObj = parseColor(style.backgroundColor);
      if (bgColorObj.hex && bgColorObj.opacity > 0) {
        items.push({
          type: 'shape',
          zIndex,
          domOrder,
          shapeType: 'rect',
          options: { x, y, w, h, fill: { color: bgColorObj.hex } },
        });
      }

      items.push({
        type: 'text',
        zIndex: zIndex + 1,
        domOrder,
        textParts: listItems,
        options: {
          x,
          y,
          w,
          h,
          align: 'left',
          valign: 'top',
          margin: 0,
          autoFit: true,
          wrap: true,
          vert: writingModeVert,
        },
      });

      return { items, stopRecursion: true };
    }
  }

  if (node.tagName === 'CANVAS') {
    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };

    const job = async () => {
      try {
        // Direct data extraction from the canvas element
        // This preserves the exact current state of the chart
        const dataUrl = node.toDataURL('image/png');

        // Basic validation
        if (dataUrl && dataUrl.length > 10) {
          item.options.data = dataUrl;
        } else {
          item.skip = true;
        }
      } catch (e) {
        // Tainted canvas (CORS issues) will throw here
        console.warn('Failed to capture canvas content:', e);
        item.skip = true;
      }
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- ASYNC JOB: SVG Tags ---
  if (node.nodeName.toUpperCase() === 'SVG') {
    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { data: null, x, y, w, h, rotate: rotation },
    };

    const job = async () => {
      // Use svgToSvg for vector output (Convert to Shape in PowerPoint)
      // Use svgToPng for rasterized output (pixel perfect)
      const converter = globalOptions.svgAsVector ? svgToSvg : svgToPng;
      const processed = await converter(node);
      if (processed) item.options.data = processed;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- ASYNC JOB: IMG Tags ---
  if (node.tagName === 'IMG') {
    let radii = {
      tl: parseFloat(style.borderTopLeftRadius) || 0,
      tr: parseFloat(style.borderTopRightRadius) || 0,
      br: parseFloat(style.borderBottomRightRadius) || 0,
      bl: parseFloat(style.borderBottomLeftRadius) || 0,
    };

    const hasAnyRadius = radii.tl > 0 || radii.tr > 0 || radii.br > 0 || radii.bl > 0;
    if (!hasAnyRadius) {
      const parent = node.parentElement;
      const parentStyle = window.getComputedStyle(parent);
      if (parentStyle.overflow !== 'visible') {
        const pRadii = {
          tl: parseFloat(parentStyle.borderTopLeftRadius) || 0,
          tr: parseFloat(parentStyle.borderTopRightRadius) || 0,
          br: parseFloat(parentStyle.borderBottomRightRadius) || 0,
          bl: parseFloat(parentStyle.borderBottomLeftRadius) || 0,
        };
        const pRect = parent.getBoundingClientRect();
        if (Math.abs(pRect.width - rect.width) < 5 && Math.abs(pRect.height - rect.height) < 5) {
          radii = pRadii;
        }
      }
    }

    const objectFit = style.objectFit || 'fill'; // default CSS behavior is fill
    const objectPosition = style.objectPosition || '50% 50%';

    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };

    const job = async () => {
      const processed = await getProcessedImage(
        node.src,
        widthPx,
        heightPx,
        radii,
        objectFit,
        objectPosition
      );
      if (processed) item.options.data = processed;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- ASYNC JOB: Icons and Other Elements ---
  if (isIconElement(node)) {
    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };
    const job = async () => {
      const pngData = await elementToCanvasImage(node, widthPx, heightPx);
      if (pngData) item.options.data = pngData;
      else item.skip = true;
    };
    return { items: [item], job, stopRecursion: true };
  }

  // Radii logic. Read once into per-corner elliptical form; derive scalar
  // form for code paths (PPTX `roundRect`, shadow halos) that don't render
  // elliptical arcs. `needsCustomShape` is true when corners are either
  // elliptical (rx≠ry) or per-corner divergent — both require an SVG path
  // because PPTX `prstGeom: roundRect` only supports a single uniform radius.
  const borderRadiusValue = parseFloat(style.borderRadius) || 0;
  const radiiXY = getCornerRadiiXY(style, widthPx, heightPx);
  const radiiScalar = radiiXYToScalar(radiiXY);
  const hasPartialBorderRadius = isElliptical(radiiXY) || isPerCorner(radiiXY);

  // --- PRIORITY SVG: Solid Fill with Partial Border Radius (Vector Cone/Tab) ---
  // Fix for "missing cone": Prioritize SVG vector generation over Raster Canvas for simple shapes with partial radii.
  // This avoids html2canvas failures on empty divs.
  const tempBg = parseColor(style.backgroundColor);
  const isTxt = isTextContainer(node);

  // BUG FIX: Don't treat as a vector shape if it has content (like text or children).
  // This prevents containers like ".glass-box" from being treated as empty shapes and stopping recursion.
  const hasContent = node.textContent.trim().length > 0 || node.children.length > 0;

  if (hasPartialBorderRadius && tempBg.hex && !isTxt && !hasContent) {
    const shapeSvg = generateCustomShapeSVG(
      widthPx,
      heightPx,
      tempBg.hex,
      tempBg.opacity,
      radiiXY
    );

    return {
      items: [
        {
          type: 'image',
          zIndex,
          domOrder,
          options: { data: shapeSvg, x, y, w, h, rotate: rotation },
        },
      ],
      stopRecursion: true, // Treat as leaf
    };
  }

  // --- ASYNC JOB: Clipped Divs via Canvas ---
  // Only capture as image if it's an empty leaf.
  // Rasterizing containers (like .glass-box) kills editability of children.
  if (hasPartialBorderRadius && isClippedByParent(node) && !hasContent) {
    const marginLeft = parseFloat(style.marginLeft) || 0;
    const marginTop = parseFloat(style.marginTop) || 0;
    x += marginLeft * PX_TO_INCH * config.styleScale;
    y += marginTop * PX_TO_INCH * config.styleScale;

    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };

    const job = async () => {
      const canvasImageData = await elementToCanvasImage(node, widthPx, heightPx);
      if (canvasImageData) item.options.data = canvasImageData;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- SYNC: Standard CSS Extraction ---
  const bgColorObj = parseColor(style.backgroundColor);
  const bgClip = style.webkitBackgroundClip || style.backgroundClip;
  const isBgClipText = bgClip === 'text';
  const bgImgStr = style.backgroundImage;
  const hasGradient = !isBgClipText && bgImgStr && bgImgStr.includes('linear-gradient');
  const urlMatch = !isBgClipText && !hasGradient && bgImgStr ? bgImgStr.match(/url\(['"]?(.*?)['"]?\)/) : null;
  const hasBgImgUrl = !!urlMatch;

  const borderColorObj = parseColor(style.borderColor);
  const borderWidth = parseFloat(style.borderWidth);
  const hasBorder = borderWidth > 0 && borderColorObj.hex;

  const borderInfo = getBorderInfo(style, config.styleScale);
  const hasUniformBorder = borderInfo.type === 'uniform';
  const hasCompositeBorder = borderInfo.type === 'composite';

  const shadowStr = style.boxShadow;
  const hasShadow = shadowStr && shadowStr !== 'none';
  const shadowPlan = hasShadow
    ? planShadows(shadowStr, widthPx, heightPx, radiiScalar, config.styleScale)
    : { primaryShadow: null, outerImage: null, innerImage: null };
  const softEdge = getSoftEdges(style.filter, config.styleScale);

  let isImageWrapper = false;
  const imgChild = Array.from(node.children).find((c) => c.tagName === 'IMG');
  if (imgChild) {
    const childW = imgChild.offsetWidth || imgChild.getBoundingClientRect().width;
    const childH = imgChild.offsetHeight || imgChild.getBoundingClientRect().height;
    if (childW >= widthPx - 2 && childH >= heightPx - 2) isImageWrapper = true;
  }

  let textPayload = null;
  const isText = isTextContainer(node);

  if (isText) {
    const textParts = [];
    let trimNextLeading = false;

    node.childNodes.forEach((child, index) => {
      // Handle <br> tags
      if (child.tagName === 'BR') {
        // 1. Trim trailing space from the *previous* text part to prevent double wrapping
        if (textParts.length > 0) {
          const lastPart = textParts[textParts.length - 1];
          if (lastPart.text && typeof lastPart.text === 'string') {
            lastPart.text = lastPart.text.trimEnd();
          }
        }

        textParts.push({ text: '', options: { breakLine: true } });

        // 2. Signal to trim leading space from the *next* text part
        trimNextLeading = true;
        return;
      }

      let rawTextVal = child.nodeType === 3 ? child.nodeValue : child.textContent;
      let nodeStyle = child.nodeType === 1 ? window.getComputedStyle(child) : style;
      // Skip inline children the browser renders fully invisible (reveal.js
      // fragment chips inside a text-container parent).
      if (
        child.nodeType === 1 &&
        (nodeStyle.display === 'none' ||
          nodeStyle.visibility === 'hidden' ||
          nodeStyle.opacity === '0')
      ) {
        return;
      }

      // Block-level inline children (e.g. `<span style="display:block">`) need
      // their own line. The browser breaks before and after them; PPTX rich
      // text doesn't honor display so we synthesize the breaks here.
      const isBlockChild =
        child.nodeType === 1 &&
        (nodeStyle.display === 'block' ||
          nodeStyle.display === 'list-item' ||
          nodeStyle.display === 'flex' ||
          nodeStyle.display === 'grid' ||
          nodeStyle.display === 'flow-root' ||
          nodeStyle.display.startsWith('table'));
      if (isBlockChild && textParts.length > 0) {
        const last = textParts[textParts.length - 1];
        if (!last.options?.breakLine) {
          if (last.text && typeof last.text === 'string') {
            last.text = last.text.trimEnd();
          }
          textParts.push({ text: '', options: { breakLine: true } });
          trimNextLeading = true;
        }
      }

      const wsProcessed = processWhitespace(rawTextVal, nodeStyle.whiteSpace);
      // processWhitespace returns either a string (collapsed modes) or an
      // array of {text}/{breakLine} segments for `pre*` modes that contain
      // newlines. Normalize to an array for the rest of the pipeline.
      const segments = Array.isArray(wsProcessed)
        ? wsProcessed
        : [{ text: wsProcessed }];

      const textOptsBase = getTextStyle(nodeStyle, config.styleScale || config.scale);
      // BUG FIX: Numbers 1 and 2 having background.
      // If this is a naked Text Node (nodeType 3), it inherits style from the parent container.
      // The parent container's background is already rendered as the Shape Fill.
      // We must NOT render it again as a Text Highlight, otherwise it looks like a solid marker on top of the shape.
      if (child.nodeType === 3 && textOptsBase.highlight) {
        delete textOptsBase.highlight;
      }

      const isCollapsed = !Array.isArray(wsProcessed);
      const isLastSegment = (segIdx) => segIdx === segments.length - 1;

      segments.forEach((seg, segIdx) => {
        if (seg.breakLine) {
          textParts.push({ text: '', options: { breakLine: true } });
          trimNextLeading = true;
          return;
        }
        let textVal = seg.text || '';

        // Only trim collapsing-mode runs. In pre/pre-wrap/pre-line/break-spaces
        // we must preserve leading/trailing spaces.
        if (isCollapsed) {
          if (index === 0 && segIdx === 0) textVal = textVal.trimStart();
          if (trimNextLeading) {
            textVal = textVal.trimStart();
            trimNextLeading = false;
          }
          if (
            index === node.childNodes.length - 1 &&
            isLastSegment(segIdx)
          ) {
            textVal = textVal.trimEnd();
          }
        } else {
          // pre-mode: still consume the trim-leading flag from a preceding
          // explicit <br>, but otherwise leave whitespace alone.
          if (trimNextLeading) {
            textVal = textVal.replace(/^[ \t]+/, '');
            trimNextLeading = false;
          }
        }

        textVal = applyTextTransform(textVal, nodeStyle.textTransform);

        if (textVal.length > 0) {
          textParts.push({
            text: sanitizeText(textVal),
            options: { ...textOptsBase },
          });
        }
      });

      if (isBlockChild) {
        textParts.push({ text: '', options: { breakLine: true } });
        trimNextLeading = true;
      }
    });

    // Drop trailing empty breakLine sentinels — they'd render as a blank line
    // at the end of the shape (especially when the last child was block).
    while (
      textParts.length > 0 &&
      textParts[textParts.length - 1].options?.breakLine &&
      textParts[textParts.length - 1].text === ''
    ) {
      textParts.pop();
    }

    if (textParts.length > 0) {
      const { align, valign, intentionalSize, lineHeightVcenter } =
        resolveContainerAlignment(style, heightPx);

      let padding = getPadding(style, config.styleScale);
      if (align === 'center' && valign === 'middle') padding = [0, 0, 0, 0];

      // The line-height vcenter trick uses line-height as the box-filler.
      // Keeping it as a paragraph lineSpacing would stretch the line-box
      // and push the glyph baseline to the bottom of the shape — drop it
      // so PPTX vcenter measures the actual glyph box.
      if (lineHeightVcenter) {
        for (const part of textParts) {
          if (part.options) delete part.options.lineSpacing;
        }
      }

      textPayload = {
        text: textParts,
        align,
        valign,
        margin: paddingToPptxMargin(padding),
        intentionalSize,
      };
    }
  }

  let bgJob = null;

  if (hasBgImgUrl || hasGradient || (softEdge && bgColorObj.hex && !isImageWrapper)) {
    if (hasBgImgUrl) {
      const bgUrl = urlMatch[1];
      const radii = {
        tl: parseFloat(style.borderTopLeftRadius) || 0,
        tr: parseFloat(style.borderTopRightRadius) || 0,
        br: parseFloat(style.borderBottomRightRadius) || 0,
        bl: parseFloat(style.borderBottomLeftRadius) || 0,
      };
      
      const bgItem = {
        type: 'image',
        zIndex,
        domOrder,
        options: { x, y, w, h, rotate: rotation, data: null },
      };
      items.push(bgItem);
      
      bgJob = async () => {
        const processed = await getProcessedImage(
          bgUrl,
          widthPx,
          heightPx,
          radii,
          style.backgroundSize || 'cover',
          style.backgroundPosition || '50% 50%'
        );
        if (processed) bgItem.options.data = processed;
        else bgItem.skip = true;
      };
    } else {
      let bgData = null;
      let padIn = 0;
      if (softEdge) {
        const svgInfo = generateBlurredSVG(
          widthPx,
          heightPx,
          bgColorObj.hex,
          borderRadiusValue,
          softEdge
        );
        bgData = svgInfo.data;
        padIn = svgInfo.padding * PX_TO_INCH * config.scale;
      } else {
        bgData = generateGradientSVG(
          widthPx,
          heightPx,
          style.backgroundImage,
          hasPartialBorderRadius ? radiiXY : borderRadiusValue,
          hasBorder ? { color: borderColorObj.hex, width: borderWidth } : null
        );
      }

      if (bgData) {
        items.push({
          type: 'image',
          zIndex,
          domOrder,
          options: {
            data: bgData,
            x: x - padIn,
            y: y - padIn,
            w: w + padIn * 2,
            h: h + padIn * 2,
            rotate: rotation,
          },
        });
      }
    }

    if (textPayload) {
      textPayload.text[0].options.fontSize =
        Number(textPayload.text[0]?.options?.fontSize?.toFixed(1)) || 12;
      items.push({
        type: 'text',
        zIndex: zIndex + 1,
        domOrder,
        textParts: textPayload.text,
        options: {
          x,
          y,
          w,
          h,
          align: textPayload.align,
          valign: textPayload.valign,
          rotate: rotation,
          margin: textPayload.margin,
          wrap: true,
          // autoFit:true emits <a:spAutoFit/> ("resize shape to fit text"),
          // which collapses a flex/grid box back to text size. When the
          // alignment readout flagged the box as intentionally sized,
          // keep the declared dimensions instead.
          autoFit: !textPayload.intentionalSize,
          vert: writingModeVert,
        },
      });
    }
    if (hasCompositeBorder) {
      const borderItems = createCompositeBorderItems(
        borderInfo.sides,
        x,
        y,
        w,
        h,
        config.styleScale,
        zIndex,
        domOrder
      );
      items.push(...borderItems);
    }
  } else if (
    (bgColorObj.hex && !isImageWrapper) ||
    hasUniformBorder ||
    hasCompositeBorder ||
    hasShadow ||
    textPayload
  ) {
    const finalAlpha = safeOpacity * bgColorObj.opacity;
    const transparency = (1 - finalAlpha) * 100;
    const useSolidFill = bgColorObj.hex && !isImageWrapper;

    // Outer multi-shadow halo behind the shape.
    if (shadowPlan.outerImage) {
      const padIn = shadowPlan.outerImage.paddingPx * PX_TO_INCH * config.styleScale;
      items.push({
        type: 'image',
        zIndex,
        domOrder,
        options: {
          data: shadowPlan.outerImage.dataUrl,
          x: x - padIn,
          y: y - padIn,
          w: w + padIn * 2,
          h: h + padIn * 2,
          rotate: rotation,
        },
      });
    }

    if (hasPartialBorderRadius && useSolidFill && !textPayload) {
      const shapeSvg = generateCustomShapeSVG(
        widthPx,
        heightPx,
        bgColorObj.hex,
        bgColorObj.opacity,
        radiiXY
      );

      items.push({
        type: 'image',
        zIndex,
        domOrder,
        options: { data: shapeSvg, x, y, w, h, rotate: rotation },
      });
    } else {
      const shapeOpts = {
        x,
        y,
        w,
        h,
        rotate: rotation,
        fill: useSolidFill
          ? { color: bgColorObj.hex, transparency: transparency }
          : { type: 'none' },
        line: hasUniformBorder ? borderInfo.options : null,
      };

      if (shadowPlan.primaryShadow) shapeOpts.shadow = shadowPlan.primaryShadow;

      // 1. Calculate dimensions first
      const minDimension = Math.min(widthPx, heightPx);

      let rawRadius = parseFloat(style.borderRadius) || 0;
      const isPercentage = style.borderRadius && style.borderRadius.toString().includes('%');

      // 2. Normalize radius to pixels
      let radiusPx = rawRadius;
      if (isPercentage) {
        radiusPx = (rawRadius / 100) * minDimension;
      }

      let shapeType = pptx.ShapeType.rect;

      // 3. Determine Shape Logic
      const isSquare = Math.abs(widthPx - heightPx) < 1;
      const isFullyRound = radiusPx >= minDimension / 2;

      // CASE A: It is an Ellipse if:
      // 1. It is explicitly "50%" (standard CSS way to make ovals/circles)
      // 2. OR it is a perfect square and fully rounded (a circle)
      if (isFullyRound && (isPercentage || isSquare)) {
        shapeType = pptx.ShapeType.ellipse;
      }
      // CASE B: It is a Rounded Rectangle (including "Pill" shapes)
      else if (radiusPx > 0) {
        shapeType = pptx.ShapeType.roundRect;
        let cappedRadiusPx = Math.min(radiusPx, minDimension / 2);
        shapeOpts.rectRadius = cappedRadiusPx * PX_TO_INCH * config.styleScale;
      }

      if (textPayload) {
        textPayload.text[0].options.fontSize =
          Number(textPayload.text[0]?.options?.fontSize?.toFixed(1)) || 12;
        const textOptions = {
          shape: shapeType,
          ...shapeOpts,
          rotate: rotation,
          align: textPayload.align,
          valign: textPayload.valign,
          margin: textPayload.margin,
          wrap: true,
          // See note in the bg-image branch: spAutoFit collapses the
          // box. Disable when the alignment branch flagged the box as
          // intentionally larger than its text.
          autoFit: !textPayload.intentionalSize,
          vert: writingModeVert,
        };
        items.push({
          type: 'text',
          zIndex,
          domOrder,
          textParts: textPayload.text,
          options: textOptions,
        });
      } else if (!hasPartialBorderRadius) {
        items.push({
          type: 'shape',
          zIndex,
          domOrder,
          shapeType,
          options: shapeOpts,
        });
      }
    }

    // Inner multi-shadow halo above the shape (clipped to shape interior).
    if (shadowPlan.innerImage) {
      items.push({
        type: 'image',
        zIndex,
        domOrder,
        options: {
          data: shadowPlan.innerImage.dataUrl,
          x, y, w, h,
          rotate: rotation,
        },
      });
    }

    // CSS `outline` — a non-space-taking ring outside the border edge,
    // offset by `outline-offset`. Emit as a separate non-filled shape.
    {
      const outlineWidthPx = parseFloat(style.outlineWidth) || 0;
      const outlineStyle = style.outlineStyle;
      if (
        outlineWidthPx > 0 &&
        outlineStyle &&
        outlineStyle !== 'none' &&
        outlineStyle !== 'hidden'
      ) {
        const outlineColor = parseColor(style.outlineColor);
        if (outlineColor.hex && outlineColor.opacity > 0) {
          const offsetPx = parseFloat(style.outlineOffset) || 0;
          // Path centerline sits at offset + width/2 outside the border edge.
          const inflatePx = offsetPx + outlineWidthPx / 2;
          const inflateIn = inflatePx * PX_TO_INCH * config.styleScale;
          const dashMap = { solid: 'solid', dashed: 'dash', dotted: 'dot' };
          const lineOpts = {
            color: outlineColor.hex,
            width: outlineWidthPx * 0.75 * config.styleScale,
            transparency: (1 - outlineColor.opacity) * 100,
            dashType: dashMap[outlineStyle] || 'solid',
          };
          const outlineOpts = {
            x: x - inflateIn,
            y: y - inflateIn,
            w: w + inflateIn * 2,
            h: h + inflateIn * 2,
            rotate: rotation,
            fill: { type: 'none' },
            line: lineOpts,
          };
          let outlineShape = pptx.ShapeType.rect;
          // Match the rounded corner: outline radius = border-radius + offset
          // + width/2 along the centerline. Ellipse if the inner shape is one.
          const baseRadiusPx =
            parseFloat(style.borderRadius) || 0;
          if (baseRadiusPx > 0) {
            outlineShape = pptx.ShapeType.roundRect;
            const outlineRadiusPx = Math.max(0, baseRadiusPx + offsetPx + outlineWidthPx / 2);
            outlineOpts.rectRadius = outlineRadiusPx * PX_TO_INCH * config.styleScale;
          }
          items.push({
            type: 'shape',
            zIndex: zIndex + 1,
            domOrder,
            shapeType: outlineShape,
            options: outlineOpts,
          });
        }
      }
    }

    if (hasCompositeBorder) {
      const borderSvgData = generateCompositeBorderSVG(
        widthPx,
        heightPx,
        borderRadiusValue,
        borderInfo.sides
      );
      if (borderSvgData) {
        items.push({
          type: 'image',
          zIndex: zIndex + 1,
          domOrder,
          options: { data: borderSvgData, x, y, w, h, rotate: rotation },
        });
      }
    }
  }

  return { items, job: bgJob, stopRecursion: !!textPayload };
}

function isComplexHierarchy(root) {
  // Use a simple tree traversal to find forbidden elements in the list structure
  const stack = [root];
  while (stack.length > 0) {
    const el = stack.pop();

    // 1. Layouts: Flex/Grid on LIs
    if (el.tagName === 'LI') {
      const s = window.getComputedStyle(el);
      if (s.display === 'flex' || s.display === 'grid' || s.display === 'inline-flex') return true;
    }

    // 2. Media / Icons
    if (['IMG', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME'].includes(el.tagName)) return true;
    if (isIconElement(el)) return true;

    // 3. Nested Lists (Flattening logic doesn't support nested bullets well yet)
    if (el !== root && (el.tagName === 'UL' || el.tagName === 'OL')) return true;

    // Recurse, but don't go too deep if not needed
    for (let i = 0; i < el.children.length; i++) {
      stack.push(el.children[i]);
    }
  }
  return false;
}

function createCompositeBorderItems(sides, x, y, w, h, scale, zIndex, domOrder) {
  const items = [];
  const pxToInch = 1 / 96;
  const common = { zIndex: zIndex + 1, domOrder, shapeType: 'rect' };

  if (sides.top.width > 0)
    items.push({
      ...common,
      options: { x, y, w, h: sides.top.width * pxToInch * scale, fill: { color: sides.top.color } },
    });
  if (sides.right.width > 0)
    items.push({
      ...common,
      options: {
        x: x + w - sides.right.width * pxToInch * scale,
        y,
        w: sides.right.width * pxToInch * scale,
        h,
        fill: { color: sides.right.color },
      },
    });
  if (sides.bottom.width > 0)
    items.push({
      ...common,
      options: {
        x,
        y: y + h - sides.bottom.width * pxToInch * scale,
        w,
        h: sides.bottom.width * pxToInch * scale,
        fill: { color: sides.bottom.color },
      },
    });
  if (sides.left.width > 0)
    items.push({
      ...common,
      options: {
        x,
        y,
        w: sides.left.width * pxToInch * scale,
        h,
        fill: { color: sides.left.color },
      },
    });

  return items;
}
