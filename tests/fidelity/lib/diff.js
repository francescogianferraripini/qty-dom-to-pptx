import sharp from 'sharp';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';

async function toRGBA(buffer, w, h) {
  const { data, info } = await sharp(buffer)
    .resize(w, h, { fit: 'fill' })
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  return { data, width: info.width, height: info.height };
}

/**
 * Foreground-aware fidelity score.
 *
 * The raw pixel-% metric is dominated by full-bleed backgrounds: a slide
 * with a burgundy background and broken foreground reports ~5% even when a
 * third of the content is visibly missing.
 *
 * We compute two block-level signals and take the worse:
 *
 *   1. EDGE delta — per 16x16 block, |srcEdges - dstEdges| / max(...).
 *      Catches missing/extra structural content (text, lines, icons).
 *      Robust to 1-2px stroke shifts because both blocks still hit similar
 *      density.
 *
 *   2. COLOR delta — per block, mean luminance difference vs the slide's
 *      modal background. Catches missing fills, gradient miscolors, broken
 *      tinted boxes — things the edge metric misses because they have no
 *      strong gradients inside.
 *
 * Both signals threshold and saturate so subpixel rendering noise doesn't
 * accumulate into a false floor. contentPercent = max(edge%, color%).
 */
const BLOCK_SIZE = 16;
const EDGE_THRESHOLD = 12; // per-pixel: ignore sub-12 luminance gradient as noise
const BLOCK_MIN_EDGES = 200; // per-block: skip near-empty blocks for edge metric
const COLOR_TOLERANCE = 15; // ignore <15-lum block-mean shifts (renderer noise)
const COLOR_SATURATION = 60; // a 60+ lum block-mean shift = full mismatch
const BG_CONTRAST_FLOOR = 8; // skip blocks within 8 lum of slide background

function luminance(rgba, w, h) {
  const out = new Uint8Array(w * h);
  for (let i = 0, j = 0; j < out.length; i += 4, j++) {
    out[j] = (rgba[i] * 299 + rgba[i + 1] * 587 + rgba[i + 2] * 114) / 1000;
  }
  return out;
}

function gradientL1(lum, w, h) {
  const out = new Uint8Array(w * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      const c = lum[i];
      const r = x + 1 < w ? lum[i + 1] : c;
      const d = y + 1 < h ? lum[i + w] : c;
      const v = Math.abs(r - c) + Math.abs(d - c);
      out[i] = v >= EDGE_THRESHOLD ? (v > 255 ? 255 : v) : 0;
    }
  }
  return out;
}

function maxPool3x3(blocks, bw, bh) {
  const out = new Float32Array(bw * bh);
  for (let y = 0; y < bh; y++) {
    for (let x = 0; x < bw; x++) {
      let m = 0;
      const y0 = y > 0 ? y - 1 : 0;
      const y1 = y + 1 < bh ? y + 1 : bh - 1;
      const x0 = x > 0 ? x - 1 : 0;
      const x1 = x + 1 < bw ? x + 1 : bw - 1;
      for (let yy = y0; yy <= y1; yy++) {
        const row = yy * bw;
        for (let xx = x0; xx <= x1; xx++) {
          const v = blocks[row + xx];
          if (v > m) m = v;
        }
      }
      out[y * bw + x] = m;
    }
  }
  return out;
}

function blockSums(map, w, h, block) {
  const bw = Math.ceil(w / block);
  const bh = Math.ceil(h / block);
  const sums = new Float32Array(bw * bh);
  for (let by = 0; by < bh; by++) {
    const y0 = by * block;
    const y1 = Math.min(y0 + block, h);
    for (let bx = 0; bx < bw; bx++) {
      const x0 = bx * block;
      const x1 = Math.min(x0 + block, w);
      let s = 0;
      for (let y = y0; y < y1; y++) {
        const row = y * w;
        for (let x = x0; x < x1; x++) s += map[row + x];
      }
      sums[by * bw + bx] = s;
    }
  }
  return { sums, bw, bh };
}

function median(floatArr) {
  const a = Array.from(floatArr).sort((x, y) => x - y);
  return a[Math.floor(a.length / 2)];
}

function contentDelta(srcRgba, dstRgba, w, h) {
  const srcLum = luminance(srcRgba, w, h);
  const dstLum = luminance(dstRgba, w, h);
  const srcEdges = gradientL1(srcLum, w, h);
  const dstEdges = gradientL1(dstLum, w, h);
  const { sums: srcEdgeSum, bw, bh } = blockSums(srcEdges, w, h, BLOCK_SIZE);
  const { sums: dstEdgeSum } = blockSums(dstEdges, w, h, BLOCK_SIZE);
  const { sums: srcLumSum } = blockSums(srcLum, w, h, BLOCK_SIZE);
  const { sums: dstLumSum } = blockSums(dstLum, w, h, BLOCK_SIZE);
  const blockArea = BLOCK_SIZE * BLOCK_SIZE;

  // ---- Edge-density delta ----
  // Max-pool 3x3 over block sums before comparing: tolerates ±1 block of
  // displacement (~±16px) so identical content shifted by half a block
  // doesn't produce a false miss. Genuinely missing content still scores
  // high — its whole neighborhood is empty.
  const srcEdgePool = maxPool3x3(srcEdgeSum, bw, bh);
  const dstEdgePool = maxPool3x3(dstEdgeSum, bw, bh);
  let edgeWeight = 0;
  let edgeMiss = 0;
  for (let i = 0; i < bw * bh; i++) {
    const s = srcEdgePool[i];
    const d = dstEdgePool[i];
    const m = Math.max(s, d);
    if (m < BLOCK_MIN_EDGES) continue;
    edgeWeight += m;
    edgeMiss += Math.abs(s - d);
  }
  const edgePercent = edgeWeight === 0 ? 0 : (edgeMiss / edgeWeight) * 100;

  // ---- Color-shift delta ----
  // Background = median of source block-mean luminances. Blocks within
  // BG_CONTRAST_FLOOR of bg in *both* renders are background-only and
  // ignored. Remaining blocks are weighted by max contrast against bg.
  const srcMeans = new Float32Array(bw * bh);
  for (let i = 0; i < srcMeans.length; i++) {
    srcMeans[i] = srcLumSum[i] / blockArea;
  }
  const bgLum = median(srcMeans);

  let colorWeight = 0;
  let colorMiss = 0;
  for (let i = 0; i < bw * bh; i++) {
    const s = srcLumSum[i] / blockArea;
    const d = dstLumSum[i] / blockArea;
    const sContrast = Math.abs(s - bgLum);
    const dContrast = Math.abs(d - bgLum);
    const contentMass = Math.max(sContrast, dContrast);
    if (contentMass < BG_CONTRAST_FLOOR) continue;
    colorWeight += contentMass;
    const diff = Math.abs(s - d);
    if (diff > COLOR_TOLERANCE) {
      const ratio = Math.min(
        1,
        (diff - COLOR_TOLERANCE) / COLOR_SATURATION,
      );
      colorMiss += ratio * contentMass;
    }
  }
  const colorPercent = colorWeight === 0 ? 0 : (colorMiss / colorWeight) * 100;

  return {
    contentPercent: Math.max(edgePercent, colorPercent),
    edgePercent,
    colorPercent,
  };
}

/**
 * Compare two PNG buffers at a common (width, height).
 * Returns:
 *   - diffPng: pixelmatch's annotated PNG (red = mismatch, yellow = AA)
 *   - mismatched / total / percent: classic raw pixel-percent
 *   - contentPercent: foreground-aware delta (gates the budget)
 *   - edgePercent / colorPercent: the two components, for diagnostics
 */
export async function diffPngs(aBuffer, bBuffer, width, height) {
  const a = await toRGBA(aBuffer, width, height);
  const b = await toRGBA(bBuffer, width, height);
  const diff = new PNG({ width, height });
  const mismatched = pixelmatch(a.data, b.data, diff.data, width, height, {
    threshold: 0.1,
    includeAA: false,
  });
  const total = width * height;
  const { contentPercent, edgePercent, colorPercent } = contentDelta(
    a.data,
    b.data,
    width,
    height,
  );
  const diffPng = PNG.sync.write(diff);
  return {
    diffPng,
    mismatched,
    total,
    percent: (mismatched / total) * 100,
    contentPercent,
    edgePercent,
    colorPercent,
  };
}
