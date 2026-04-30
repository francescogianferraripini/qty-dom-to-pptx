// src/utils.js

// canvas context for color normalization
let _ctx;
function getCtx() {
  if (!_ctx) _ctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true });
  return _ctx;
}

// Strip XML 1.0–illegal C0 control characters before they reach pptxgenjs.
// pptxgenjs entity-escapes &/</> itself, but it does not strip control
// characters, and PowerPoint refuses files that contain them.
// Allowed: \t (0x09), \n (0x0A), \r (0x0D).
// eslint-disable-next-line no-control-regex
const CONTROL_CHARS_RE = /[\x00-\x08\x0B\x0C\x0E-\x1F]/g;
export function sanitizeText(text) {
  if (typeof text !== 'string') return text;
  return text.replace(CONTROL_CHARS_RE, '');
}

// Document language for locale-aware text-transform. Falls back to undefined
// (which makes Intl APIs use the runtime default).
function docLang() {
  if (typeof document === 'undefined') return undefined;
  const lang = document.documentElement && document.documentElement.lang;
  return lang || undefined;
}

let _segmenter = null;
function getCapitalizeSegmenter() {
  if (_segmenter !== null) return _segmenter;
  if (typeof Intl === 'undefined' || typeof Intl.Segmenter !== 'function') {
    _segmenter = false;
    return false;
  }
  try {
    _segmenter = new Intl.Segmenter(docLang(), { granularity: 'word' });
  } catch {
    _segmenter = false;
  }
  return _segmenter;
}

export function applyTextTransform(text, transform) {
  if (!text || !transform || transform === 'none') return text;
  const lang = docLang();
  if (transform === 'uppercase') return text.toLocaleUpperCase(lang);
  if (transform === 'lowercase') return text.toLocaleLowerCase(lang);
  if (transform === 'capitalize') {
    const seg = getCapitalizeSegmenter();
    if (seg) {
      let out = '';
      for (const part of seg.segment(text)) {
        if (part.isWordLike && part.segment.length > 0) {
          // Capitalize first character of each word; leave the rest untouched.
          // Matches the spec better than \b\w/g, which mangles non-ASCII.
          out += part.segment[0].toLocaleUpperCase(lang) + part.segment.slice(1);
        } else {
          out += part.segment;
        }
      }
      return out;
    }
    return text.replace(/(^|\s)(\S)/g, (_, pre, c) => pre + c.toLocaleUpperCase(lang));
  }
  return text;
}

// Process whitespace per CSS `white-space`. Returns either a plain string
// (for collapsing modes) or an array of segments split on preserved newlines:
//   [{ text }, { breakLine: true }, { text }, ...]
// Tabs are converted to spaces uniformly so PPTX renders predictably.
export function processWhitespace(text, ws) {
  if (typeof text !== 'string' || text.length === 0) return text;
  switch (ws) {
    case 'pre':
    case 'pre-wrap':
    case 'break-spaces': {
      // Preserve runs of whitespace and line breaks. \r\n → \n.
      const normalized = text.replace(/\r\n?/g, '\n').replace(/\t/g, '    ');
      if (!normalized.includes('\n')) return normalized;
      const segs = [];
      const lines = normalized.split('\n');
      lines.forEach((line, i) => {
        if (line.length > 0) segs.push({ text: line });
        if (i < lines.length - 1) segs.push({ breakLine: true });
      });
      return segs;
    }
    case 'pre-line': {
      // Collapse runs of spaces/tabs but preserve newlines.
      const normalized = text
        .replace(/\r\n?/g, '\n')
        .replace(/[ \t]+/g, ' ');
      if (!normalized.includes('\n')) return normalized;
      const segs = [];
      const lines = normalized.split('\n');
      lines.forEach((line, i) => {
        if (line.length > 0) segs.push({ text: line });
        if (i < lines.length - 1) segs.push({ breakLine: true });
      });
      return segs;
    }
    case 'normal':
    case 'nowrap':
    default:
      return text.replace(/[\n\r\t]+/g, ' ').replace(/\s{2,}/g, ' ');
  }
}

// Map CSS generic font-family keywords to PowerPoint-friendly fallbacks.
const GENERIC_FONT_MAP = {
  serif: 'Cambria',
  'sans-serif': 'Calibri',
  monospace: 'Consolas',
  cursive: 'Comic Sans MS',
  fantasy: 'Impact',
  'system-ui': 'Calibri',
  'ui-serif': 'Cambria',
  'ui-sans-serif': 'Calibri',
  'ui-monospace': 'Consolas',
  'ui-rounded': 'Calibri',
  emoji: 'Segoe UI Emoji',
  math: 'Cambria Math',
  fangsong: 'SimSun',
};

// Names that frequently lead a CSS font stack on Mac/Linux but aren't shipped
// with Windows PowerPoint. When such a name appears AND the chain ends with a
// generic (e.g. `monospace`), we skip it so the generic's mapping wins —
// otherwise PowerPoint/LibreOffice render the unknown name with the document
// default (a proportional face), which loses monospace fidelity for code
// blocks. Embeddable webfonts (Poppins, Inter, etc.) are NOT in this list and
// remain the first choice so font embedding still works.
const NON_WINDOWS_FONTS = new Set([
  'sfmono-regular', 'sf mono', '-apple-system', 'blinkmacsystemfont',
  'menlo', 'monaco', 'apple color emoji',
  'liberation mono', 'liberation sans', 'liberation serif',
  'dejavu sans', 'dejavu sans mono', 'dejavu serif',
  'ubuntu', 'ubuntu mono', 'cantarell', 'noto sans',
]);

// Resolves a CSS `font-family` value to the *single* font name we emit as the
// OOXML `typeface` attribute. The OOXML schema treats `typeface` as one font
// name — not a CSS-style fallback list — so PowerPoint and LibreOffice both
// look up the entire string verbatim and fall back to the document default
// when the name doesn't resolve. Returning a comma list here was silently
// breaking code-block fidelity on the quantyca deck.
//
// Strategy: walk the chain left-to-right, mapping generics inline. Skip names
// in NON_WINDOWS_FONTS when the chain has a terminal generic (so monospace's
// `Consolas` mapping wins over a Mac-only `SFMono-Regular`). Otherwise the
// first concrete name wins, which preserves embedded webfonts.
//
// Important: pptxgenjs concatenates the returned string straight into an XML
// attribute value without entity-escaping, so we strip embedded quotes.
export function resolveFontFaceList(fontFamilyStr) {
  if (!fontFamilyStr) return 'Calibri';

  const rawTokens = fontFamilyStr.split(',');
  const tokens = [];
  for (let token of rawTokens) {
    let name = token.trim();
    if (
      (name.startsWith('"') && name.endsWith('"')) ||
      (name.startsWith("'") && name.endsWith("'"))
    ) {
      name = name.slice(1, -1).trim();
    }
    name = name.replace(/["']/g, '');
    if (name) tokens.push(name);
  }
  if (!tokens.length) return 'Calibri';

  const lastGeneric = GENERIC_FONT_MAP[tokens[tokens.length - 1].toLowerCase()];

  for (const name of tokens) {
    const lower = name.toLowerCase();
    const generic = GENERIC_FONT_MAP[lower];
    if (generic) return generic;
    if (lastGeneric && NON_WINDOWS_FONTS.has(lower)) continue;
    return name;
  }

  return lastGeneric || 'Calibri';
}

// Walks ancestors of `root` and accumulates uniform-scale components from
// each ancestor's `transform` matrix. `getBoundingClientRect()` is already
// post-transform, so positions don't need correcting — but `getComputedStyle`
// returns logical (pre-transform) lengths, and those need to be multiplied
// by `ancestorScale` before the existing px→pt conversion.
//
// Returns a single uniform scalar (the geometric mean of sx and sy). Real-
// world cases (reveal.js fitting, framework zoom) use uniform scale; for
// non-uniform transforms we accept the small distortion of using the mean
// rather than threading {sx, sy} through every helper.
export function getAncestorScale(root) {
  if (!root || typeof window === 'undefined') return 1;
  let sx = 1;
  let sy = 1;
  let el = root.parentElement;
  while (el) {
    let style;
    try {
      style = window.getComputedStyle(el);
    } catch {
      break;
    }
    const t = style && style.transform;
    if (t && t !== 'none') {
      const sc = scaleFromTransform(t);
      if (sc) {
        sx *= sc.sx;
        sy *= sc.sy;
      }
    }
    el = el.parentElement;
  }
  // Guard against degenerate / zero scales that would zero-out font sizes.
  if (!isFinite(sx) || sx <= 0) sx = 1;
  if (!isFinite(sy) || sy <= 0) sy = 1;
  return Math.sqrt(sx * sy);
}

// Extract a uniform scale factor from a CSS transform string. Returns 1 for
// transforms that contain no scale component (rotation-only, translate-only,
// or `none`). Used by the traversal-time cumulative scale tracker so font-
// size readings can be corrected for transforms that live BELOW root (the
// reveal.js `.slides` case).
export function nodeOwnScale(transformStr) {
  if (!transformStr || transformStr === 'none') return 1;
  const sc = scaleFromTransform(transformStr);
  if (!sc) return 1;
  if (!isFinite(sc.sx) || sc.sx <= 0 || !isFinite(sc.sy) || sc.sy <= 0) return 1;
  return Math.sqrt(sc.sx * sc.sy);
}

// Decode a `content` value as returned by getComputedStyle. Strips the
// outer quotes and processes CSS hex escapes (`\d'acqua` → `D'acqua`).
// Returns the empty string for `none`, `normal`, or unparseable values
// (counters, attr(), open-quote, etc.).
export function decodePseudoContent(contentRaw) {
  if (!contentRaw) return '';
  const s = contentRaw.trim();
  if (s === 'none' || s === 'normal' || s === 'no-open-quote' || s === 'no-close-quote') return '';
  // Match a single quoted string at the start. CSS allows a sequence of
  // tokens (e.g. `"foo " counter(x)`); we only render the leading literal
  // — counters/attrs require live DOM context that's expensive to mirror.
  const m = s.match(/^(['"])((?:\\.|(?!\1).)*)\1/);
  if (!m) return '';
  const inner = m[2];
  let out = '';
  let i = 0;
  while (i < inner.length) {
    const c = inner[i];
    if (c === '\\') {
      const hex = inner.slice(i + 1).match(/^([0-9a-fA-F]{1,6})\s?/);
      if (hex) {
        const code = parseInt(hex[1], 16);
        if (code) out += String.fromCodePoint(code);
        i += 1 + hex[0].length;
      } else if (i + 1 < inner.length) {
        out += inner[i + 1];
        i += 2;
      } else {
        i += 1;
      }
    } else {
      out += c;
      i++;
    }
  }
  return out;
}

// Parse a CSS length OR percentage. Returns numeric pixels (resolving
// percentages against `ref`) or null for `auto`/missing.
function parseLengthOrPct(v, ref) {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s || s === 'auto') return null;
  if (s.endsWith('%')) {
    const pct = parseFloat(s);
    if (!isFinite(pct)) return null;
    return (pct / 100) * (ref || 0);
  }
  const n = parseFloat(s);
  return isFinite(n) ? n : null;
}

// Parse one of the four position sides (top/right/bottom/left) for an
// absolutely-positioned pseudo. Percentages resolve against `ref` (the
// parent's width for left/right, height for top/bottom).
function parseSide(v, ref) {
  return parseLengthOrPct(v, ref);
}

// Measure a text-based pseudo with `width: auto` by mirroring its content
// into a hidden, absolutely-positioned span and reading the rendered rect.
// Browsers don't expose pseudo-element rects directly (no `getBoxQuads`
// pseudo support outside Firefox), so a proxy is the only portable option.
function measureTextProxy(pStyle, text) {
  if (typeof document === 'undefined' || !text) return null;
  let proxy;
  try {
    proxy = document.createElement('span');
    const props = [
      'fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'fontVariant',
      'fontStretch', 'letterSpacing', 'wordSpacing', 'textTransform',
      'paddingLeft', 'paddingRight', 'paddingTop', 'paddingBottom',
      'borderLeftWidth', 'borderRightWidth', 'borderTopWidth', 'borderBottomWidth',
      'borderLeftStyle', 'borderRightStyle', 'borderTopStyle', 'borderBottomStyle',
      'lineHeight', 'whiteSpace', 'boxSizing', 'textIndent', 'fontFeatureSettings',
    ];
    for (const p of props) {
      const v = pStyle[p];
      if (v) proxy.style[p] = v;
    }
    proxy.style.position = 'absolute';
    proxy.style.visibility = 'hidden';
    proxy.style.pointerEvents = 'none';
    proxy.style.left = '-99999px';
    proxy.style.top = '-99999px';
    proxy.style.display = 'inline-block';
    proxy.style.maxWidth = 'none';
    proxy.textContent = text;
    document.body.appendChild(proxy);
    const rect = proxy.getBoundingClientRect();
    document.body.removeChild(proxy);
    proxy = null;
    if (!rect.width || !rect.height) return null;
    return { width: rect.width, height: rect.height };
  } catch {
    if (proxy && proxy.parentNode) proxy.parentNode.removeChild(proxy);
    return null;
  }
}

// Detect whether a pseudo-element has a layout box worth rendering — i.e.
// it contributes a visible rectangle (background, border, transform, or
// inline content text) rather than nothing at all. Returns null if there's
// nothing to draw, otherwise a description of where the box sits relative
// to the parent's bounding rect along with the styles needed to render it.
//
// This is a best-effort approximation. Pixel-perfect pseudo-element
// positioning would require `getBoxQuads({pseudo})`, which is not portably
// supported. We cover the patterns that actually appear in real-world decks:
// hero/divider full-bleed backgrounds (auto width/height resolved from
// inset shorthand), accent underlines, corner brand marks (text content with
// auto width measured via DOM proxy), abs-positioned badges, circular dots.
export function measurePseudoBox(parent, which /* 'before' | 'after' */) {
  if (!parent || parent.nodeType !== 1 || typeof window === 'undefined') return null;
  let pStyle;
  try {
    pStyle = window.getComputedStyle(parent, `::${which}`);
  } catch {
    return null;
  }
  if (!pStyle) return null;

  const display = pStyle.display;
  const content = pStyle.content;
  // `content: normal` and `none` mean the pseudo isn't generated.
  if (!content || content === 'none' || content === 'normal' || display === 'none') {
    return null;
  }
  if (pStyle.visibility === 'hidden') return null;
  const opacityNum = parseFloat(pStyle.opacity);
  if (!isNaN(opacityNum) && opacityNum === 0) return null;

  const contentText = decodePseudoContent(content);

  const parentRect = parent.getBoundingClientRect();
  const parentStyle = window.getComputedStyle(parent);
  const position = pStyle.position;

  // Resolve dimensions. % resolves against parent rect (the closest
  // positioned ancestor would be more correct, but for the decks we care
  // about the parent is positioned via reveal's slide layout and is the
  // containing block). `auto` stays null for now and may be filled in later
  // from inset sides or a text proxy.
  let wRaw = parseLengthOrPct(pStyle.width, parentRect.width);
  let hRaw = parseLengthOrPct(pStyle.height, parentRect.height);

  const sideLeft = position === 'absolute' || position === 'fixed'
    ? parseSide(pStyle.left, parentRect.width) : null;
  const sideRight = position === 'absolute' || position === 'fixed'
    ? parseSide(pStyle.right, parentRect.width) : null;
  const sideTop = position === 'absolute' || position === 'fixed'
    ? parseSide(pStyle.top, parentRect.height) : null;
  const sideBottom = position === 'absolute' || position === 'fixed'
    ? parseSide(pStyle.bottom, parentRect.height) : null;

  let left = NaN;
  let top = NaN;

  if (position === 'absolute' || position === 'fixed') {
    const ml = parseFloat(pStyle.marginLeft) || 0;
    const mt = parseFloat(pStyle.marginTop) || 0;

    // Width from inset: when both left and right are set and width is auto,
    // the box stretches between them. This is the `inset: 0 0 0 50%` case.
    if (wRaw == null && sideLeft != null && sideRight != null) {
      left = parentRect.left + sideLeft + ml;
      const right = parentRect.right - sideRight;
      wRaw = right - left;
    } else if (sideLeft != null) {
      left = parentRect.left + sideLeft + ml;
    } else if (sideRight != null && wRaw != null) {
      left = parentRect.right - sideRight - wRaw;
    }

    if (hRaw == null && sideTop != null && sideBottom != null) {
      top = parentRect.top + sideTop + mt;
      const bottom = parentRect.bottom - sideBottom;
      hRaw = bottom - top;
    } else if (sideTop != null) {
      top = parentRect.top + sideTop + mt;
    } else if (sideBottom != null && hRaw != null) {
      top = parentRect.bottom - sideBottom - hRaw;
    }

    if (isNaN(left)) left = parentRect.left;
    if (isNaN(top)) top = parentRect.top;
  } else {
    // Block / static / relative pseudo. Position at parent's content edge:
    // ::before stacks at the top, ::after at the bottom. For inline-block
    // pseudos this is approximate but matches the dominant decorative usage.
    const padLeft = parseFloat(parentStyle.paddingLeft) || 0;
    const padTop = parseFloat(parentStyle.paddingTop) || 0;
    const padBottom = parseFloat(parentStyle.paddingBottom) || 0;
    const borderTop = parseFloat(parentStyle.borderTopWidth) || 0;
    const borderBottom = parseFloat(parentStyle.borderBottomWidth) || 0;
    const ml = parseFloat(pStyle.marginLeft) || 0;
    const mt = parseFloat(pStyle.marginTop) || 0;
    const mb = parseFloat(pStyle.marginBottom) || 0;

    left = parentRect.left + borderTop + padLeft + ml;
    if (which === 'before') {
      top = parentRect.top + borderTop + padTop + mt;
    } else {
      top = parentRect.bottom - borderBottom - padBottom - mb - (hRaw || 0);
    }
  }

  // Text-based pseudo with auto sizing: proxy-measure the rendered text.
  // Re-resolve any side-anchored coordinate that depended on the missing dim.
  if ((wRaw == null || wRaw <= 0 || hRaw == null || hRaw <= 0) && contentText) {
    const measured = measureTextProxy(pStyle, contentText);
    if (measured) {
      if (wRaw == null || wRaw <= 0) wRaw = measured.width;
      if (hRaw == null || hRaw <= 0) hRaw = measured.height;
      if (position === 'absolute' || position === 'fixed') {
        if (sideLeft == null && sideRight != null) {
          left = parentRect.right - sideRight - wRaw;
        }
        if (sideTop == null && sideBottom != null) {
          top = parentRect.bottom - sideBottom - hRaw;
        }
      } else if (which === 'after') {
        // Recompute bottom-anchored ::after now that we know its height.
        const padBottom = parseFloat(parentStyle.paddingBottom) || 0;
        const borderBottom = parseFloat(parentStyle.borderBottomWidth) || 0;
        const mb = parseFloat(pStyle.marginBottom) || 0;
        top = parentRect.bottom - borderBottom - padBottom - mb - hRaw;
      }
    }
  }

  if (wRaw == null || hRaw == null || wRaw <= 0 || hRaw <= 0) return null;

  // Visibility check: at least one paint property must produce a visible
  // mark. A pseudo with empty content and no fill/border/transform would
  // never show, so skip it.
  const bg = parseColor(pStyle.backgroundColor);
  const hasBg = bg.hex && bg.opacity > 0;
  const borderW = parseFloat(pStyle.borderWidth) || 0;
  const borderColor = parseColor(pStyle.borderColor);
  const hasBorder = borderW > 0 && borderColor.hex && borderColor.opacity > 0;
  const hasTransform = pStyle.transform && pStyle.transform !== 'none';
  const bgImg = pStyle.backgroundImage;
  const hasBgImage = bgImg && bgImg !== 'none';
  if (!hasBg && !hasBorder && !hasTransform && !hasBgImage && !contentText) return null;

  return {
    rect: { left, top, width: wRaw, height: hRaw },
    style: pStyle,
    parentStyle,
    contentText,
  };
}

function scaleFromTransform(transformStr) {
  const m2 = transformStr.match(/^matrix\(([^)]+)\)$/);
  if (m2) {
    const v = m2[1].split(',').map((s) => parseFloat(s.trim()));
    if (v.length >= 4) {
      return {
        sx: Math.sqrt(v[0] * v[0] + v[1] * v[1]),
        sy: Math.sqrt(v[2] * v[2] + v[3] * v[3]),
      };
    }
  }
  const m3 = transformStr.match(/^matrix3d\(([^)]+)\)$/);
  if (m3) {
    const v = m3[1].split(',').map((s) => parseFloat(s.trim()));
    if (v.length >= 16) {
      return {
        sx: Math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]),
        sy: Math.sqrt(v[4] * v[4] + v[5] * v[5] + v[6] * v[6]),
      };
    }
  }
  return null;
}

function getTableBorder(style, side, scale) {
  const widthStr = style[`border${side}Width`];
  const styleStr = style[`border${side}Style`];
  const colorStr = style[`border${side}Color`];

  const width = parseFloat(widthStr) || 0;
  if (width === 0 || styleStr === 'none' || styleStr === 'hidden') {
    return null;
  }

  const color = parseColor(colorStr);
  if (!color.hex || color.opacity === 0) return null;

  let dash = 'solid';
  if (styleStr === 'dashed') dash = 'dash';
  if (styleStr === 'dotted') dash = 'dot';

  return {
    pt: width * 0.75 * scale, // Convert px to pt
    color: color.hex,
    type: dash,
  };
}

/**
 * Extracts native table data for PptxGenJS.
 */
export function extractTableData(node, scale, styleScale = scale) {
  const rows = [];
  const colWidths = [];

  // 1. Calculate Column Widths based on the first row of cells
  // We look at the first <tr>'s children to determine visual column widths.
  // Note: This assumes a fixed grid. Complex colspan/rowspan on the first row
  // might skew widths, but getBoundingClientRect captures the rendered result.
  const firstRow = node.querySelector('tr');
  // Track each colWidths slot's nowrap flag so we can inflate columns whose
  // CSS asks the browser to shrink-wrap them (`width:1%; white-space:nowrap`).
  // The browser's auto-sizing produces a column tight to its widest content
  // *for the font Chrome chose*; PPTX/LibreOffice often falls back to a
  // slightly different (typically wider) face, which makes the same content
  // overflow and wrap — visible on highlight.js line-number columns where
  // 2-digit numbers like "10" split into "1" / "0".
  const colNowrap = [];
  if (firstRow) {
    const cells = Array.from(firstRow.children);
    cells.forEach((cell) => {
      const rect = cell.getBoundingClientRect();
      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const wIn = (rect.width * (1 / 96) * scale) / colspan;
      const cs = window.getComputedStyle(cell);
      const isNowrap = cs.whiteSpace === 'nowrap' || cs.whiteSpace === 'pre';
      for (let i = 0; i < colspan; i++) {
        colWidths.push(wIn);
        colNowrap.push(isNowrap);
      }
    });
  }
  // Inflate nowrap columns by ~10%, donating that width from the widest
  // wrapping column so the table's total width stays put.
  colNowrap.forEach((isNowrap, i) => {
    if (!isNowrap) return;
    const bonus = colWidths[i] * 0.1;
    let widestI = -1;
    let widest = 0;
    colWidths.forEach((w, j) => {
      if (j === i || colNowrap[j]) return;
      if (w > widest) { widest = w; widestI = j; }
    });
    if (widestI >= 0 && colWidths[widestI] - bonus > 0) {
      colWidths[i] += bonus;
      colWidths[widestI] -= bonus;
    }
  });

  const tableStyle = window.getComputedStyle(node);
  const borderSpacing = tableStyle.borderSpacing.split(' ');
  const hSpace = parseFloat(borderSpacing[0]) || 0;
  const vSpace = parseFloat(borderSpacing[1] || borderSpacing[0]) || 0;
  const hSpaceIn = hSpace * (1 / 96) * styleScale;
  const vSpaceIn = vSpace * (1 / 96) * styleScale;

  // 2. Iterate Rows
  const trList = node.querySelectorAll('tr');
  trList.forEach((tr) => {
    const rowData = [];
    const cellList = Array.from(tr.children).filter((c) => ['TD', 'TH'].includes(c.tagName));

      cellList.forEach((cell) => {
        const style = window.getComputedStyle(cell);
        const cellParts = collectTextParts(cell, style, styleScale);
        // Fallback to plain text if collectTextParts returns empty/invalid
        const cellText = (cellParts && cellParts.length > 0) ? cellParts
          : sanitizeText(cell.innerText.replace(/[\n\r\t]+/g, ' ').trim());

      // A. Text Style
      const textStyle = getTextStyle(style, styleScale);

      // B. Cell Background
      let bg = parseColor(style.backgroundColor);
      if (
        (!bg.hex || bg.opacity === 0) &&
        style.backgroundImage &&
        style.backgroundImage !== 'none'
      ) {
        const fallback = getGradientFallbackColor(style.backgroundImage);
        if (fallback) bg = parseColor(fallback);
      }
      const fill = bg.hex && bg.opacity > 0 ? { color: bg.hex } : null;

      // C. Alignment
      let align = 'left';
      if (style.textAlign === 'center') align = 'center';
      if (style.textAlign === 'right' || style.textAlign === 'end') align = 'right';

      let valign = 'top';
      if (style.verticalAlign === 'middle') valign = 'middle';
      if (style.verticalAlign === 'bottom') valign = 'bottom';

      // D. Padding (Margins in PPTX) — keep in inches.
      // PptxGenJS picks units adaptively: if margin[0] >= 1 it treats the whole
      // array as points, else as inches. Mixing intended-pt and intended-in
      // values (e.g. tiny top + large right) flips the interpretation and
      // explodes large values to inches. Pass everything in inches so the
      // unit branch is stable regardless of magnitude.
      const padding = getPadding(style, styleScale);
      const margin = [
        padding[0] + vSpaceIn / 2, // top
        padding[1] + hSpaceIn / 2, // right
        padding[2] + vSpaceIn / 2, // bottom
        padding[3] + hSpaceIn / 2, // left
      ];

      // E. Borders (logical width → styleScale)
      const borderTop = getTableBorder(style, 'Top', styleScale);
      const borderRight = getTableBorder(style, 'Right', styleScale);
      const borderBottom = getTableBorder(style, 'Bottom', styleScale);
      const borderLeft = getTableBorder(style, 'Left', styleScale);

      // F. Construct Cell Object
      rowData.push({
        text: cellText,
        options: {
          color: textStyle.color,
          fontFace: textStyle.fontFace,
          fontSize: textStyle.fontSize,
          bold: textStyle.bold,
          italic: textStyle.italic,
          underline: textStyle.underline,

          fill: fill,
          align: align,
          valign: valign,
          margin: margin,

          rowspan: parseInt(cell.getAttribute('rowspan')) || null,
          colspan: parseInt(cell.getAttribute('colspan')) || null,

          border: [borderTop, borderRight, borderBottom, borderLeft],
        },
      });
    });

    if (rowData.length > 0) {
      rows.push(rowData);
    }
  });

  return { rows, colWidths };
}

// Checks if any parent element has overflow: hidden which would clip this element
export function isClippedByParent(node) {
  let parent = node.parentElement;
  while (parent && parent !== document.body) {
    const style = window.getComputedStyle(parent);
    const overflow = style.overflow;
    if (overflow === 'hidden' || overflow === 'clip') {
      return true;
    }
    parent = parent.parentElement;
  }
  return false;
}

// Helper to save gradient text
// Helper to save gradient text: extracts the first color from a gradient string
export function getGradientFallbackColor(bgImage) {
  if (!bgImage || bgImage === 'none') return null;

  // 1. Extract content inside function(...)
  // Handles linear-gradient(...), radial-gradient(...), repeating-linear-gradient(...)
  const match = bgImage.match(/gradient\((.*)\)/);
  if (!match) return null;

  const content = match[1];

  // 2. Split by comma, respecting parentheses (to avoid splitting inside rgb(), oklch(), etc.)
  const parts = [];
  let current = '';
  let parenDepth = 0;

  for (const char of content) {
    if (char === '(') parenDepth++;
    if (char === ')') parenDepth--;
    if (char === ',' && parenDepth === 0) {
      parts.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  if (current) parts.push(current.trim());

  // 3. Find first part that is a color (skip angle/direction)
  for (const part of parts) {
    // Ignore directions (to right) or angles (90deg, 0.5turn)
    if (/^(to\s|[\d.]+(deg|rad|turn|grad))/.test(part)) continue;

    // Extract color: Remove trailing position (e.g. "red 50%" -> "red")
    // Regex matches whitespace + number + unit at end of string
    const colorPart = part.replace(/\s+(-?[\d.]+(%|px|em|rem|ch|vh|vw)?)$/, '');

    // Check if it's not just a number (some gradients might have bare numbers? unlikely in standard syntax)
    if (colorPart) return colorPart;
  }

  return null;
}

function mapDashType(style) {
  if (style === 'dashed') return 'dash';
  if (style === 'dotted') return 'dot';
  return 'solid';
}

/**
 * Analyzes computed border styles and determines the rendering strategy.
 */
export function getBorderInfo(style, scale) {
  const topColor = parseColor(style.borderTopColor);
  const rightColor = parseColor(style.borderRightColor);
  const bottomColor = parseColor(style.borderBottomColor);
  const leftColor = parseColor(style.borderLeftColor);
  const top = {
    width: parseFloat(style.borderTopWidth) || 0,
    style: style.borderTopStyle,
    color: topColor.hex,
    opacity: topColor.opacity,
  };
  const right = {
    width: parseFloat(style.borderRightWidth) || 0,
    style: style.borderRightStyle,
    color: rightColor.hex,
    opacity: rightColor.opacity,
  };
  const bottom = {
    width: parseFloat(style.borderBottomWidth) || 0,
    style: style.borderBottomStyle,
    color: bottomColor.hex,
    opacity: bottomColor.opacity,
  };
  const left = {
    width: parseFloat(style.borderLeftWidth) || 0,
    style: style.borderLeftStyle,
    color: leftColor.hex,
    opacity: leftColor.opacity,
  };

  const hasAnyBorder = top.width > 0 || right.width > 0 || bottom.width > 0 || left.width > 0;
  if (!hasAnyBorder) return { type: 'none' };

  // Fully-transparent borders: many CSS resets place `border: 1px solid
  // transparent` so :hover / :focus can flip color without shifting layout.
  // The layout space is already baked into getBoundingClientRect, so we skip
  // the stroke entirely. Treat as no-border for downstream consumers.
  const allTransparent =
    top.opacity === 0 && right.opacity === 0 && bottom.opacity === 0 && left.opacity === 0;
  if (allTransparent) return { type: 'none' };

  // Check if all sides are uniform
  const isUniform =
    top.width === right.width &&
    top.width === bottom.width &&
    top.width === left.width &&
    top.style === right.style &&
    top.style === bottom.style &&
    top.style === left.style &&
    top.color === right.color &&
    top.color === bottom.color &&
    top.color === left.color;

  if (isUniform) {
    return {
      type: 'uniform',
      options: {
        width: top.width * 0.75 * scale,
        color: top.color,
        transparency: (1 - topColor.opacity) * 100,
        dashType: mapDashType(top.style),
      },
    };
  } else {
    return {
      type: 'composite',
      sides: { top, right, bottom, left },
    };
  }
}

/**
 * Generates an SVG image for composite borders that respects border-radius.
 */
export function generateCompositeBorderSVG(w, h, radius, sides) {
  radius = radius / 2; // Adjust for SVG rendering
  const clipId = 'clip_' + Math.random().toString(36).substr(2, 9);
  let borderRects = '';

  if (sides.top.width > 0 && sides.top.color) {
    borderRects += `<rect x="0" y="0" width="${w}" height="${sides.top.width}" fill="#${sides.top.color}" />`;
  }
  if (sides.right.width > 0 && sides.right.color) {
    borderRects += `<rect x="${w - sides.right.width}" y="0" width="${sides.right.width}" height="${h}" fill="#${sides.right.color}" />`;
  }
  if (sides.bottom.width > 0 && sides.bottom.color) {
    borderRects += `<rect x="0" y="${h - sides.bottom.width}" width="${w}" height="${sides.bottom.width}" fill="#${sides.bottom.color}" />`;
  }
  if (sides.left.width > 0 && sides.left.color) {
    borderRects += `<rect x="0" y="0" width="${sides.left.width}" height="${h}" fill="#${sides.left.color}" />`;
  }

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
        <defs>
            <clipPath id="${clipId}">
                <rect x="0" y="0" width="${w}" height="${h}" rx="${radius}" ry="${radius}" />
            </clipPath>
        </defs>
        <g clip-path="url(#${clipId})">
            ${borderRects}
        </g>
    </svg>`;

  return 'data:image/svg+xml;base64,' + btoa(svg);
}

/**
 * Generates an SVG data URL for a solid shape with non-uniform corner radii.
 * Accepts either scalar radii (`{tl, tr, br, bl}`) or elliptical per-corner
 * pairs (`{tl: {x, y}, tr: {x, y}, ...}`).
 */
export function generateCustomShapeSVG(w, h, color, opacity, radii) {
  const r = normalizeRadiiXY(radii);
  clampRadiiXY(r, w, h);
  return wrapShapeSvg(
    w,
    h,
    buildEllipticalPath(w, h, r),
    `fill="#${color}" fill-opacity="${opacity}"`
  );
}

// Convert either scalar `{tl, tr, br, bl}` or elliptical `{tl:{x,y}, ...}`
// radii into the elliptical form. Mutates a fresh copy.
function normalizeRadiiXY(radii) {
  const out = {};
  for (const k of ['tl', 'tr', 'br', 'bl']) {
    const v = radii[k];
    if (v && typeof v === 'object') {
      out[k] = { x: v.x || 0, y: v.y || 0 };
    } else {
      const n = v || 0;
      out[k] = { x: n, y: n };
    }
  }
  return out;
}

// CSS Backgrounds 3 §5.5: when adjacent corner radii would overlap, scale
// every radius by the same factor so the largest pair fits the available
// edge length. Computed independently for each axis (rx vs ry) since the
// shorthand only enforces overlap-free along its own axis.
function clampRadiiXY(r, w, h) {
  const { tl, tr, br, bl } = r;
  const factor = Math.min(
    w / (tl.x + tr.x) || Infinity,
    w / (bl.x + br.x) || Infinity,
    h / (tl.y + bl.y) || Infinity,
    h / (tr.y + br.y) || Infinity,
    1
  );
  if (factor < 1) {
    for (const k of ['tl', 'tr', 'br', 'bl']) {
      r[k].x *= factor;
      r[k].y *= factor;
    }
  }
}

function buildEllipticalPath(w, h, r) {
  const { tl, tr, br, bl } = r;
  return (
    `M ${tl.x} 0 ` +
    `L ${w - tr.x} 0 ` +
    `A ${tr.x} ${tr.y} 0 0 1 ${w} ${tr.y} ` +
    `L ${w} ${h - br.y} ` +
    `A ${br.x} ${br.y} 0 0 1 ${w - br.x} ${h} ` +
    `L ${bl.x} ${h} ` +
    `A ${bl.x} ${bl.y} 0 0 1 0 ${h - bl.y} ` +
    `L 0 ${tl.y} ` +
    `A ${tl.x} ${tl.y} 0 0 1 ${tl.x} 0 ` +
    `Z`
  );
}

function wrapShapeSvg(w, h, pathD, fillAttrs, defs = '') {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    defs +
    `<path d="${pathD}" ${fillAttrs} />` +
    `</svg>`;
  return 'data:image/svg+xml;base64,' + btoa(svg);
}

// --- REPLACE THE EXISTING parseColor FUNCTION ---
// Sentinel hex used to detect when a fillStyle assignment was rejected
// by the browser (invalid CSS color). Picked to be deterministic and
// extremely unlikely to appear as real content.
const PARSE_COLOR_SENTINEL = '#01fe02';
const PARSE_COLOR_SENTINEL_NORMALIZED = '#01fe02';

export function parseColor(str) {
  if (!str || str === 'transparent' || (typeof str === 'string' && str.trim() === 'rgba(0, 0, 0, 0)')) {
    return { hex: null, opacity: 0 };
  }

  // Pre-normalize the rare 16-digit hex form (#rrrrggggbbbbaaaa, 4 digits per
  // channel) before handing to the browser — Chromium rejects it. Take the
  // high byte of each channel.
  let probe = str;
  if (typeof probe === 'string') {
    const trimmed = probe.trim();
    if (/^#[0-9a-fA-F]{16}$/.test(trimmed)) {
      const r = trimmed.slice(1, 3);
      const g = trimmed.slice(5, 7);
      const b = trimmed.slice(9, 11);
      const a = trimmed.slice(13, 15);
      probe = `#${r}${g}${b}${a}`;
    } else if (/^#[0-9a-fA-F]{12}$/.test(trimmed)) {
      // Legacy 4-digit-per-channel without alpha (#rrrrggggbbbb).
      const r = trimmed.slice(1, 3);
      const g = trimmed.slice(5, 7);
      const b = trimmed.slice(9, 11);
      probe = `#${r}${g}${b}`;
    }
  }

  const ctx = getCtx();
  // Detect parse failure: if assignment to fillStyle is rejected, fillStyle
  // keeps its prior value. Seed a sentinel so the rejected case is
  // identifiable instead of silently returning the previous color (typically
  // black, which the engine would then ship as a real, wrong color).
  ctx.fillStyle = PARSE_COLOR_SENTINEL;
  ctx.fillStyle = probe;
  const computed = ctx.fillStyle;
  if (computed === PARSE_COLOR_SENTINEL_NORMALIZED) {
    return { hex: null, opacity: 0, parseFailed: true };
  }

  // 1. Handle Hex Output (e.g. #ff0000) - Fast Path
  if (computed.startsWith('#')) {
    let hex = computed.slice(1);
    let opacity = 1;
    if (hex.length === 3)
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    if (hex.length === 4)
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    if (hex.length === 8) {
      opacity = parseInt(hex.slice(6), 16) / 255;
      hex = hex.slice(0, 6);
    }
    return { hex: hex.toUpperCase(), opacity };
  }

  // 2. Handle RGB/RGBA Output (standard) - Fast Path
  if (computed.startsWith('rgb')) {
    const match = computed.match(/[\d.]+/g);
    if (match && match.length >= 3) {
      const r = parseInt(match[0]);
      const g = parseInt(match[1]);
      const b = parseInt(match[2]);
      const a = match.length > 3 ? parseFloat(match[3]) : 1;
      const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
      return { hex, opacity: a };
    }
  }

  // 3. Fallback: Browser returned a format we don't parse (oklch, lab, color(srgb...), etc.)
  // Use Canvas API to convert to sRGB. Note: this clips wide-gamut colors to
  // sRGB — see SUPPORTED.md.
  ctx.clearRect(0, 0, 1, 1);
  ctx.fillRect(0, 0, 1, 1);
  const data = ctx.getImageData(0, 0, 1, 1).data;
  // data = [r, g, b, a]
  const r = data[0];
  const g = data[1];
  const b = data[2];
  const a = data[3] / 255;

  if (a === 0) return { hex: null, opacity: 0 };

  const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  return { hex, opacity: a };
}

export function getPadding(style, scale) {
  const pxToInch = 1 / 96;
  return [
    (parseFloat(style.paddingTop) || 0) * pxToInch * scale,
    (parseFloat(style.paddingRight) || 0) * pxToInch * scale,
    (parseFloat(style.paddingBottom) || 0) * pxToInch * scale,
    (parseFloat(style.paddingLeft) || 0) * pxToInch * scale,
  ];
}

// pptxgenjs's `inset` option only accepts a scalar (an array silently drops);
// its `margin` option accepts an array `[left, right, bottom, top]` in points
// and writes through to lIns/rIns/bIns/tIns. Convert getPadding's CSS-order
// inches into that shape so CSS padding lands in the rendered text frame.
export function paddingToPptxMargin(padding) {
  return [padding[3] * 72, padding[1] * 72, padding[2] * 72, padding[0] * 72];
}

export function getSoftEdges(filterStr, scale) {
  if (!filterStr || filterStr === 'none') return null;
  const match = filterStr.match(/blur\(([\d.]+)px\)/);
  if (match) return parseFloat(match[1]) * 0.75 * scale;
  return null;
}

export function getTextStyle(style, scale) {
  let colorObj = parseColor(style.color);

  const bgClip = style.webkitBackgroundClip || style.backgroundClip;
  if (colorObj.opacity === 0 && bgClip === 'text') {
    const fallback = getGradientFallbackColor(style.backgroundImage);
    if (fallback) colorObj = parseColor(fallback);
  }

  let lineSpacing = null;
  const fontSizePx = parseFloat(style.fontSize);
  const lhStr = style.lineHeight;

  if (lhStr && lhStr !== 'normal') {
    let lhPx = parseFloat(lhStr);

    // Edge Case: If browser returns a raw multiplier (e.g. "1.5")
    // we must multiply by font size to get the height in pixels.
    // (Note: getComputedStyle usually returns 'px', but inline styles might differ)
    if (/^[0-9.]+$/.test(lhStr)) {
      lhPx = lhPx * fontSizePx;
    }

    if (!isNaN(lhPx) && lhPx > 0) {
      // Convert Pixel Height to Point Height (1px = 0.75pt)
      // And apply the global layout scale.
      lineSpacing = lhPx * 0.75 * scale;
    }
  }

  // --- Spacing (Margins) ---
  // Convert CSS margins (px) to PPTX Paragraph Spacing (pt).
  let paraSpaceBefore = 0;
  let paraSpaceAfter = 0;

  const mt = parseFloat(style.marginTop) || 0;
  const mb = parseFloat(style.marginBottom) || 0;

  if (mt > 0) paraSpaceBefore = mt * 0.75 * scale;
  if (mb > 0) paraSpaceAfter = mb * 0.75 * scale;

  // textDecoration is the legacy shorthand (e.g. "underline solid rgb(...)").
  // textDecorationLine is the modern split — read both so we work across
  // browsers and serialized styles.
  const decoration =
    (style.textDecorationLine || style.textDecoration || '').toString();
  const hasUnderline = decoration.includes('underline');
  const hasLineThrough = decoration.includes('line-through');

  return {
    color: colorObj.hex || '000000',
    fontFace: resolveFontFaceList(style.fontFamily),
    fontSize: Number((fontSizePx * 0.75 * scale).toFixed(1)),
    bold: parseInt(style.fontWeight) >= 600,
    italic: style.fontStyle === 'italic',
    underline: hasUnderline,
    ...(hasLineThrough && { strike: 'sngStrike' }),
    // Only add if we have a valid value
    ...(lineSpacing && { lineSpacing }),
    ...(paraSpaceBefore > 0 && { paraSpaceBefore }),
    ...(paraSpaceAfter > 0 && { paraSpaceAfter }),
    // Map background color to highlight if present
    ...(parseColor(style.backgroundColor).hex
      ? { highlight: parseColor(style.backgroundColor).hex }
      : {}),
    // Mapping letter-spacing to charSpacing
    ...(style.letterSpacing && style.letterSpacing !== 'normal'
      ? { charSpacing: parseFloat(style.letterSpacing) * 0.75 * scale }
      : {}),
  };
}

/**
 * Determines if a given DOM node is primarily a text container.
 * Updated to correctly reject Icon elements so they are rendered as images.
 */
export function isTextContainer(node) {
  const hasText = node.textContent.trim().length > 0;
  if (!hasText) return false;

  const children = Array.from(node.children);
  if (children.length === 0) return true;

  // When the parent is flex/grid with multiple children, those children are
  // blockified and laid out independently (e.g. `justify-content: space-between`
  // distributes them along the main axis). Computed `display` still reads as
  // "inline" for spans, but flattening them into a single inline-text run loses
  // their per-item positions. Force per-child render items in that case.
  const parentStyle = window.getComputedStyle(node);
  const parentIsFlexOrGrid = ['flex', 'inline-flex', 'grid', 'inline-grid'].includes(
    parentStyle.display
  );
  const parentWritingMode = parentStyle.writingMode || 'horizontal-tb';

  const isSafeInline = (el) => {
    // 1. Reject Web Components / Custom Elements
    if (el.tagName.includes('-')) return false;
    // 2. Reject Explicit Images/SVGs
    if (el.tagName === 'IMG' || el.tagName === 'SVG') return false;

    // 3. Reject Class-based Icons (FontAwesome, Material, Bootstrap, etc.)
    // If an <i> or <span> has icon classes, it is a visual object, not text.
    if (el.tagName === 'I' || el.tagName === 'SPAN') {
      const cls = el.getAttribute('class') || '';
      if (
        cls.includes('fa-') ||
        cls.includes('fas') ||
        cls.includes('far') ||
        cls.includes('fab') ||
        cls.includes('material-icons') ||
        cls.includes('bi-') ||
        cls.includes('icon')
      ) {
        return false;
      }
    }

    const style = window.getComputedStyle(el);
    const display = style.display;

    // Reject children that opt into their own visual orientation — vertical
    // writing-mode different from the parent, or a non-trivial transform
    // (rotate, etc.). These cannot be represented as a flat inline run inside
    // the parent's text flow; they need their own render item so we capture the
    // rotated bbox and emit the correct OOXML `vert` value.
    if ((style.writingMode || 'horizontal-tb') !== parentWritingMode) return false;
    if (style.transform && style.transform !== 'none') return false;

    // Multi-child flex/grid containers blockify their children — see comment
    // above isSafeInline.
    if (parentIsFlexOrGrid && children.length > 1) return false;

    // 4. Standard Inline Tag Check
    const isInlineTag = ['SPAN', 'B', 'STRONG', 'EM', 'I', 'A', 'SMALL', 'MARK'].includes(
      el.tagName
    );
    const isInlineDisplay = display.includes('inline');

    if (!isInlineTag && !isInlineDisplay) return false;

    // 5. Structural Styling Check
    // If a child has a background or border, it's a layout block, not a simple text span.
    const bgColor = parseColor(style.backgroundColor);
    const hasVisibleBg = bgColor.hex && bgColor.opacity > 0;
    const hasBorder =
      parseFloat(style.borderWidth) > 0 && parseColor(style.borderColor).opacity > 0;

    if (hasVisibleBg || hasBorder) {
      // Relaxed check: Allow inline elements with background/border to be treated as text.
      // They will be rendered as highlighted text runs (no border support in text runs though).
      // This preserves text flow for "badges".
      // return false;
    }

    // 4. Check for empty shapes (visual objects without text, like dots)
    const hasContent = el.textContent.trim().length > 0;
    if (!hasContent && (hasVisibleBg || hasBorder)) {
      return false;
    }

    return true;
  };

  return children.every(isSafeInline);
}

export function getRotation(transformStr) {
  return Math.round(decomposeMatrix2D(transformStr).rotation);
}

/**
 * Decomposes a CSS computed `transform` matrix into translate / rotate /
 * scale / skew components. `getComputedStyle` always returns the matrix form
 * (`matrix(...)` for 2D or `matrix3d(...)` for 3D-augmented), so this works
 * regardless of how the author wrote the transform.
 *
 * Decomposition: M = T(tx, ty) · R(θ) · [[sx, sx·tan(k)], [0, sy]]
 *
 * - rotation in degrees
 * - skewX in degrees — a non-zero value (>~0.5°) indicates the element is
 *   sheared, which PPTX shapes cannot represent natively. Callers should
 *   rasterize or accept the distortion.
 * - sx, sy: scale factors (post-rotation removal)
 *
 * Returns identity for `none` or unparseable inputs.
 */
export function decomposeMatrix2D(transformStr) {
  const id = { tx: 0, ty: 0, sx: 1, sy: 1, rotation: 0, skewX: 0 };
  if (!transformStr || transformStr === 'none') return id;
  let v = null;
  const m2 = transformStr.match(/^matrix\(([^)]+)\)$/);
  if (m2) {
    v = m2[1].split(',').map((s) => parseFloat(s.trim()));
    if (v.length < 6) return id;
  } else {
    const m3 = transformStr.match(/^matrix3d\(([^)]+)\)$/);
    if (!m3) return id;
    const all = m3[1].split(',').map((s) => parseFloat(s.trim()));
    if (all.length < 16) return id;
    // 2D submatrix of matrix3d: column-major, take [m11, m12, m21, m22, m41, m42].
    v = [all[0], all[1], all[4], all[5], all[12], all[13]];
  }
  const [a, b, c, d, e, f] = v;
  if (![a, b, c, d, e, f].every(isFinite)) return id;

  const rotationRad = Math.atan2(b, a);
  const sx = Math.hypot(a, b);
  // Remove rotation from (c, d): c' = cosθ·c + sinθ·d ; d' = -sinθ·c + cosθ·d
  const cosθ = Math.cos(rotationRad);
  const sinθ = Math.sin(rotationRad);
  const cRot = cosθ * c + sinθ * d;
  const dRot = -sinθ * c + cosθ * d;
  const sy = dRot;
  const skewXRad = sx > 1e-9 ? Math.atan2(cRot, sy) : 0;

  return {
    tx: e,
    ty: f,
    sx: isFinite(sx) && sx > 0 ? sx : 1,
    sy: isFinite(sy) && Math.abs(sy) > 1e-9 ? sy : 1,
    rotation: (rotationRad * 180) / Math.PI,
    skewX: (skewXRad * 180) / Math.PI,
  };
}

/** True if the element's computed transform contains a non-trivial skew. */
export function hasSkew(transformStr) {
  return Math.abs(decomposeMatrix2D(transformStr).skewX) > 0.5;
}

const horizMap = (v) => {
  if (v === 'center') return 'center';
  if (v === 'flex-end' || v === 'end') return 'right';
  if (v === 'flex-start' || v === 'start') return 'left';
  return null;
};
const vertMap = (v) => {
  if (v === 'center') return 'middle';
  if (v === 'flex-end' || v === 'end') return 'bottom';
  if (v === 'flex-start' || v === 'start') return 'top';
  return null;
};

/**
 * Resolve PPTX text alignment from a CSS layout style. Honors flex
 * (with `flex-direction` axis swap), grid, the `place-items` /
 * `place-content` shorthands, `text-align`, and the legacy
 * `line-height`-equals-height vertical centering trick.
 *
 * `intentionalSize` is true when the source box is sized larger than
 * the text on purpose (centered/end-aligned in flex/grid, or the
 * line-height vcenter pattern). When set, callers should disable
 * pptxgenjs `autoFit: true` (which emits `<a:spAutoFit/>` and resizes
 * the shape down to text) so the declared box dimensions survive.
 *
 * `lineHeightVcenter` is true when valign='middle' was inferred from
 * the legacy `line-height === height` trick. The matching `lineSpacing`
 * on the text runs (set by `getTextStyle` from `line-height`) would
 * otherwise stretch the line-box to the full shape height and push the
 * baseline to its bottom — callers should drop `lineSpacing` from each
 * run so PPTX's vertical centering measures the glyph box, not the
 * stretched line-box.
 */
export function resolveContainerAlignment(style, heightPx) {
  const display = style.display || '';
  const isFlex = display.includes('flex');
  const isGrid = display.includes('grid');

  const flexDirection = style.flexDirection || 'row';
  const isColumn = flexDirection === 'column' || flexDirection === 'column-reverse';

  let align = style.textAlign || 'left';
  if (align === 'start') align = 'left';
  if (align === 'end') align = 'right';
  if (!['left', 'right', 'center', 'justify'].includes(align)) align = 'left';

  let valign = 'top';
  let intentionalSize = false;
  let lineHeightVcenter = false;

  // computed-style getters resolve `place-items` / `place-content`
  // shorthands to their long-hand pairs, so reading align-items /
  // justify-items / align-content / justify-content is enough.
  const justifyContent = style.justifyContent || 'normal';
  const alignItems = style.alignItems || 'normal';
  const alignContent = style.alignContent || 'normal';
  const justifyItems = style.justifyItems || 'normal';

  if (isFlex) {
    // Main axis: justify-content. Cross axis: align-items. Swapped for column.
    const mainV = isColumn ? vertMap(justifyContent) : null;
    const mainH = !isColumn ? horizMap(justifyContent) : null;
    const crossV = !isColumn ? vertMap(alignItems) : null;
    const crossH = isColumn ? horizMap(alignItems) : null;
    if (mainH) { align = mainH; intentionalSize = true; }
    if (crossH) { align = crossH; intentionalSize = true; }
    if (mainV) { valign = mainV; intentionalSize = true; }
    if (crossV) { valign = crossV; intentionalSize = true; }
  } else if (isGrid) {
    // For a single-text grid item, `place-items` (item-axis) and
    // `place-content` (track-axis) both end up centering the run. Prefer
    // the item axis but fall back to the content axis when items are at
    // their `normal` / default.
    const h = horizMap(justifyItems) || horizMap(justifyContent);
    const v = vertMap(alignItems) || vertMap(alignContent);
    if (h) { align = h; intentionalSize = true; }
    if (v) { valign = v; intentionalSize = true; }
  }

  // Line-height-based vertical centering: a fixed-height bar with
  // `line-height` set to that height. Browser renders the text in the
  // middle of the line-box; PPTX uses font metrics so without an
  // explicit valign='middle' the run sits at the top of the shape.
  if (valign !== 'middle') {
    const fontSize = parseFloat(style.fontSize) || 0;
    const lineHeight = parseFloat(style.lineHeight) || 0; // 'normal' → NaN → 0
    if (
      fontSize > 0 &&
      lineHeight > fontSize * 1.4 &&
      heightPx > 0 &&
      Math.abs(lineHeight - heightPx) <= Math.max(2, heightPx * 0.05)
    ) {
      valign = 'middle';
      intentionalSize = true;
      lineHeightVcenter = true;
    }
  }

  return { align, valign, intentionalSize, lineHeightVcenter };
}

export function getWritingModeVert(writingMode, textOrientation) {
  const isUpright = textOrientation === 'upright';

  switch (writingMode) {
    case 'vertical-rl':
      // Latin in `vertical-rl` (default `mixed` orientation) rotates each
      // character 90° CW so the line reads top-to-bottom — that is OOXML
      // `vert`. `eaVert`/`wordArtVertRtl` stack characters upright, which is
      // correct for `text-orientation: upright`.
      return isUpright ? 'wordArtVertRtl' : 'vert';
    case 'vertical-lr':
      // Same character orientation as `vertical-rl`; only the multi-line
      // stacking direction differs. Single-line axis labels look identical so
      // map to `vert` as well.
      return isUpright ? 'wordArtVert' : 'vert';
    case 'sideways-rl':
      return 'vert';
    case 'sideways-lr':
      return 'vert270';
    default:
      return null;
  }
}

/**
 * The canonical CSS pattern for "axis label that reads bottom-to-top" is
 * `writing-mode: vertical-rl; transform: rotate(180deg)` (for left-side axis
 * labels). The combo flips a top-to-bottom vertical run into bottom-to-top.
 *
 * Returns `{ vert, rotation }` where the writing-mode `rotate(180deg)` has
 * been folded into the OOXML `vert` value, so the caller can apply zero
 * shape rotation. For other rotation values the original vert/rotation are
 * returned unchanged — those should be applied as a normal shape rotation
 * around the bounding box.
 */
export function combineVertWithRotation(vert, rotation) {
  if (!vert || !rotation) return { vert, rotation };
  if (Math.abs(rotation) !== 180) return { vert, rotation };

  // 180° flips the reading direction along the line axis.
  const flip = {
    vert: 'vert270',
    vert270: 'vert',
    eaVert: 'eaVert', // upright chars — 180° rotates the box, not the chars
    wordArtVert: 'wordArtVertRtl',
    wordArtVertRtl: 'wordArtVert',
  };
  return { vert: flip[vert] || vert, rotation: 0 };
}

/**
 * Converts an SVG node to a PNG data URL (rasterized)
 */
export function svgToPng(node) {
  return new Promise((resolve) => {
    const clone = node.cloneNode(true);
    const rect = node.getBoundingClientRect();
    const width = rect.width || 300;
    const height = rect.height || 150;

    inlineSvgStyles(node, clone);
    clone.setAttribute('width', width);
    clone.setAttribute('height', height);
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

    const xml = new XMLSerializer().serializeToString(clone);
    const svgUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(xml)}`;
    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const scale = 3;
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve(null);
    img.src = svgUrl;
  });
}

/**
 * Converts an SVG node to an SVG data URL (preserves vector format)
 * This allows "Convert to Shape" in PowerPoint
 */
export function svgToSvg(node) {
  return new Promise((resolve) => {
    try {
      const clone = node.cloneNode(true);
      const rect = node.getBoundingClientRect();
      const width = rect.width || 300;
      const height = rect.height || 150;

      inlineSvgStyles(node, clone);
      clone.setAttribute('width', width);
      clone.setAttribute('height', height);
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

      // Ensure xmlns:xlink is present for any xlink:href attributes
      if (clone.querySelector('[*|href]') || clone.innerHTML.includes('xlink:')) {
        clone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
      }

      const xml = new XMLSerializer().serializeToString(clone);
      // Use base64 encoding for better compatibility with PowerPoint
      const svgUrl = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(xml)))}`;
      resolve(svgUrl);
    } catch (e) {
      console.warn('SVG serialization failed:', e);
      resolve(null);
    }
  });
}

/**
 * Helper to inline computed styles into an SVG clone
 */
function inlineSvgStyles(source, target) {
  const computed = window.getComputedStyle(source);
  const properties = [
    'fill',
    'stroke',
    'stroke-width',
    'stroke-linecap',
    'stroke-linejoin',
    'opacity',
    'font-family',
    'font-size',
    'font-weight',
  ];

  if (computed.fill === 'none') target.setAttribute('fill', 'none');
  else if (computed.fill) target.style.fill = computed.fill;

  if (computed.stroke === 'none') target.setAttribute('stroke', 'none');
  else if (computed.stroke) target.style.stroke = computed.stroke;

  properties.forEach((prop) => {
    if (prop !== 'fill' && prop !== 'stroke') {
      const val = computed[prop];
      if (val && val !== 'auto') target.style[prop] = val;
    }
  });

  for (let i = 0; i < source.children.length; i++) {
    if (target.children[i]) inlineSvgStyles(source.children[i], target.children[i]);
  }
}

/**
 * Tokenize a single comma-separated box-shadow layer into typed parts.
 * Handles parenthesized color tokens (rgba(...), oklch(...)) as a unit.
 */
function tokenizeShadowLayer(layer) {
  const tokens = [];
  let buf = '';
  let depth = 0;
  for (let i = 0; i < layer.length; i++) {
    const ch = layer[i];
    if (ch === '(') {
      depth++;
      buf += ch;
    } else if (ch === ')') {
      depth--;
      buf += ch;
    } else if (/\s/.test(ch) && depth === 0) {
      if (buf) {
        tokens.push(buf);
        buf = '';
      }
    } else {
      buf += ch;
    }
  }
  if (buf) tokens.push(buf);
  return tokens;
}

/**
 * Parse a CSS box-shadow value into an array of layers. Each layer is
 * `{ inset, x, y, blur, spread, color, opacity }` with raw px values and
 * the color resolved through parseColor. Layers whose color resolves to
 * fully transparent are dropped — they have no visual effect.
 */
export function parseShadowList(shadowStr) {
  if (!shadowStr || shadowStr === 'none') return [];
  const layers = shadowStr.split(/,(?![^()]*\))/);
  const out = [];
  for (const raw of layers) {
    const trimmed = raw.trim();
    if (!trimmed) continue;
    const tokens = tokenizeShadowLayer(trimmed);
    let inset = false;
    const lengths = [];
    let colorToken = null;
    for (const tok of tokens) {
      if (tok === 'inset') {
        inset = true;
        continue;
      }
      // A length token: optional sign, digits, optional decimal, optional unit.
      // We only honor px (computed styles always normalize to px anyway).
      if (/^-?(\d*\.\d+|\d+)(px)?$/.test(tok)) {
        lengths.push(parseFloat(tok));
      } else {
        colorToken = tok;
      }
    }
    if (lengths.length < 2) continue;
    const [x, y, blur = 0, spread = 0] = lengths;
    const colorObj = colorToken ? parseColor(colorToken) : { hex: '000000', opacity: 1 };
    if (!colorObj.hex || colorObj.opacity === 0) continue;
    out.push({
      inset,
      x,
      y,
      blur,
      spread,
      color: colorObj.hex,
      opacity: colorObj.opacity,
    });
  }
  return out;
}

/**
 * Convert a single parsed shadow layer to the legacy PPTX-style shadow
 * option that pptxgenjs accepts on `shape.shadow`. Outer/inset is mapped
 * to type 'outer' / 'inner'.
 */
export function shadowLayerToPptx(layer, scale) {
  const distance = Math.sqrt(layer.x * layer.x + layer.y * layer.y);
  let angle = Math.atan2(layer.y, layer.x) * (180 / Math.PI);
  if (angle < 0) angle += 360;
  return {
    type: layer.inset ? 'inner' : 'outer',
    angle,
    blur: layer.blur * 0.75 * scale,
    offset: distance * 0.75 * scale,
    color: layer.color,
    opacity: layer.opacity,
  };
}

export function getVisibleShadow(shadowStr, scale) {
  const list = parseShadowList(shadowStr);
  if (!list.length) return null;
  return shadowLayerToPptx(list[0], scale);
}

/**
 * Compute per-corner radii in CSS px, resolving % to the smaller dimension.
 * Returns scalar radii (a single value per corner). Used by shadow-composite
 * helpers that don't render elliptical arcs — see `getCornerRadiiXY` for the
 * full elliptical form.
 */
export function getCornerRadii(style, widthPx, heightPx) {
  const minDim = Math.min(widthPx, heightPx);
  const parse = (v) => {
    if (!v) return 0;
    const s = String(v);
    // computed style for an elliptical corner returns "rx ry" (e.g. "80px 24px");
    // collapse to the larger axis so the scalar radius covers the visible curve.
    const tokens = s.trim().split(/\s+/);
    if (tokens.length === 2) {
      const a = parseR(tokens[0], minDim);
      const b = parseR(tokens[1], minDim);
      return Math.max(a, b);
    }
    return parseR(s, minDim);
  };
  return {
    tl: parse(style.borderTopLeftRadius),
    tr: parse(style.borderTopRightRadius),
    br: parse(style.borderBottomRightRadius),
    bl: parse(style.borderBottomLeftRadius),
  };
}

function parseR(s, basis) {
  if (!s) return 0;
  if (s.includes('%')) {
    const pct = parseFloat(s);
    return isFinite(pct) ? (pct / 100) * basis : 0;
  }
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

/**
 * Compute per-corner *elliptical* radii in CSS px. Each corner returns
 * `{ x, y }` where `x` is the horizontal radius (resolved against widthPx for
 * percentages) and `y` is the vertical radius (resolved against heightPx).
 * For uniform corners both axes match. The `border-radius: rx / ry` syntax
 * (e.g. `80px / 24px`) and per-corner long-hands are both handled — browsers
 * expand the shorthand into per-corner long-hands at computed-style time.
 *
 * Use this for emitting elliptical SVG arc paths; use the scalar
 * `getCornerRadii` for shadow halos that approximate with circular geometry.
 */
export function getCornerRadiiXY(style, widthPx, heightPx) {
  const parsePair = (v) => {
    if (!v) return { x: 0, y: 0 };
    const tokens = String(v).trim().split(/\s+/);
    const rx = parseR(tokens[0] || '0', widthPx);
    const ry = parseR(tokens[1] !== undefined ? tokens[1] : tokens[0] || '0', heightPx);
    return { x: rx, y: ry };
  };
  return {
    tl: parsePair(style.borderTopLeftRadius),
    tr: parsePair(style.borderTopRightRadius),
    br: parsePair(style.borderBottomRightRadius),
    bl: parsePair(style.borderBottomLeftRadius),
  };
}

/** True if any corner has a non-zero elliptical (rx≠ry) component. */
export function isElliptical(radiiXY) {
  for (const k of ['tl', 'tr', 'br', 'bl']) {
    const r = radiiXY[k];
    if (Math.abs(r.x - r.y) > 0.5) return true;
  }
  return false;
}

/** True if corners are not all equal (per-corner divergence). */
export function isPerCorner(radiiXY) {
  const t = radiiXY.tl;
  for (const k of ['tr', 'br', 'bl']) {
    const r = radiiXY[k];
    if (Math.abs(r.x - t.x) > 0.5 || Math.abs(r.y - t.y) > 0.5) return true;
  }
  return false;
}

function tracePath(ctx, x, y, w, h, radii) {
  const tl = Math.max(0, Math.min(radii.tl, w / 2, h / 2));
  const tr = Math.max(0, Math.min(radii.tr, w / 2, h / 2));
  const br = Math.max(0, Math.min(radii.br, w / 2, h / 2));
  const bl = Math.max(0, Math.min(radii.bl, w / 2, h / 2));
  ctx.moveTo(x + tl, y);
  ctx.lineTo(x + w - tr, y);
  if (tr > 0) ctx.quadraticCurveTo(x + w, y, x + w, y + tr);
  ctx.lineTo(x + w, y + h - br);
  if (br > 0) ctx.quadraticCurveTo(x + w, y + h, x + w - br, y + h);
  ctx.lineTo(x + bl, y + h);
  if (bl > 0) ctx.quadraticCurveTo(x, y + h, x, y + h - bl);
  ctx.lineTo(x, y + tl);
  if (tl > 0) ctx.quadraticCurveTo(x, y, x + tl, y);
}

function hexToRgbCss(hex, opacity) {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${opacity})`;
}

/**
 * Compose multiple OUTER box-shadow layers into a single PNG to be placed
 * behind the shape. Returns `{ dataUrl, paddingPx }` where paddingPx is the
 * margin around the shape's rect on each side (the image's outer rect
 * extends past the shape rect by paddingPx on every side).
 *
 * Implementation: render each shadow as a canvas drop-shadow on the rounded
 * rect, then subtract the shape area so only the shadow halo remains.
 */
export function composeOuterShadows(widthPx, heightPx, radii, shadows) {
  const list = (shadows || []).filter((s) => !s.inset);
  if (!list.length) return null;
  let pad = 0;
  for (const s of list) {
    pad = Math.max(pad, Math.abs(s.x) + s.blur + Math.max(0, s.spread) + 4);
    pad = Math.max(pad, Math.abs(s.y) + s.blur + Math.max(0, s.spread) + 4);
  }
  pad = Math.ceil(pad);
  const W = Math.ceil(widthPx + pad * 2);
  const H = Math.ceil(heightPx + pad * 2);
  if (typeof document === 'undefined') return null;
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  for (const s of list) {
    const sw = widthPx + s.spread * 2;
    const sh = heightPx + s.spread * 2;
    if (sw <= 0 || sh <= 0) continue;
    const sr = {
      tl: Math.max(0, radii.tl + s.spread),
      tr: Math.max(0, radii.tr + s.spread),
      br: Math.max(0, radii.br + s.spread),
      bl: Math.max(0, radii.bl + s.spread),
    };
    ctx.save();
    ctx.shadowColor = hexToRgbCss(s.color, s.opacity);
    // Canvas shadowBlur ≈ 2 × CSS blur stddev; CSS blur radius is roughly
    // 2 × stddev, so 1:1 is close enough for visual equivalence.
    ctx.shadowBlur = s.blur;
    ctx.shadowOffsetX = s.x;
    ctx.shadowOffsetY = s.y;
    ctx.fillStyle = '#000';
    ctx.beginPath();
    tracePath(ctx, pad - s.spread, pad - s.spread, sw, sh, sr);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
  }

  // Cut the shape region — the shape itself is drawn separately and would
  // otherwise overlap with this image's central solid fill (each layer's
  // unblurred core).
  ctx.save();
  ctx.globalCompositeOperation = 'destination-out';
  ctx.fillStyle = '#000';
  ctx.beginPath();
  tracePath(ctx, pad, pad, widthPx, heightPx, radii);
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  return { dataUrl: canvas.toDataURL('image/png'), paddingPx: pad };
}

/**
 * Compose multiple INSET box-shadow layers into a single PNG sized exactly
 * to the shape rect. The image is intended to be overlaid on top of the
 * shape (inner shadows render above the shape's fill / border).
 *
 * Implementation: clip to the rounded shape, then fill an "outside the
 * shape" region with the shadow color. Because the fill is offset and
 * blurred via canvas shadow*, only the shadow halo bleeds into the
 * clipped interior.
 */
export function composeInsetShadows(widthPx, heightPx, radii, shadows) {
  const list = (shadows || []).filter((s) => s.inset);
  if (!list.length) return null;
  if (typeof document === 'undefined') return null;
  // Render with a margin so the "outer rim" of the inverse-shape path is
  // actually rasterized inside the canvas — otherwise the browser would
  // shadow-blur only the rasterized portion, which for a fits-the-canvas
  // shape is just the four rounded-corner cutouts. With margin, the rim is
  // a thick belt around the shape and casts a proper inset shadow.
  let M = 64;
  for (const s of list) {
    M = Math.max(M, Math.ceil(Math.abs(s.x) + s.blur + Math.abs(s.spread) + 16));
  }
  const W = Math.ceil(widthPx) + M * 2;
  const H = Math.ceil(heightPx) + M * 2;
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  for (const s of list) {
    const ix = M + s.spread;
    const iy = M + s.spread;
    const iw = widthPx - s.spread * 2;
    const ih = heightPx - s.spread * 2;
    if (iw <= 0 || ih <= 0) continue;
    const ir = {
      tl: Math.max(0, radii.tl - s.spread),
      tr: Math.max(0, radii.tr - s.spread),
      br: Math.max(0, radii.br - s.spread),
      bl: Math.max(0, radii.bl - s.spread),
    };
    ctx.save();
    ctx.shadowColor = hexToRgbCss(s.color, s.opacity);
    ctx.shadowBlur = s.blur;
    ctx.shadowOffsetX = s.x;
    ctx.shadowOffsetY = s.y;
    // Opaque fill: canvas multiplies the source pixel alpha into the shadow
    // alpha. With a translucent fill the shadow ends up doubly-attenuated.
    ctx.fillStyle = '#000';

    // Even-odd: outer rect (≈ canvas) MINUS the spread-shrunk inner shape
    // leaves a "frame" of fill around the shape. The frame's inner edge
    // traces the shape boundary; canvas computes shadow from that edge,
    // offset and blurred. We mask to the shape afterwards.
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(W, 0);
    ctx.lineTo(W, H);
    ctx.lineTo(0, H);
    ctx.closePath();
    tracePath(ctx, ix, iy, iw, ih, ir);
    ctx.closePath();
    ctx.fill('evenodd');
    ctx.restore();
  }

  // Mask to the shape only — everything outside the shape (the frame fill
  // and any shadow that bled into the margin) is discarded.
  ctx.save();
  ctx.globalCompositeOperation = 'destination-in';
  ctx.fillStyle = '#000';
  ctx.beginPath();
  tracePath(ctx, M, M, widthPx, heightPx, radii);
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  // Crop to shape size for the output image. The cropped PNG aligns to the
  // shape's PPTX rect with no padding.
  const out = document.createElement('canvas');
  out.width = Math.ceil(widthPx);
  out.height = Math.ceil(heightPx);
  const octx = out.getContext('2d');
  if (!octx) return null;
  octx.drawImage(canvas, M, M, Math.ceil(widthPx), Math.ceil(heightPx), 0, 0, out.width, out.height);
  return { dataUrl: out.toDataURL('image/png'), paddingPx: 0 };
}

/**
 * Generates an SVG image for gradients, supporting degrees and keywords.
 */
export function generateGradientSVG(w, h, bgString, radius, border) {
  try {
    const match = bgString.match(/linear-gradient\((.*)\)/);
    if (!match) return null;
    const content = match[1];

    // Split by comma, ignoring commas inside parentheses (e.g. rgba())
    const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
    if (parts.length < 2) return null;

    let x1 = '0%',
      y1 = '0%',
      x2 = '0%',
      y2 = '100%';
    let stopsStartIndex = 0;
    const firstPart = parts[0].toLowerCase();

    // 1. Check for Keywords (to right, etc.)
    if (firstPart.startsWith('to ')) {
      stopsStartIndex = 1;
      const direction = firstPart.replace('to ', '').trim();
      switch (direction) {
        case 'top':
          y1 = '100%';
          y2 = '0%';
          break;
        case 'bottom':
          y1 = '0%';
          y2 = '100%';
          break;
        case 'left':
          x1 = '100%';
          x2 = '0%';
          break;
        case 'right':
          x2 = '100%';
          break;
        case 'top right':
          x1 = '0%';
          y1 = '100%';
          x2 = '100%';
          y2 = '0%';
          break;
        case 'top left':
          x1 = '100%';
          y1 = '100%';
          x2 = '0%';
          y2 = '0%';
          break;
        case 'bottom right':
          x2 = '100%';
          y2 = '100%';
          break;
        case 'bottom left':
          x1 = '100%';
          y2 = '100%';
          break;
      }
    }
    // 2. Check for Degrees (45deg, 90deg, etc.)
    else if (firstPart.match(/^-?[\d.]+(deg|rad|turn|grad)$/)) {
      stopsStartIndex = 1;
      const val = parseFloat(firstPart);
      // CSS 0deg is Top (North), 90deg is Right (East), 180deg is Bottom (South)
      // We convert this to SVG coordinates on a unit square (0-100%).
      // Formula: Map angle to perimeter coordinates.
      if (!isNaN(val)) {
        const deg = firstPart.includes('rad') ? val * (180 / Math.PI) : val;
        const cssRad = ((deg - 90) * Math.PI) / 180; // Correct CSS angle offset

        // Calculate standard vector for rectangle center (50, 50)
        const scale = 50; // Distance from center to edge (approx)
        const cos = Math.cos(cssRad); // Y component (reversed in SVG)
        const sin = Math.sin(cssRad); // X component

        // Invert Y for SVG coordinate system
        x1 = (50 - sin * scale).toFixed(1) + '%';
        y1 = (50 + cos * scale).toFixed(1) + '%';
        x2 = (50 + sin * scale).toFixed(1) + '%';
        y2 = (50 - cos * scale).toFixed(1) + '%';
      }
    }

    // 3. Process Color Stops
    let stopsXML = '';
    const stopParts = parts.slice(stopsStartIndex);

    stopParts.forEach((part, idx) => {
      // Parse "Color Position" (e.g., "red 50%")
      // Regex looks for optional space + number + unit at the end of the string
      let color = part;
      let offset = Math.round((idx / (stopParts.length - 1)) * 100) + '%';

      const posMatch = part.match(/^(.*?)\s+(-?[\d.]+(?:%|px)?)$/);
      if (posMatch) {
        color = posMatch[1];
        offset = posMatch[2];
      }

      // Handle RGBA/RGB for SVG compatibility
      let opacity = 1;
      if (color.includes('rgba')) {
        const rgbaMatch = color.match(/[\d.]+/g);
        if (rgbaMatch && rgbaMatch.length >= 4) {
          opacity = rgbaMatch[3];
          color = `rgb(${rgbaMatch[0]},${rgbaMatch[1]},${rgbaMatch[2]})`;
        }
      }

      stopsXML += `<stop offset="${offset}" stop-color="${color.trim()}" stop-opacity="${opacity}"/>`;
    });

    let strokeAttr = '';
    if (border) {
      strokeAttr = `stroke="#${border.color}" stroke-width="${border.width}"`;
    }

    // Accept scalar (number), per-corner scalars (`{tl, tr, br, bl}`), or
    // elliptical per-corner (`{tl: {x, y}, ...}`).
    let radiiInput;
    if (typeof radius === 'object' && radius !== null) {
      radiiInput = radius;
    } else {
      const n = radius || 0;
      radiiInput = { tl: n, tr: n, br: n, bl: n };
    }
    const r = normalizeRadiiXY(radiiInput);
    clampRadiiXY(r, w, h);
    const pathD = buildEllipticalPath(w, h, r);

    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
          <defs>
            <linearGradient id="grad" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">
              ${stopsXML}
            </linearGradient>
          </defs>
          <path d="${pathD}" fill="url(#grad)" ${strokeAttr} />
      </svg>`;

    return 'data:image/svg+xml;base64,' + btoa(svg);
  } catch (e) {
    console.warn('Gradient generation failed:', e);
    return null;
  }
}

export function generateBlurredSVG(w, h, color, radius, blurPx) {
  const padding = blurPx * 3;
  const fullW = w + padding * 2;
  const fullH = h + padding * 2;
  const x = padding;
  const y = padding;
  let shapeTag = '';
  const isCircle = radius >= Math.min(w, h) / 2 - 1 && Math.abs(w - h) < 2;

  if (isCircle) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    shapeTag = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="#${color}" filter="url(#f1)" />`;
  } else {
    shapeTag = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${radius}" ry="${radius}" fill="#${color}" filter="url(#f1)" />`;
  }

  const svg = `
  <svg xmlns="http://www.w3.org/2000/svg" width="${fullW}" height="${fullH}" viewBox="0 0 ${fullW} ${fullH}">
    <defs>
      <filter id="f1" x="-50%" y="-50%" width="200%" height="200%">
        <feGaussianBlur in="SourceGraphic" stdDeviation="${blurPx}" />
      </filter>
    </defs>
    ${shapeTag}
  </svg>`;

  return {
    data: 'data:image/svg+xml;base64,' + btoa(svg),
    padding: padding,
  };
}

// src/utils.js

// ... (keep all existing exports) ...

/**
 * Traverses the target DOM and collects all unique font-family names used.
 */
export function getUsedFontFamilies(root) {
  const families = new Set();

  function scan(node) {
    if (node.nodeType === 1) {
      // Element
      const style = window.getComputedStyle(node);
      const fontList = style.fontFamily.split(',');
      // The first font in the stack is the primary one
      const primary = fontList[0].trim().replace(/['"]/g, '');
      if (primary) families.add(primary);
    }
    for (const child of node.childNodes) {
      scan(child);
    }
  }

  // Handle array of roots or single root
  const elements = Array.isArray(root) ? root : [root];
  elements.forEach((el) => {
    const node = typeof el === 'string' ? document.querySelector(el) : el;
    if (node) scan(node);
  });

  return families;
}

// Lightweight @font-face parser used as a fallback when document.styleSheets
// can't expose cssRules due to CORS. We only need family + src; we leave
// weight / style alone because pptxgenjs ignores them anyway.
function parseFontFaceRules(cssText, baseHref) {
  const rules = [];
  const re = /@font-face\s*\{([^}]+)\}/gi;
  let m;
  while ((m = re.exec(cssText))) {
    const body = m[1];
    const familyMatch = body.match(/font-family\s*:\s*(?:["']([^"']+)["']|([^;]+));/i);
    const srcMatch = body.match(/src\s*:\s*([^;]+);/i);
    if (!familyMatch || !srcMatch) continue;
    const family = (familyMatch[1] || familyMatch[2] || '').trim().replace(/^['"]|['"]$/g, '');
    let src = srcMatch[1];
    // Resolve any relative url(...) against the stylesheet URL so that the
    // resulting absolute URL works when fetched cross-origin.
    if (baseHref) {
      try {
        src = src.replace(/url\((['"]?)([^)'"]+)\1\)/g, (_, q, u) => {
          if (/^(https?:|data:)/i.test(u)) return `url(${q}${u}${q})`;
          try {
            return `url(${q}${new URL(u, baseHref).href}${q})`;
          } catch {
            return `url(${q}${u}${q})`;
          }
        });
      } catch {
        /* leave src untouched */
      }
    }
    rules.push({ family, src });
  }
  return rules;
}

/**
 * Scans document.styleSheets to find @font-face URLs for the requested families.
 * Returns an array of { name, url } objects.
 */
export async function getAutoDetectedFonts(usedFamilies) {
  const foundFonts = [];
  const processedUrls = new Set();

  // Helper to extract clean URL from CSS src string
  const extractUrl = (srcStr) => {
    // Look for url("...") or url('...') or url(...)
    // Prioritize woff, ttf, otf. Avoid woff2 if possible as handling is harder,
    // but if it's the only one, take it (convert logic handles it best effort).
    const matches = srcStr.match(/url\((['"]?)(.*?)\1\)/g);
    if (!matches) return null;

    // Filter for preferred formats
    let chosenUrl = null;
    for (const match of matches) {
      const urlRaw = match.replace(/url\((['"]?)(.*?)\1\)/, '$2');
      // Skip data URIs for now (unless you want to support base64 embedding)
      if (urlRaw.startsWith('data:')) continue;

      if (urlRaw.includes('.ttf') || urlRaw.includes('.otf') || urlRaw.includes('.woff')) {
        chosenUrl = urlRaw;
        break; // Found a good one
      }
      // Fallback
      if (!chosenUrl) chosenUrl = urlRaw;
    }
    return chosenUrl;
  };

  // Sheets we couldn't read via CSSOM. We re-fetch them as text and parse
  // @font-face blocks ourselves; this rescues icon fonts hosted on CDNs that
  // serve the font with permissive CORS but the *stylesheet* without it.
  const blockedSheetHrefs = [];

  for (const sheet of Array.from(document.styleSheets)) {
    try {
      // Accessing cssRules on cross-origin sheets (like Google Fonts) might fail
      // if CORS headers aren't set. We wrap in try/catch.
      const rules = sheet.cssRules || sheet.rules;
      if (!rules) continue;

      for (const rule of Array.from(rules)) {
        if (rule.constructor.name === 'CSSFontFaceRule' || rule.type === 5) {
          const familyName = rule.style.getPropertyValue('font-family').replace(/['"]/g, '').trim();

          if (usedFamilies.has(familyName)) {
            const src = rule.style.getPropertyValue('src');
            const url = extractUrl(src);

            if (url && !processedUrls.has(url)) {
              processedUrls.add(url);
              foundFonts.push({ name: familyName, url: url });
            }
          }
        }
      }
    } catch {
      // SecurityError is common for external stylesheets without CORS.
      // We retry these by fetching the raw .css and parsing it ourselves.
      if (sheet.href) blockedSheetHrefs.push(sheet.href);
    }
  }

  for (const href of blockedSheetHrefs) {
    try {
      const cssText = await (await fetch(href)).text();
      const fontFaces = parseFontFaceRules(cssText, href);
      for (const ff of fontFaces) {
        if (!usedFamilies.has(ff.family)) continue;
        const url = extractUrl(ff.src);
        if (!url || processedUrls.has(url)) continue;
        processedUrls.add(url);
        foundFonts.push({ name: ff.family, url });
      }
    } catch (e) {
      console.warn(`Cannot fetch stylesheet for font auto-detection: ${href}`, e);
    }
  }

  return foundFonts;
}

export function collectTextParts(node, parentStyle, scale) {
  const parts = [];

  // Skip subtrees that the browser is rendering as fully invisible. Reveal.js
  // fragments, hidden tabs, etc. compute to opacity:0 / visibility:hidden but
  // still live in the DOM — without this guard the text container path emits
  // them as visible runs.
  if (node.nodeType === 1) {
    const cs = window.getComputedStyle(node);
    if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') {
      return parts;
    }
  }

  // Check for CSS Content (::before) - often used for icons
  if (node.nodeType === 1) {
    const beforeStyle = window.getComputedStyle(node, '::before');
    const content = beforeStyle.content;
    if (content && content !== 'none' && content !== 'normal' && content !== '""') {
      // Strip quotes
      const cleanContent = content.replace(/^['"]|['"]$/g, '');
      if (cleanContent.trim()) {
        parts.push({
          text: sanitizeText(cleanContent + ' '), // Add space after icon
          options: getTextStyle(window.getComputedStyle(node), scale),
        });
      }
    }
  }

  let trimNextLeading = false;

  node.childNodes.forEach((child, index) => {
    if (child.nodeType === 3) {
      // Text. Use parent style for whitespace handling (text nodes inherit it).
      const styleToUse = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
      const ws = styleToUse.whiteSpace;
      const wsProcessed = processWhitespace(child.nodeValue, ws);
      const segments = Array.isArray(wsProcessed)
        ? wsProcessed
        : [{ text: wsProcessed }];
      // `pre`, `pre-wrap`, and `break-spaces` preserve significant whitespace
      // — including leading runs that processWhitespace would otherwise return
      // as a plain string (when no newline forces array form). Only the truly
      // collapsing modes (`normal`, `nowrap`) should trim leading/trailing.
      const isCollapsed = ws !== 'pre' && ws !== 'pre-wrap' && ws !== 'break-spaces' && ws !== 'pre-line';

      segments.forEach((seg, segIdx) => {
        if (seg.breakLine) {
          parts.push({ text: '', options: { breakLine: true } });
          trimNextLeading = true;
          return;
        }
        let val = seg.text || '';

        if (isCollapsed) {
          if (index === 0 && segIdx === 0) val = val.trimStart();
          if (trimNextLeading) {
            val = val.trimStart();
            trimNextLeading = false;
          }
          if (
            index === node.childNodes.length - 1 &&
            segIdx === segments.length - 1
          ) {
            val = val.trimEnd();
          }
        } else if (trimNextLeading) {
          val = val.replace(/^[ \t]+/, '');
          trimNextLeading = false;
        }

        if (val) {
          val = applyTextTransform(val, styleToUse.textTransform);
          parts.push({
            text: sanitizeText(val),
            options: getTextStyle(styleToUse, scale),
          });
        }
      });
    } else if (child.nodeType === 1) {
      if (child.tagName === 'BR') {
        if (parts.length > 0) {
          const lastPart = parts[parts.length - 1];
          if (lastPart.text && typeof lastPart.text === 'string') {
            lastPart.text = lastPart.text.trimEnd();
          }
        }
        parts.push({ text: '', options: { breakLine: true } });
        trimNextLeading = true;
      } else {
        const isBlock = ['DIV', 'P', 'LI'].includes(child.tagName);
        if (isBlock && parts.length > 0 && !parts[parts.length - 1].options?.breakLine) {
          parts.push({ text: '', options: { breakLine: true } });
        }

        const childParts = collectTextParts(child, parentStyle, scale);
        if (childParts.length > 0) parts.push(...childParts);

        if (isBlock) {
          parts.push({ text: '', options: { breakLine: true } });
          trimNextLeading = true;
        }
      }
    }
  });

  // Cleanup potential trailing empty breakLines
  while (parts.length > 0 && parts[parts.length - 1].options?.breakLine && parts[parts.length - 1].text === '') {
    parts.pop();
  }

  return parts;
}
