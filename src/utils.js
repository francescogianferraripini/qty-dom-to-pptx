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

// Splits a CSS font-family value, honoring quoted segments, and resolves
// generic keywords to a concrete fallback PowerPoint will display. Returns a
// comma-separated string suitable for pptxgenjs `fontFace`.
//
// Important: pptxgenjs concatenates the returned string straight into an XML
// attribute value (e.g. `<a:latin typeface="...">`) without entity-escaping,
// so we MUST strip every embedded `"` here or the resulting OOXML is invalid
// and PowerPoint/LibreOffice will fail to render the slide. We also drop
// stray single quotes for symmetry.
export function resolveFontFaceList(fontFamilyStr) {
  if (!fontFamilyStr) return 'Calibri';
  // Tokenize on commas at depth 0. We then strip outer quotes per-token.
  const rawTokens = fontFamilyStr.split(',');
  const out = [];
  const seen = new Set();
  for (let token of rawTokens) {
    let name = token.trim();
    // Strip a single layer of matching outer quotes if present.
    if (
      (name.startsWith('"') && name.endsWith('"')) ||
      (name.startsWith("'") && name.endsWith("'"))
    ) {
      name = name.slice(1, -1).trim();
    }
    // Defensive: even if a name still contains stray quotes, remove them so
    // the attribute value can never break the XML.
    name = name.replace(/["']/g, '');
    if (!name) continue;
    const generic = GENERIC_FONT_MAP[name.toLowerCase()];
    if (generic) name = generic;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(name);
  }
  return out.length ? out.join(', ') : 'Calibri';
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
  if (firstRow) {
    const cells = Array.from(firstRow.children);
    cells.forEach((cell) => {
      const rect = cell.getBoundingClientRect();
      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const wIn = (rect.width * (1 / 96) * scale) / colspan;
      for (let i = 0; i < colspan; i++) {
        colWidths.push(wIn);
      }
    });
  }

  const tableStyle = window.getComputedStyle(node);
  const borderSpacing = tableStyle.borderSpacing.split(' ');
  const hSpace = parseFloat(borderSpacing[0]) || 0;
  const vSpace = parseFloat(borderSpacing[1] || borderSpacing[0]) || 0;
  const hSpacePt = hSpace * 0.75 * styleScale;
  const vSpacePt = vSpace * 0.75 * styleScale;

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

      // D. Padding (Margins in PPTX)
      // CSS Padding px -> PPTX Margin pt (logical px → styleScale)
      const padding = getPadding(style, styleScale);
      // getPadding returns [top, right, bottom, left] in inches relative to scale
      // PptxGenJS expects points (pt) for margin: [t, r, b, l]
      // or discrete properties. Let's use discrete for clarity.
      const margin = [
        padding[0] * 72 + vSpacePt / 2, // top
        padding[1] * 72 + hSpacePt / 2, // right
        padding[2] * 72 + vSpacePt / 2, // bottom
        padding[3] * 72 + hSpacePt / 2, // left
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
  const top = {
    width: parseFloat(style.borderTopWidth) || 0,
    style: style.borderTopStyle,
    color: parseColor(style.borderTopColor).hex,
  };
  const right = {
    width: parseFloat(style.borderRightWidth) || 0,
    style: style.borderRightStyle,
    color: parseColor(style.borderRightColor).hex,
  };
  const bottom = {
    width: parseFloat(style.borderBottomWidth) || 0,
    style: style.borderBottomStyle,
    color: parseColor(style.borderBottomColor).hex,
  };
  const left = {
    width: parseFloat(style.borderLeftWidth) || 0,
    style: style.borderLeftStyle,
    color: parseColor(style.borderLeftColor).hex,
  };

  const hasAnyBorder = top.width > 0 || right.width > 0 || bottom.width > 0 || left.width > 0;
  if (!hasAnyBorder) return { type: 'none' };

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
        transparency: (1 - parseColor(style.borderTopColor).opacity) * 100,
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
 */
export function generateCustomShapeSVG(w, h, color, opacity, radii) {
  let { tl, tr, br, bl } = radii;

  // Clamp radii using CSS spec logic (avoid overlap)
  const factor = Math.min(
    w / (tl + tr) || Infinity,
    h / (tr + br) || Infinity,
    w / (br + bl) || Infinity,
    h / (bl + tl) || Infinity
  );

  if (factor < 1) {
    tl *= factor;
    tr *= factor;
    br *= factor;
    bl *= factor;
  }

  const path = `
    M ${tl} 0
    L ${w - tr} 0
    A ${tr} ${tr} 0 0 1 ${w} ${tr}
    L ${w} ${h - br}
    A ${br} ${br} 0 0 1 ${w - br} ${h}
    L ${bl} ${h}
    A ${bl} ${bl} 0 0 1 0 ${h - bl}
    L 0 ${tl}
    A ${tl} ${tl} 0 0 1 ${tl} 0
    Z
  `;

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
      <path d="${path}" fill="#${color}" fill-opacity="${opacity}" />
    </svg>`;

  return 'data:image/svg+xml;base64,' + btoa(svg);
}

// --- REPLACE THE EXISTING parseColor FUNCTION ---
export function parseColor(str) {
  if (!str || str === 'transparent' || str.trim() === 'rgba(0, 0, 0, 0)') {
    return { hex: null, opacity: 0 };
  }

  const ctx = getCtx();
  ctx.fillStyle = str;
  const computed = ctx.fillStyle;

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
  // Use Canvas API to convert to sRGB
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
  if (!transformStr || transformStr === 'none') return 0;
  const values = transformStr.split('(')[1].split(')')[0].split(',');
  if (values.length < 4) return 0;
  const a = parseFloat(values[0]);
  const b = parseFloat(values[1]);
  return Math.round(Math.atan2(b, a) * (180 / Math.PI));
}

export function getWritingModeVert(writingMode, textOrientation) {
  const isUpright = textOrientation === 'upright';

  switch (writingMode) {
    case 'vertical-rl':
      return isUpright ? 'wordArtVertRtl' : 'eaVert';
    case 'vertical-lr':
      return isUpright ? 'wordArtVert' : 'mongolianVert';
    case 'sideways-rl':
      return 'vert';
    case 'sideways-lr':
      return 'vert270';
    default:
      return null;
  }
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

export function getVisibleShadow(shadowStr, scale) {
  if (!shadowStr || shadowStr === 'none') return null;
  const shadows = shadowStr.split(/,(?![^()]*\))/);
  for (let s of shadows) {
    s = s.trim();
    if (s.startsWith('rgba(0, 0, 0, 0)')) continue;
    const match = s.match(
      /(rgba?\([^)]+\)|#[0-9a-fA-F]+)\s+(-?[\d.]+)px\s+(-?[\d.]+)px\s+([\d.]+)px/
    );
    if (match) {
      const colorStr = match[1];
      const x = parseFloat(match[2]);
      const y = parseFloat(match[3]);
      const blur = parseFloat(match[4]);
      const distance = Math.sqrt(x * x + y * y);
      let angle = Math.atan2(y, x) * (180 / Math.PI);
      if (angle < 0) angle += 360;
      const colorObj = parseColor(colorStr);
      return {
        type: 'outer',
        angle: angle,
        blur: blur * 0.75 * scale,
        offset: distance * 0.75 * scale,
        color: colorObj.hex || '000000',
        opacity: colorObj.opacity,
      };
    }
  }
  return null;
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

    let tl = 0, tr = 0, br = 0, bl = 0;
    if (typeof radius === 'object' && radius !== null) {
      tl = radius.tl || 0;
      tr = radius.tr || 0;
      br = radius.br || 0;
      bl = radius.bl || 0;
    } else {
      tl = tr = br = bl = radius || 0;
    }

    const factor = Math.min(
      w / (tl + tr) || Infinity,
      h / (tr + br) || Infinity,
      w / (br + bl) || Infinity,
      h / (bl + tl) || Infinity
    );

    if (factor < 1) {
      tl *= factor; tr *= factor; br *= factor; bl *= factor;
    }

    // Generate absolute path based on radius bounds
    const pathD = `M ${tl} 0 L ${w - tr} 0 A ${tr} ${tr} 0 0 1 ${w} ${tr} L ${w} ${h - br} A ${br} ${br} 0 0 1 ${w - br} ${h} L ${bl} ${h} A ${bl} ${bl} 0 0 1 0 ${h - bl} L 0 ${tl} A ${tl} ${tl} 0 0 1 ${tl} 0 Z`;

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
      const isCollapsed = !Array.isArray(wsProcessed);

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
