import { describe, it, expect } from 'vitest';
import {
  generateGradientSVG,
  generateCompositeBorderSVG,
  generateCustomShapeSVG,
} from '../../src/utils.js';

/**
 * The generators return `data:image/svg+xml;base64,...` URLs. Snapshot the
 * decoded SVG so diffs stay readable, and normalize random IDs so the
 * snapshot is deterministic.
 */
function decode(dataUrl) {
  if (!dataUrl) return null;
  const match = dataUrl.match(/^data:image\/svg\+xml;base64,(.*)$/);
  if (!match) throw new Error(`Not an svg+xml data URL: ${dataUrl.slice(0, 60)}`);
  let svg = atob(match[1]);
  // Strip random ids so snapshots don't churn.
  svg = svg.replace(/clip_[a-z0-9]+/g, 'clip_X');
  // Normalize whitespace (the source uses indented multi-line template strings).
  svg = svg.replace(/\s+/g, ' ').trim();
  return svg;
}

describe('generateGradientSVG', () => {
  it('renders a top-to-bottom two-stop linear gradient', () => {
    const svg = decode(
      generateGradientSVG(100, 50, 'linear-gradient(#ff0000, #0000ff)', 0, null),
    );
    expect(svg).toMatchSnapshot();
  });

  it('honors the "to right" keyword', () => {
    const svg = decode(
      generateGradientSVG(
        200,
        100,
        'linear-gradient(to right, #ffffff 0%, #000000 100%)',
        0,
        null,
      ),
    );
    expect(svg).toMatchSnapshot();
  });

  it('honors degree angles', () => {
    const svg = decode(
      generateGradientSVG(
        200,
        100,
        'linear-gradient(135deg, #ff7e5f 0%, #feb47b 50%, #86a8e7 100%)',
        0,
        null,
      ),
    );
    expect(svg).toMatchSnapshot();
  });

  it('returns null on a non-gradient string', () => {
    expect(generateGradientSVG(100, 100, '#ff0000', 0, null)).toBeNull();
  });
});

describe('generateCompositeBorderSVG', () => {
  it('renders all four sides with different colors and widths', () => {
    const svg = decode(
      generateCompositeBorderSVG(200, 100, 8, {
        top: { width: 2, color: 'ff0000' },
        right: { width: 4, color: '00ff00' },
        bottom: { width: 6, color: '0000ff' },
        left: { width: 8, color: 'ffff00' },
      }),
    );
    expect(svg).toMatchSnapshot();
  });

  it('omits zero-width sides', () => {
    const svg = decode(
      generateCompositeBorderSVG(100, 100, 0, {
        top: { width: 2, color: '000000' },
        right: { width: 0, color: null },
        bottom: { width: 0, color: null },
        left: { width: 0, color: null },
      }),
    );
    expect(svg).toMatchSnapshot();
  });
});

describe('generateCustomShapeSVG', () => {
  it('renders uniform corner radii', () => {
    const svg = decode(
      generateCustomShapeSVG(200, 100, 'ff5722', 1, {
        tl: 16,
        tr: 16,
        br: 16,
        bl: 16,
      }),
    );
    expect(svg).toMatchSnapshot();
  });

  it('renders mixed corner radii', () => {
    const svg = decode(
      generateCustomShapeSVG(200, 100, '4285f4', 0.8, {
        tl: 32,
        tr: 8,
        br: 32,
        bl: 8,
      }),
    );
    expect(svg).toMatchSnapshot();
  });

  it('clamps radii that would otherwise overlap', () => {
    const svg = decode(
      generateCustomShapeSVG(40, 40, '000000', 1, {
        tl: 100,
        tr: 100,
        br: 100,
        bl: 100,
      }),
    );
    expect(svg).toMatchSnapshot();
  });
});
