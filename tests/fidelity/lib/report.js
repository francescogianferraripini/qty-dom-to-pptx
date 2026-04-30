import { writeFile } from 'node:fs/promises';
import path from 'node:path';

/**
 * results: Array<{
 *   name: string,
 *   sourcePng: string, // relative path
 *   pptxPng: string,
 *   diffPng: string,
 *   percent: number,          // raw pixel-percent
 *   contentPercent: number,   // block-level content-presence delta (gates budget)
 *   budget: number,
 *   mismatched: number,
 *   total: number,
 *   error?: string,
 * }>
 */
export async function writeReport(reportPath, results) {
  const rows = results
    .map((r) => {
      const cp = r.contentPercent ?? 0;
      const budget = r.budget ?? Infinity;
      const cls = r.error
        ? 'err'
        : cp > budget
          ? 'bad'
          : cp <= budget * 0.4
            ? 'good'
            : 'ok';
      const stat = r.error
        ? `<span class="err">ERROR: ${escapeHtml(r.error)}</span>`
        : `<strong>${cp.toFixed(2)}%</strong> content` +
          `<br><span class="raw">edge ${(r.edgePercent ?? 0).toFixed(1)}% · color ${(r.colorPercent ?? 0).toFixed(1)}%</span>` +
          `<br><span class="raw">${r.percent.toFixed(2)}% raw (${r.mismatched}/${r.total})</span>` +
          `<br><span class="raw">budget ${budget.toFixed(0)}%</span>`;
      const liveLink = r.casePath
        ? `<a href="${escapeHtml(r.casePath)}" target="_blank">live</a>`
        : '';
      const pptxLink = r.pptxPath
        ? `<a href="${escapeHtml(r.pptxPath)}">pptx</a>`
        : '';
      const links =
        liveLink || pptxLink
          ? `<div class="links">${[liveLink, pptxLink].filter(Boolean).join(' · ')}</div>`
          : '';
      return `
      <tr class="${cls}">
        <td class="name">${escapeHtml(r.name)}${links}</td>
        <td class="stat">${stat}</td>
        <td>${r.sourcePng ? `<img src="${escapeHtml(r.sourcePng)}" />` : ''}</td>
        <td>${r.pptxPng ? `<img src="${escapeHtml(r.pptxPng)}" />` : ''}</td>
        <td>${r.diffPng ? `<img src="${escapeHtml(r.diffPng)}" />` : ''}</td>
      </tr>`;
    })
    .join('\n');

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>dom-to-pptx fidelity report</title>
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 24px; background: #fafafa; }
  h1 { margin: 0 0 12px; }
  .meta { color: #666; margin-bottom: 16px; font-size: 13px; }
  table { border-collapse: collapse; width: 100%; background: white; }
  th, td { border: 1px solid #e0e0e0; padding: 8px; text-align: left; vertical-align: top; }
  th { background: #f5f5f5; }
  td.name { font-family: monospace; white-space: nowrap; }
  td.name .links { font-size: 11px; margin-top: 4px; font-family: monospace; }
  td.name .links a { color: #1565c0; text-decoration: none; }
  td.name .links a:hover { text-decoration: underline; }
  td.stat { font-family: monospace; white-space: nowrap; }
  td.stat .raw { color: #888; font-size: 11px; }
  img { max-width: 320px; height: auto; display: block; }
  tr.good td.stat { color: #2e7d32; }
  tr.ok   td.stat { color: #ef6c00; }
  tr.bad  td.stat { color: #c62828; font-weight: 700; }
  tr.err  td.stat { color: #c62828; font-weight: 700; }
</style>
</head>
<body>
  <h1>dom-to-pptx fidelity report</h1>
  <div class="meta">
    Generated ${new Date().toISOString()} —
    ${results.length} cases,
    ${results.filter((r) => !r.error).length} succeeded,
    ${results.filter((r) => r.error).length} errored
  </div>
  <table>
    <thead>
      <tr><th>Case</th><th>Δ (content / raw)</th><th>Source</th><th>PPTX</th><th>Diff</th></tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>
</body>
</html>`;

  await writeFile(reportPath, html, 'utf8');
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  })[c]);
}
