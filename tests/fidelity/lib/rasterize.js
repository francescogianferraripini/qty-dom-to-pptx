import { spawn } from 'node:child_process';
import { mkdir, readFile, readdir, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';

/**
 * Convert a PPTX file (on disk) to a PNG of slide 1 using LibreOffice headless.
 * Returns a Buffer of the PNG.
 */
export async function pptxToPng(pptxPath) {
  const outDir = path.join(
    os.tmpdir(),
    `dom-to-pptx-fidelity-${Date.now()}-${Math.random().toString(36).slice(2)}`,
  );
  await mkdir(outDir, { recursive: true });
  const userProfile = path.join(outDir, '.lo-profile');

  try {
    await runLibreOffice([
      `-env:UserInstallation=file://${userProfile}`,
      '--headless',
      '--convert-to',
      'png',
      '--outdir',
      outDir,
      pptxPath,
    ]);

    const files = await readdir(outDir);
    const png = files.find((f) => f.toLowerCase().endsWith('.png'));
    if (!png) {
      throw new Error(
        `LibreOffice produced no PNG for ${pptxPath} (got: ${files.join(', ')})`,
      );
    }
    return await readFile(path.join(outDir, png));
  } finally {
    await rm(outDir, { recursive: true, force: true });
  }
}

function runLibreOffice(args) {
  return new Promise((resolve, reject) => {
    const child = spawn('libreoffice', args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stderr = '';
    child.stderr.on('data', (d) => (stderr += d.toString()));
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`libreoffice exited ${code}: ${stderr}`));
    });
  });
}

/**
 * Helper: write buffer to disk first, then rasterize. Some tests prefer this signature.
 */
export async function pptxBufferToPng(pptxBuffer, scratchDir, name) {
  await mkdir(scratchDir, { recursive: true });
  const pptxPath = path.join(scratchDir, `${name}.pptx`);
  await writeFile(pptxPath, pptxBuffer);
  return pptxToPng(pptxPath);
}
