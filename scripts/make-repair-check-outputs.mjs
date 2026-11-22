/**
 * Builds the manual repair-check output set for release-gate item 6
 * (docs/release-gate.md): every workbook calcplus can open in the
 * compatibility corpus, saved through the production open/save pipeline.
 *
 * For each fixture it writes two files into fixtures/repair-check/:
 *   <name>.noop.xlsx    - opened and saved untouched
 *   <name>.edited.xlsx  - one cell edited through the command layer, then saved
 *
 * The maintainer then opens the folder in Excel 365 and confirms none of the
 * files shows a repair prompt. The deliberately hostile fixtures (corrupt,
 * encrypted, zip-bomb) and the Strict file are rejected by design
 * (file-compat Tier C) and are skipped by name, like verify-repair-free.mjs.
 *
 * Run with: npm run repair-outputs
 */
import { mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const compatDir = join(root, 'fixtures', 'compat');
const outDir = join(root, 'fixtures', 'repair-check');
const rejectedByDesign = new Set([
  'corrupt.xlsx',
  'encrypted.xlsx',
  'zip-bomb.xlsx',
  'strict-ooxml.xlsx',
]);

function collectPackages(path) {
  const absolute = resolve(path);
  const stats = statSync(absolute);
  if (stats.isFile()) return /\.(xlsx|xlsm)$/i.test(absolute) ? [absolute] : [];
  const files = [];
  for (const entry of readdirSync(absolute).sort()) {
    files.push(...collectPackages(join(absolute, entry)));
  }
  return files;
}

async function loadModules() {
  const server = await createServer({
    root,
    configFile: join(root, 'vite.config.ts'),
    appType: 'custom',
    logLevel: 'silent',
    server: { middlewareMode: true },
  });
  try {
    const open = await server.ssrLoadModule('/src/ooxml/open.ts');
    const save = await server.ssrLoadModule('/src/ooxml/save.ts');
    const commands = await server.ssrLoadModule('/src/commands/session.ts');
    const value = await server.ssrLoadModule('/src/model/value.ts');
    return {
      openWorkbook: open.openWorkbook,
      saveWorkbook: save.saveWorkbook,
      DocumentSession: commands.DocumentSession,
      numberValue: value.numberValue,
    };
  } finally {
    await server.close();
  }
}

function outputName(file, mode) {
  const relativePath = relative(compatDir, file);
  const extension = relativePath.slice(relativePath.lastIndexOf('.'));
  return relativePath.slice(0, -extension.length) + `.${mode}${extension}`;
}

const { openWorkbook, saveWorkbook, DocumentSession, numberValue } = await loadModules();

rmSync(outDir, { recursive: true, force: true });

const manifest = { generatedAt: new Date().toISOString(), written: [], skipped: [], failed: [] };

for (const file of collectPackages(compatDir)) {
  const name = relative(compatDir, file);
  if (rejectedByDesign.has(name)) {
    manifest.skipped.push({ file: name, reason: 'rejected by design (file-compat Tier C)' });
    continue;
  }
  const bytes = new Uint8Array(readFileSync(file));

  try {
    // A no-op save: the identity path.
    const opened = openWorkbook(bytes);
    const noop = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });

    // An edited save: one cell through the command layer, so the patch writer
    // runs with real dirty maps. A cell inside a shared or array formula group
    // is refused by design (a split grouped formula loses data, D-8), so try a
    // few candidates and use the first that edits cleanly.
    const editedOpen = openWorkbook(bytes);
    const session = new DocumentSession(editedOpen.workbook);
    const sheet = editedOpen.workbook.sheets[0];
    const candidates = [
      [0, 0],
      [0, 1],
      [1, 0],
      [0, 25],
      [0, 50],
      [25, 25],
    ];
    for (const [row, col] of candidates) {
      session.dispatch('cell.set', {
        sheetId: sheet.id,
        row,
        col,
        value: numberValue(1),
      });
      if (session.dirtyCellMap().size > 0) break;
    }
    if (session.dirtyCellMap().size === 0) {
      throw new Error('no candidate cell accepted the cell.set edit');
    }
    const edited = saveWorkbook({
      workbook: editedOpen.workbook,
      package: editedOpen.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: session.dirtyMergeSheets(),
    });

    for (const [mode, result] of [
      ['noop', noop],
      ['edited', edited],
    ]) {
      const outPath = join(outDir, outputName(file, mode));
      mkdirSync(dirname(outPath), { recursive: true });
      writeFileSync(outPath, result.bytes);
      manifest.written.push(outputName(file, mode));
    }
    console.log(`ok       ${name}`);
  } catch (error) {
    manifest.failed.push({ file: name, reason: error.message });
    console.error(`FAILED   ${name}: ${error.message}`);
  }
}

mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');

console.log('');
console.log(`Wrote ${manifest.written.length} files to fixtures/repair-check/.`);
console.log(`Skipped: ${manifest.skipped.length}. Failed: ${manifest.failed.length}.`);
console.log('Open the folder in Excel 365 and confirm no file shows a repair prompt.');

if (manifest.failed.length > 0) process.exitCode = 1;
