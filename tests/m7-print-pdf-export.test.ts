import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { exportPdfPlan, pdfFileName } from '@commands/chrome';
import inventory from '@commands/command-inventory.json';
import { createCommandRegistry } from '@commands/registry';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { makePage, makePlan, textItem } from './helpers/print-pdf';

/**
 * M7 lane 3 — File → Export PDF (print-pdf §2; ui-chrome §4.1: the map entry
 * lands before UI).
 *
 * The command flow runs for real: the plan goes through PDFKit with the
 * bundled fonts (served from disk through a fetch stub), and the download path
 * is stubbed at the DOM edge so the test can inspect the Blob it receives.
 */

const originalFetch = globalThis.fetch;
const originalCreateObjectURL = URL.createObjectURL;
const originalRevokeObjectURL = URL.revokeObjectURL;

interface Anchor {
  href: string;
  download: string;
  style: Record<string, string>;
  click?: () => void;
  remove?: () => void;
}

let anchor: Anchor;
let captured: Blob | null;

beforeEach(() => {
  vi.useFakeTimers();
  captured = null;
  anchor = { href: '', download: '', style: {}, click: () => undefined, remove: () => undefined };
  (globalThis as unknown as { document: unknown }).document = {
    createElement: () => anchor,
    body: { appendChild: () => undefined },
  };
  URL.createObjectURL = ((blob: Blob) => {
    captured = blob;
    return 'blob:pdf';
  }) as typeof URL.createObjectURL;
  URL.revokeObjectURL = (() => undefined) as typeof URL.revokeObjectURL;
  globalThis.fetch = (async (input: string | URL | Request) => {
    const url = String(input);
    const bytes = new Uint8Array(await readFile(fileURLToPath(url)));
    return new Response(bytes);
  }) as typeof fetch;
});

afterEach(() => {
  vi.useRealTimers();
  globalThis.fetch = originalFetch;
  URL.createObjectURL = originalCreateObjectURL;
  URL.revokeObjectURL = originalRevokeObjectURL;
  delete (globalThis as unknown as { document?: unknown }).document;
});

describe('file.exportPdf', () => {
  it('is in the command map and the registry, and enabled with a workbook open', () => {
    const entry = inventory.find((candidate) => candidate.id === 'file.exportPdf');
    expect(entry).toBeDefined();
    expect(entry?.implemented).toBe(true);
    expect(entry?.category).toBe('file');
    expect(entry?.matrixFeature).toBe(
      'Print and PDF export: print area, orientation, fit-to-width, margins',
    );

    const registry = createCommandRegistry();
    const command = registry.require('file.exportPdf');
    const workbook = createEmptyWorkbook();
    workbook.sheets.push(createSheet('Sheet1', 'xl/worksheets/sheet1.xml'));
    const context = { workbook, activeSheetId: null, selection: null, activeCell: null };
    expect(command.isEnabled(context)).toBe(true);
    expect(command.isEnabled({ ...context, busy: true })).toBe(false);
  });

  it('names the download after the workbook, or the sheet when unsaved', () => {
    expect(pdfFileName('Report.xlsx', 'Sheet1')).toBe('Report.pdf');
    expect(pdfFileName('Macro.XLSM', 'Sheet1')).toBe('Macro.pdf');
    expect(pdfFileName(null, 'Budget')).toBe('Budget.pdf');
    expect(pdfFileName('  ', 'Budget')).toBe('Budget.pdf');
  });

  it('renders a plan and downloads a PDF Blob', async () => {
    const plan = makePlan([makePage(0, [textItem({ text: 'Export me' })])]);
    const summary = await exportPdfPlan(plan, 'Report.xlsx');

    expect(summary).toEqual({ fileName: 'Report.pdf', pageCount: 1, fallbackPages: [] });
    expect(anchor.download).toBe('Report.pdf');
    expect(captured).not.toBeNull();
    expect(captured?.type).toBe('application/pdf');
    const bytes = new Uint8Array(await captured!.arrayBuffer());
    expect(Buffer.from(bytes.subarray(0, 8)).toString('latin1')).toBe('%PDF-1.3');
    expect(Buffer.from(bytes.subarray(-6)).toString('latin1')).toBe('%%EOF\n');
  });
});
