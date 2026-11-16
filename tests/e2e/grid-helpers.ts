import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, type Page } from '@playwright/test';

export interface CellRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface DebugApi {
  getState(): {
    fileName: string | null;
    active: { row: number; col: number };
    selection: { anchor: { row: number; col: number }; focus: { row: number; col: number } };
    focusZone: string;
    status: string;
    canUndo: boolean;
  };
  getCellDisplay(row: number, col: number): string | null;
  getCellRect(row: number, col: number): CellRect | null;
  getScroll(): { x: number; y: number };
  getFreeze(): { rows: number; cols: number };
  getCellStyle(
    row: number,
    col: number,
  ): {
    font: {
      name: string;
      size: number;
      bold: boolean;
      italic: boolean;
      underline: boolean;
      underlineStyle: string;
      strike: boolean;
      color: string | null;
    };
    fill: { pattern: string; color: string | null; backgroundColor: string | null };
    border: {
      left: { style: string; color: string | null; weight: string | null };
      right: { style: string; color: string | null; weight: string | null };
      top: { style: string; color: string | null; weight: string | null };
      bottom: { style: string; color: string | null; weight: string | null };
    };
    alignment: {
      horizontal: string;
      vertical: string;
      wrapText: boolean;
      indent: number;
      textRotation: number;
      shrinkToFit: boolean;
    };
    numberFormat: string | null;
    numberFormatColor: string | null;
    displayText: string;
    merge: { anchorRow: number; anchorCol: number; rowSpan: number; colSpan: number } | null;
  } | null;
  getFillHandle(): { x: number; y: number; width: number; height: number } | null;
}

declare global {
  interface Window {
    __calcplus: DebugApi;
  }
}

export function fixturePath(name: string): string {
  return fileURLToPath(new URL(`../../fixtures/compat/${name}`, import.meta.url));
}

export function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(`../../fixtures/compat/${name}`, import.meta.url)));
}

export async function openFixture(page: Page, name: string): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles(fixturePath(name));
  await expect(page.getByTestId('file-name')).toHaveText(name);
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

export async function cellRect(page: Page, row: number, col: number): Promise<CellRect> {
  const rect = await page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellRect(targetRow as number, targetCol as number),
    [row, col] as const,
  );
  if (rect === null) throw new Error(`No rectangle for row ${row}, column ${col}`);
  return rect;
}

export async function clickCell(page: Page, row: number, col: number): Promise<void> {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const rect = await cellRect(page, row, col);
  await page.mouse.click(box.x + rect.x + rect.width / 2, box.y + rect.y + rect.height / 2);
}

/** The colour of a pixel inside the grid canvas, in canvas coordinates. */
export async function pixelAt(
  page: Page,
  x: number,
  y: number,
): Promise<{ r: number; g: number; b: number; a: number }> {
  return page.evaluate(
    ([pointX, pointY]) => {
      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
      if (canvas === null) throw new Error('No grid canvas');
      const context = canvas.getContext('2d');
      if (context === null) throw new Error('No canvas context');
      const ratio = canvas.width / canvas.getBoundingClientRect().width;
      const data = context.getImageData(
        Math.round(pointX * ratio),
        Math.round(pointY * ratio),
        1,
        1,
      ).data;
      return { r: data[0] ?? 0, g: data[1] ?? 0, b: data[2] ?? 0, a: data[3] ?? 0 };
    },
    [x, y] as const,
  );
}

/**
 * A signature of the painted pixels of a region of the canvas. Two regions
 * with the same signature look the same to the user.
 */
export async function regionSignature(
  page: Page,
  x: number,
  y: number,
  width: number,
  height: number,
): Promise<string> {
  return page.evaluate(
    ([regionX, regionY, regionWidth, regionHeight]) => {
      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
      if (canvas === null) throw new Error('No grid canvas');
      const context = canvas.getContext('2d');
      if (context === null) throw new Error('No canvas context');
      const ratio = canvas.width / canvas.getBoundingClientRect().width;
      const data = context.getImageData(
        Math.round(regionX * ratio),
        Math.round(regionY * ratio),
        Math.round(regionWidth * ratio),
        Math.round(regionHeight * ratio),
      ).data;
      const parts: string[] = [];
      for (let index = 0; index < data.length; index += 16) {
        parts.push(`${data[index]},${data[index + 1]},${data[index + 2]}`);
      }
      return parts.join('|');
    },
    [x, y, width, height] as const,
  );
}

/** The test id of the element that holds the keyboard focus. */
export async function focusedTestId(page: Page): Promise<string | null> {
  return page.evaluate(() => document.activeElement?.getAttribute('data-testid') ?? null);
}
