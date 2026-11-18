import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import { formattedSheetXlsx } from '../helpers/formatted-sheet';

/**
 * A merged range that crosses a frozen divider is still one cell. Its label is
 * centred in the merged rectangle and clipped at the merged edge, never at the
 * edge of the pane that holds the anchor cell.
 *
 * The formatted fixture freezes column A and rows 1-2, so A1:D1 and A11:D11
 * both start in a frozen pane and end in the scrolling one. The divider sits at
 * x=221 and the merged ranges reach x=684. The label in A11 is white on the
 * imported navy fill, so the painted ink can be measured against the fill.
 */
test.describe('merged label across the frozen divider', () => {
  test.use({
    viewport: { width: 1280, height: 920 },
    deviceScaleFactor: 1,
    colorScheme: 'light',
    reducedMotion: 'reduce',
  });

  test('paints the merged range and its label across the whole merge', async ({ page }) => {
    await page.goto('/?fsa=0');
    // The same pinned grid viewport the visual baseline uses.
    await page.addStyleTag({
      content: '.grid-area { flex: 0 0 auto; height: 640px; padding: 12px; }',
    });
    await page.getByTestId('file-input').setInputFiles({
      name: 'm4a-formatted.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: Buffer.from(formattedSheetXlsx()),
    });
    await page.waitForFunction(
      () => window.__calcplus.getCellDisplay(0, 0) === 'Quarterly report - formatted',
    );
    await page.evaluate(() => document.fonts.ready.then(() => true));
    // Wait for the scheduled canvas paint, not just the projection update.
    await page.evaluate(
      () =>
        new Promise<void>((resolve) => {
          requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
        }),
    );

    const geometry = await page.evaluate(() => {
      const anchor = window.__calcplus.getCellRect(10, 0);
      const last = window.__calcplus.getCellRect(10, 3);
      const title = window.__calcplus.getCellRect(0, 0);
      const titleLast = window.__calcplus.getCellRect(0, 3);
      if (anchor === null || last === null || title === null || titleLast === null) return null;
      const span = (first: typeof anchor, final: typeof anchor) => ({
        left: first.x + 1,
        top: first.y + 1,
        right: final.x + final.width - 1,
        bottom: first.y + first.height - 1,
        divider: first.x + first.width,
      });
      return { label: span(anchor, last), title: span(title, titleLast) };
    });
    expect(geometry).not.toBeNull();
    const { label, title } = geometry!;

    // The frozen divider really does fall inside both merged ranges.
    expect(label.divider).toBeLessThan(label.right);
    expect(title.divider).toBeLessThan(title.right);

    const measured = await page.evaluate(
      (bands: { label: typeof label; title: typeof title }) => {
        const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
        const context = canvas?.getContext('2d') ?? null;
        if (context === null) return null;
        const scan = (band: typeof label): Record<string, number> => {
          const width = band.right - band.left;
          const height = band.bottom - band.top;
          const data = context.getImageData(band.left, band.top, width, height).data;
          let fill = 0;
          let ink = 0;
          let inkLeft = Number.POSITIVE_INFINITY;
          let inkRight = -1;
          for (let row = 0; row < height; row += 1) {
            for (let column = 0; column < width; column += 1) {
              const offset = (row * width + column) * 4;
              const red = data[offset] ?? 0;
              const green = data[offset + 1] ?? 0;
              const blue = data[offset + 2] ?? 0;
              // The imported fill is #1f4e78 and the label font is white.
              if (red < 70 && green < 110 && blue > 90 && blue < 160) {
                fill += 1;
              } else if (red > 200 && green > 200 && blue > 200) {
                ink += 1;
                const x = band.left + column;
                if (x < inkLeft) inkLeft = x;
                if (x > inkRight) inkRight = x;
              }
            }
          }
          return {
            fill,
            ink,
            inkLeft,
            inkRight,
            pixels: width * height,
            fillFraction: fill / (width * height),
          };
        };
        return { label: scan(bands.label), title: scan(bands.title) };
      },
      { label, title },
    );
    expect(measured).not.toBeNull();
    const { label: labelPixels, title: titlePixels } = measured!;

    // The anchor's fill covers the whole merged range, not just the anchor
    // column: a range cut at the divider leaves the rest of the band white and
    // the fill fraction collapses to the anchor column's share.
    expect(labelPixels.fillFraction, JSON.stringify(labelPixels)).toBeGreaterThan(0.7);
    expect(titlePixels.fillFraction, JSON.stringify(titlePixels)).toBeGreaterThan(0.7);

    // The label ink is one run that reaches well into the covered columns, so
    // the text is neither cut at the anchor column nor buried by the covered
    // cells' fills painted afterwards.
    expect(labelPixels.inkRight, JSON.stringify(labelPixels)).toBeGreaterThan(label.divider + 120);
    expect(labelPixels.ink, JSON.stringify(labelPixels)).toBeGreaterThan(100);

    // The title merge is white bold 12px; its ink sits well inside the merge,
    // so it only survives if the whole merged range is painted.
    expect(titlePixels.inkRight, JSON.stringify(titlePixels)).toBeGreaterThan(title.divider + 100);
    expect(titlePixels.ink, JSON.stringify(titlePixels)).toBeGreaterThan(100);
  });
});
