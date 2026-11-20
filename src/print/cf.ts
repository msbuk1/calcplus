import { cfPaintOps, darken, lighten } from '@grid/cf-paint';
import type { CfIconSetId } from '@model/conditional-format';
import type { ProjectedCell, SheetProjection } from '@view/types';
import type {
  PrintBox,
  PrintConditionalPaint,
  PrintConditionalText,
  PrintIconFamily,
} from './types';

/**
 * Conditional formatting for print (conditional formatting §5.1, print-pdf §4).
 *
 * The grid's `cfPaintOps` already resolves rules into an ordered paint plan:
 * priority, stop-if-true, the merged dxf, colour scales, data bars, icons and
 * the CF number format. This module does not evaluate anything itself — it
 * calls the same function with a print rectangle at 100% and converts the
 * canvas-pixel ops into the plan's cell-local points.
 *
 * Units: `cfPaintOps` takes CSS pixels at 96 dpi. The adapter supplies the
 * cell (or merged range) size in sheet points, so the rect here is that size
 * divided by 0.75. Every result is multiplied back by 0.75, which keeps an
 * icon at Excel's 12-device-pixel size and a bar's fractions of the cell
 * exact. The print scale is applied later, uniformly, by the planner.
 */

const PT_PER_PIXEL = 0.75;

/**
 * The icon families and palettes. This mirrors `ICON_STYLES` in
 * `src/grid/cf-paint.ts` (which does not export it); the print plan resolves
 * both so the SVG and PDF renderers need only draw the seven glyph families.
 * Keep the two tables in step.
 */
const ICON_STYLES: Readonly<
  Record<CfIconSetId, { readonly family: PrintIconFamily; readonly colors: readonly string[] }>
> = Object.freeze({
  '3Arrows': { family: 'arrow', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  '3ArrowsGray': { family: 'arrow', colors: ['#D9D9D9', '#A6A6A6', '#7F7F7F'] },
  '4Arrows': { family: 'arrow', colors: ['#F8696B', '#FFB628', '#FFEB84', '#63BE7B'] },
  '4ArrowsGray': { family: 'arrow', colors: ['#D9D9D9', '#B0B0B0', '#A6A6A6', '#7F7F7F'] },
  '5Arrows': { family: 'arrow', colors: ['#F8696B', '#F9966B', '#FFEB84', '#63BE7B', '#2E9E4F'] },
  '5ArrowsGray': {
    family: 'arrow',
    colors: ['#D9D9D9', '#C4C4C4', '#A6A6A6', '#8C8C8C', '#7F7F7F'],
  },
  '3TrafficLights1': { family: 'traffic', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  '3TrafficLights2': { family: 'traffic', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  '4TrafficLights': { family: 'traffic', colors: ['#000000', '#F8696B', '#FFEB84', '#63BE7B'] },
  '3Flags': { family: 'flag', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  '3Signs': { family: 'sign', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  '3Symbols': { family: 'sign', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
  '3Symbols2': { family: 'sign', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
  '4Rating': { family: 'rating', colors: ['#D9D9D9', '#A6A6A6', '#7F7F7F', '#000000'] },
  '5Rating': {
    family: 'rating',
    colors: ['#D9D9D9', '#D9D9D9', '#A6A6A6', '#7F7F7F', '#000000'],
  },
  '5Quarters': {
    family: 'quarter',
    colors: ['#D9D9D9', '#D9D9D9', '#A6A6A6', '#7F7F7F', '#000000'],
  },
  '4RedToBlack': { family: 'dot', colors: ['#000000', '#4D4D4D', '#A60000', '#F8696B'] },
});

/**
 * Resolve one cell's conditional formatting for print. `width` and `height`
 * are the cell's (or the anchor's merged range's) size in **sheet points** at
 * 100%. Returns `null` when the sheet has no CF, or when no rule paints the
 * cell, so a CF-free plan is byte-for-byte the plan it was before CF support.
 */
export function printConditionalFor(
  projection: SheetProjection,
  cell: ProjectedCell,
  width: number,
  height: number,
): PrintConditionalPaint | null {
  const conditional = projection.conditionalFormatting;
  if (conditional === null || width <= 0 || height <= 0) return null;
  const evaluation = conditional.evaluate(cell.row, cell.col, {
    value: cell.value,
    displayText: cell.displayText,
  });
  if (!evaluation.painted) return null;

  // §5.1 step 7: a CF number format drives the text the cell paints.
  const code = evaluation.overlay.numberFormat?.code;
  const cfDisplayText = code === undefined ? null : conditional.textFor(cell.value, code);
  const rect = {
    x: 0,
    y: 0,
    width: width / PT_PER_PIXEL,
    height: height / PT_PER_PIXEL,
  };
  const ops = cfPaintOps({
    evaluation,
    rect,
    zoom: 1,
    style: cell.style,
    displayText: cell.displayText,
    cfDisplayText,
  });

  let fillColor: string | null = null;
  let dataBar: PrintConditionalPaint['dataBar'] = null;
  let text: PrintConditionalText | null = null;
  let border: PrintConditionalPaint['border'] = null;
  let icon: PrintConditionalPaint['icon'] = null;

  for (const op of ops) {
    switch (op.kind) {
      case 'fill':
        fillColor = op.color;
        break;
      case 'dataBar': {
        const box = scaleRect(op.rect);
        dataBar = {
          box,
          color: op.color,
          gradient:
            op.fill === 'gradient'
              ? op.negative
                ? { from: op.color, to: lighten(op.color, 0.45) }
                : { from: lighten(op.color, 0.45), to: op.color }
              : null,
          borderColor: op.border ? darken(op.color, 0.25) : null,
          axisX: op.axisX === null ? null : op.axisX * PT_PER_PIXEL,
          axisColor: op.axisColor,
        };
        break;
      }
      case 'text':
        text = {
          hidden: op.hidden,
          shift: op.shift * PT_PER_PIXEL,
          style: op.style,
          text: op.text,
          layoutText: op.layoutText,
        };
        break;
      case 'border':
        border = op.border;
        break;
      case 'icon': {
        const style = ICON_STYLES[op.iconSet];
        if (style === undefined) break;
        const index = Math.min(op.iconCount - 1, Math.max(0, op.iconIndex));
        const color = style.colors[Math.min(style.colors.length - 1, index)] ?? '#7F7F7F';
        icon = {
          box: {
            x: op.box.x * PT_PER_PIXEL,
            y: op.box.y * PT_PER_PIXEL,
            width: op.box.size * PT_PER_PIXEL,
            height: op.box.size * PT_PER_PIXEL,
          },
          family: style.family,
          color,
          iconIndex: op.iconIndex,
          iconCount: op.iconCount,
        };
        break;
      }
    }
  }

  // `cfPaintOps` always emits a text op for a painted cell, so this fallback
  // only guards a future shape change.
  return {
    fillColor,
    dataBar,
    text: text ?? {
      hidden: false,
      shift: 0,
      style: cell.style,
      text: cell.displayText,
      layoutText: true,
    },
    border,
    icon,
  };
}

function scaleRect(rect: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}): PrintBox {
  return {
    x: rect.x * PT_PER_PIXEL,
    y: rect.y * PT_PER_PIXEL,
    width: rect.width * PT_PER_PIXEL,
    height: rect.height * PT_PER_PIXEL,
  };
}
