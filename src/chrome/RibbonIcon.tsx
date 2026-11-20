import type { JSX } from 'react';
import { RIBBON_ICONS, RIBBON_ICON_SIZE } from './ribbon-icons';
import type { RibbonIconName } from './ribbon-icons';

/**
 * Draws one icon from the calcplus family. The control owns the accessible
 * name; the SVG is decorative and shares the control's `currentColor`.
 */
export function RibbonIcon({ name }: { readonly name: RibbonIconName }): JSX.Element {
  const icon = RIBBON_ICONS[name];
  return (
    <svg
      className="ribbon-glyph"
      viewBox={`0 0 ${RIBBON_ICON_SIZE} ${RIBBON_ICON_SIZE}`}
      aria-hidden="true"
      focusable="false"
    >
      {icon.paths.map((path, index) => (
        <path key={`outline-${index}`} d={path} />
      ))}
      {icon.filled?.map((path, index) => (
        <path key={`filled-${index}`} d={path} className="ribbon-glyph-filled" />
      ))}
      {icon.muted?.map((path, index) => (
        <path key={`muted-${index}`} d={path} className="ribbon-glyph-muted" />
      ))}
      {icon.dashed?.map((path, index) => (
        <path key={`dashed-${index}`} d={path} strokeDasharray="1.6 1.4" />
      ))}
      {icon.circles?.map((circle, index) => (
        <circle
          key={`circle-${index}`}
          cx={circle.cx}
          cy={circle.cy}
          r={circle.radius}
          className={circle.filled === true ? 'ribbon-glyph-filled' : undefined}
        />
      ))}
    </svg>
  );
}
