import type { JSX, RefObject } from 'react';
import appIcon from '../../assets/calcplus.svg';
import { QuickAccessToolbar } from './QuickAccessToolbar';
import type { CommandStateResolver } from './ribbon-model';

export interface AppHeaderProps {
  readonly fileName: string | null;
  readonly dirty: boolean;
  readonly resolveState: CommandStateResolver;
  readonly onCommand: (commandId: string, args?: unknown) => void;
  readonly onFocus?: () => void;
  readonly headerRef?: RefObject<HTMLElement | null>;
}

/**
 * Title strip: the Quick Access Toolbar at the top left, then the document
 * name. The toolbar is above the ribbon and part of the ribbon focus zone
 * (ui-chrome §1); unsaved state shows on its Save button.
 */
export function AppHeader({
  fileName,
  dirty,
  resolveState,
  onCommand,
  onFocus,
  headerRef,
}: AppHeaderProps): JSX.Element {
  return (
    <header className="app-header" ref={headerRef} onFocusCapture={onFocus}>
      <QuickAccessToolbar dirty={dirty} resolveState={resolveState} onCommand={onCommand} />
      <div className="brand">
        <img className="brand-icon" src={appIcon} alt="" width={26} height={26} />
        <span className="brand-name">calcplus</span>
      </div>
      <div className="file-meta">
        <span className="file-name" data-testid="file-name">
          {fileName ?? 'No workbook open'}
        </span>
      </div>
    </header>
  );
}
