import type { JSX } from 'react';
import appIcon from '../../assets/calcplus.svg';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';

export interface AboutDialogProps {
  readonly onClose: (reason: DialogCloseReason) => void;
}

/** The v1 About dialog. It is the shell's first filled Dialog, not one of the deferred feature dialogs. */
export function AboutDialog({ onClose }: AboutDialogProps): JSX.Element {
  return (
    <Dialog
      title="About calcplus"
      onClose={onClose}
      defaultAction={{ label: 'Close', onSelect: () => onClose('dismiss') }}
    >
      <div className="about-dialog">
        <img className="about-icon" src={appIcon} alt="" width={44} height={44} />
        <div>
          <h3 className="about-name">calcplus</h3>
          <p className="about-version">Version 0.1.0</p>
          <p>
            A browser-based, single-user spreadsheet that follows Excel desktop behaviour for the
            features it supports.
          </p>
          <p>
            Calcplus is free software under GPL-3.0-only. Copyright © calcplus contributors. This
            program comes with absolutely no warranty; see the licence for details.
          </p>
          <p className="about-smartscreen">
            The Windows download is unsigned, so SmartScreen may warn "Windows protected your PC"
            until the download builds reputation. Choose <strong>More info</strong>, check the name,
            then <strong>Run anyway</strong>. You can verify the download against the SHA256SUMS
            file on the release page.
          </p>
          <p>
            <a
              href="https://github.com/msbuk1/calcplus/blob/main/LICENSE"
              target="_blank"
              rel="noreferrer"
            >
              GNU GPL version 3 full text
            </a>
            {' · '}
            <a
              href="https://github.com/msbuk1/calcplus/blob/main/THIRD_PARTY_NOTICES.md"
              target="_blank"
              rel="noreferrer"
            >
              Third-party notices
            </a>
            {' · '}
            <a
              href="https://github.com/msbuk1/calcplus/blob/main/NOTICE"
              target="_blank"
              rel="noreferrer"
            >
              Copyright and warranty notice
            </a>
          </p>
          <p>
            <a href="https://github.com/msbuk1/calcplus" target="_blank" rel="noreferrer">
              Corresponding source
            </a>
          </p>
        </div>
      </div>
    </Dialog>
  );
}
