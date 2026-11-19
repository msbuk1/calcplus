import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import {
  SIGNATURE_INVALIDATION_MESSAGE,
  confirmSignatureInvalidation,
  planSignatureSave,
} from '@io/confirm';
import { numberValue } from '@model/value';
import { openWorkbook, type OpenedWorkbook } from '@ooxml/open';
import { buildExportReport } from '@ooxml/report';
import { saveWorkbook, type SaveWorkbookResult } from '@ooxml/save';
import { listPreservedParts } from '@ooxml/part-policy';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

/**
 * M6 — digital signatures warn before writing (file-compat §6.5, §7, §10.13).
 *
 * "Signature parts are preserved, but any save invalidates them. The export
 * report warns before writing, and the saved file is reported as having an
 * invalid signature." The gate is the one `App.saveFile` runs next to the
 * macro plan hook; this suite exercises it with the real save pipeline behind
 * it, so a declined confirmation provably writes nothing.
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SIGNATURE_PART = '_xmlsignatures/sig1.xml';

const WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`;

const SIGNATURE = `${HEADER}
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><SignedInfo/></Signature>`;

/** A workbook whose package carries a signature part, as Excel writes one. */
function signedPackage(): Uint8Array {
  return buildXlsx(WORKSHEET, { extraParts: { [SIGNATURE_PART]: SIGNATURE } });
}

function unsignedPackage(): Uint8Array {
  return buildXlsx(WORKSHEET);
}

function saveSession(opened: OpenedWorkbook, session: DocumentSession): SaveWorkbookResult {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: session.dirtyCellMap(),
    dirtyDimensions: session.dirtyDimensionSnapshot(),
    dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
  });
}

/**
 * The `App.saveFile` gate in the same order: ask, and only then write. A `null`
 * result is the declined path — the caller returns before `saveWorkbook`.
 */
function saveThroughGate(
  opened: OpenedWorkbook,
  session: DocumentSession,
  answer: boolean,
): {
  readonly asked: string[];
  readonly declined: boolean;
  readonly saved: SaveWorkbookResult | null;
} {
  const asked: string[] = [];
  const proceed = planSignatureSave({
    hasDigitalSignatures: opened.compatibility.hasDigitalSignatures,
    changed: session.hasUnsavedChanges,
    confirmFn: (message) => {
      asked.push(message);
      return answer;
    },
  });
  if (!proceed) return { asked, declined: true, saved: null };
  return { asked, declined: false, saved: saveSession(opened, session) };
}

function editedSession(opened: OpenedWorkbook): DocumentSession {
  const session = new DocumentSession(opened.workbook);
  session.dispatch('cell.set', {
    sheetId: opened.workbook.sheets[0]!.id,
    row: 1,
    col: 0,
    value: numberValue(2),
  });
  return session;
}

describe('F-sig digital signatures warn before writing', () => {
  it('asks with the contract wording before writing a signed workbook', () => {
    const opened = openWorkbook(signedPackage());
    expect(opened.compatibility.hasDigitalSignatures).toBe(true);

    const { asked, saved } = saveThroughGate(opened, editedSession(opened), true);
    expect(asked).toEqual([SIGNATURE_INVALIDATION_MESSAGE]);
    expect(SIGNATURE_INVALIDATION_MESSAGE).toBe(
      "Saving will invalidate the workbook's digital signature. Save anyway?",
    );
    expect(saved).not.toBeNull();
    expect(saved!.noOp).toBe(false);
  });

  it('writes nothing when the confirmation is declined and keeps the package intact', () => {
    const original = signedPackage();
    const opened = openWorkbook(original);
    const { asked, declined, saved } = saveThroughGate(opened, editedSession(opened), false);
    expect(asked).toHaveLength(1);
    expect(declined).toBe(true);
    expect(saved).toBeNull();
    // Nothing was written, so the signature is still the one that was opened.
    expect(unzipPart(original, SIGNATURE_PART)).toBe(SIGNATURE);
  });

  it('accepts and writes, preserving the signature part but reporting it invalidated', () => {
    const original = signedPackage();
    const opened = openWorkbook(original);
    const session = editedSession(opened);
    const { saved } = saveThroughGate(opened, session, true);
    expect(saved).not.toBeNull();

    // The parts are preserved; the meaning is not, so the report says so.
    expect(unzipPart(saved!.bytes, SIGNATURE_PART)).toBe(SIGNATURE);
    const report = buildExportReport({
      workbook: opened.workbook,
      before: opened.package,
      after: saved!.package,
      report: saved!.report,
      compatibility: opened.compatibility,
      noOp: saved!.noOp,
      structuralChange: false,
      preserved: listPreservedParts({
        package: opened.package,
        entries: saved!.package.entries,
        workbook: opened.workbook,
        format: 'xlsx',
      }),
    });
    const signature = report.entries.filter((entry) => entry.kind === 'signature');
    expect(signature).toHaveLength(1);
    expect(signature[0]!.category).toBe('warning');
    expect(report.needsAttention).toBe(true);
  });

  it('does not ask for a save with nothing to write', () => {
    const opened = openWorkbook(signedPackage());
    const session = new DocumentSession(opened.workbook);
    expect(session.hasUnsavedChanges).toBe(false);

    const asked: string[] = [];
    const proceed = planSignatureSave({
      hasDigitalSignatures: opened.compatibility.hasDigitalSignatures,
      changed: session.hasUnsavedChanges,
      confirmFn: (message) => {
        asked.push(message);
        return false;
      },
    });
    // A no-op save returns the original bytes, so the signature stays valid.
    expect(proceed).toBe(true);
    expect(asked).toEqual([]);
    expect(saveSession(opened, session).noOp).toBe(true);
  });

  it('never asks for an unsigned workbook', () => {
    const opened = openWorkbook(unsignedPackage());
    expect(opened.compatibility.hasDigitalSignatures).toBe(false);

    const asked: string[] = [];
    const proceed = planSignatureSave({
      hasDigitalSignatures: opened.compatibility.hasDigitalSignatures,
      changed: true,
      confirmFn: (message) => {
        asked.push(message);
        return false;
      },
    });
    expect(proceed).toBe(true);
    expect(asked).toEqual([]);
  });

  it('exposes the confirmation on its own for a plain prompt', () => {
    const asked: string[] = [];
    expect(
      confirmSignatureInvalidation((message) => {
        asked.push(message);
        return false;
      }),
    ).toBe(false);
    expect(asked).toEqual([SIGNATURE_INVALIDATION_MESSAGE]);
  });
});
