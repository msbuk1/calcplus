import type { DefinedName, WorkbookProtectionRecord } from '@model/types';
import {
  RELATIONSHIPS_NS,
  STRICT_SPREADSHEET_NS,
  SPREADSHEET_NS,
  attributeValue,
  decodeXml,
  openTagStart,
  parseXml,
} from './xml';
import { WorkbookOpenError } from './errors';

export interface WorkbookSheetReference {
  readonly name: string;
  readonly sheetId: number;
  readonly relationshipId: string;
  readonly state: 'visible' | 'hidden' | 'veryHidden';
}

export interface WorkbookXmlInfo {
  readonly date1904: boolean;
  readonly sheets: WorkbookSheetReference[];
  readonly definedNames: DefinedName[];
  readonly activeTab: number;
  readonly calcOnLoad: boolean;
  readonly fullCalcOnLoad: boolean;
  /**
   * The workbook's `<workbookProtection>` element, verbatim (file-compat §6.7).
   * Preserved and never enforced; the element itself lives in the workbook part,
   * which only a name or sheet-list change rewrites.
   */
  readonly protection: WorkbookProtectionRecord | null;
}

export function parseWorkbookXml(bytes: Uint8Array, partName: string): WorkbookXmlInfo {
  const xml = decodeXml(bytes);
  const sheets: WorkbookSheetReference[] = [];
  const definedNames: DefinedName[] = [];
  let date1904 = false;
  let activeTab = 0;
  let calcOnLoad = false;
  let fullCalcOnLoad = false;
  let protection: WorkbookProtectionRecord | null = null;
  let sawWorkbook = false;
  let inSheets = false;
  let inDefinedNames = false;
  let inDefinedName = false;

  const appendDefinedNameText = (text: string): void => {
    if (!inDefinedName) return;
    const last = definedNames[definedNames.length - 1];
    if (last !== undefined) {
      definedNames[definedNames.length - 1] = {
        ...last,
        formula: last.formula + text,
      };
    }
  };

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        if (tag.uri === STRICT_SPREADSHEET_NS) {
          throw new WorkbookOpenError(
            'strict-ooxml',
            'This workbook uses an OOXML variant calcplus does not support yet.',
          );
        }
        if (tag.local === 'workbook' && tag.uri === SPREADSHEET_NS) {
          sawWorkbook = true;
          return;
        }
        if (!sawWorkbook) return;
        switch (tag.local) {
          case 'workbookPr': {
            date1904 = isTruthy(attributeValue(tag, 'date1904'));
            break;
          }
          case 'workbookProtection': {
            // The open tag is the whole element in every producer: the schema's
            // children are empty. Preserve the authored bytes so a rewrite of
            // the workbook part carries protection through untouched.
            protection = { rawXml: xml.slice(openTagStart(xml, position), position) };
            break;
          }
          case 'workbookView': {
            const value = Number(attributeValue(tag, 'activeTab') ?? '0');
            if (Number.isInteger(value) && value >= 0) activeTab = value;
            break;
          }
          case 'sheets': {
            inSheets = true;
            break;
          }
          case 'sheet': {
            if (!inSheets) break;
            const name = attributeValue(tag, 'name') ?? '';
            const numericSheetId = Number(attributeValue(tag, 'sheetId') ?? '0');
            const relationshipId = relationshipIdAttribute(tag);
            const stateText = attributeValue(tag, 'state') ?? 'visible';
            const state =
              stateText === 'hidden' || stateText === 'veryHidden' ? stateText : 'visible';
            if (name === '' || relationshipId === '') {
              throw new WorkbookOpenError('corrupt', `Malformed sheet entry in ${partName}`);
            }
            sheets.push({ name, sheetId: numericSheetId, relationshipId, state });
            break;
          }
          case 'definedNames': {
            inDefinedNames = true;
            break;
          }
          case 'definedName': {
            if (inDefinedNames) {
              const name = attributeValue(tag, 'name') ?? '';
              const localSheetIdText = attributeValue(tag, 'localSheetId');
              inDefinedName = true;
              definedNames.push({
                name,
                formula: '',
                sheetId:
                  localSheetIdText === undefined
                    ? null
                    : (`pending-sheet-index:${localSheetIdText}` as DefinedName['sheetId']),
              });
            }
            break;
          }
          case 'calcPr': {
            calcOnLoad = isTruthy(attributeValue(tag, 'calcOnLoad'));
            fullCalcOnLoad = isTruthy(attributeValue(tag, 'fullCalcOnLoad'));
            break;
          }
          default:
            break;
        }
      },
      onCloseTag(tag) {
        if (tag.local === 'sheets') inSheets = false;
        if (tag.local === 'definedName') inDefinedName = false;
        if (tag.local === 'definedNames') inDefinedNames = false;
      },
      onText(text) {
        appendDefinedNameText(text);
      },
      onCdata(text) {
        appendDefinedNameText(text);
      },
    },
    partName,
  );

  if (!sawWorkbook) {
    throw new WorkbookOpenError('corrupt', `${partName} is not a SpreadsheetML workbook part.`);
  }
  if (sheets.length === 0) {
    throw new WorkbookOpenError('corrupt', 'The workbook contains no sheets.');
  }

  return {
    date1904,
    sheets,
    definedNames,
    activeTab,
    calcOnLoad,
    fullCalcOnLoad,
    protection,
  };
}

function relationshipIdAttribute(tag: {
  attributes: Record<string, { local: string; uri?: string; value: string }>;
}): string {
  for (const attribute of Object.values(tag.attributes)) {
    if (attribute.local === 'id' && attribute.uri === RELATIONSHIPS_NS) return attribute.value;
  }
  return '';
}

function isTruthy(value: string | undefined): boolean {
  return value === '1' || value === 'true' || value === 'on';
}
