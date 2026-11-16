import { workbookDateSystem } from '@model/date-system';
import type { CommandDefinition } from './types';

export interface DateSystemArgs {
  readonly date1904: boolean;
}

/**
 * Switch the workbook date system. The mapping from serials to dates changes,
 * and the formula engine is rebuilt by the calculation scheduler; the authored
 * values and formulas stay untouched.
 */
export const workbookDateSystemCommand: CommandDefinition<DateSystemArgs> = {
  id: 'workbook.dateSystem',
  label: 'Date system',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args) {
    if (context.workbook.date1904 === args.date1904) {
      return { changes: [] };
    }
    return {
      changes: [
        {
          type: 'workbook.dateSystem',
          before: context.workbook.date1904,
          after: args.date1904,
        },
      ],
    };
  },
};

/** The 1900/1904 flag as the UI and reports label it. */
export function dateSystemLabel(workbook: { readonly date1904: boolean }): '1900' | '1904' {
  return workbookDateSystem(workbook.date1904);
}
