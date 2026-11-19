import type { CommandDefinition } from './types';

/**
 * Chrome-only commands. Their `run` is a no-op: the application shell owns the
 * browser file APIs and transient view state, and dispatches these IDs through
 * its command host. They still live in the command registry so enablement and
 * checked state have one source (command-system §1).
 */

const fileNewCommand: CommandDefinition<void> = {
  id: 'file.new',
  label: 'New',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileOpenCommand: CommandDefinition<void> = {
  id: 'file.open',
  label: 'Open',
  category: 'file',
  keybindings: ['Ctrl+O'],
  isEnabled(context) {
    return context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileSaveCommand: CommandDefinition<void> = {
  id: 'file.save',
  label: 'Save',
  category: 'file',
  keybindings: ['Ctrl+S'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileSaveAsCommand: CommandDefinition<void> = {
  id: 'file.saveAs',
  label: 'Save As',
  category: 'file',
  keybindings: ['Ctrl+Shift+S'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const viewGridlinesCommand: CommandDefinition<void> = {
  id: 'view.gridlines',
  label: 'Gridlines',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  isChecked(context) {
    return context.view?.showGridLines ?? false;
  },
  run() {
    return { changes: [] };
  },
};

const viewHeadingsCommand: CommandDefinition<void> = {
  id: 'view.headings',
  label: 'Headings',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  isChecked(context) {
    return context.view?.showHeadings ?? false;
  },
  run() {
    return { changes: [] };
  },
};

const viewZoom100Command: CommandDefinition<void> = {
  id: 'view.zoom100',
  label: 'Zoom 100%',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

const nameBoxNavigateCommand: CommandDefinition<void> = {
  id: 'nameBox.navigate',
  label: 'Go to address',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.activeSheetId !== null;
  },
  run() {
    return { changes: [] };
  },
};

const helpAboutCommand: CommandDefinition<void> = {
  id: 'help.about',
  label: 'About calcplus',
  category: 'help',
  keybindings: [],
  isEnabled() {
    return true;
  },
  run() {
    return { changes: [] };
  },
};

const fileCompatibilityReportCommand: CommandDefinition<void> = {
  id: 'file.compatibilityReport',
  label: 'Compatibility Report',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

const fileExportReportCommand: CommandDefinition<void> = {
  id: 'file.exportReport',
  label: 'Export Report',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

export function chromeCommands(): readonly CommandDefinition<never>[] {
  return [
    fileNewCommand,
    fileOpenCommand,
    fileSaveCommand,
    fileSaveAsCommand,
    viewGridlinesCommand,
    viewHeadingsCommand,
    viewZoom100Command,
    nameBoxNavigateCommand,
    fileCompatibilityReportCommand,
    fileExportReportCommand,
    helpAboutCommand,
  ] as unknown as readonly CommandDefinition<never>[];
}
