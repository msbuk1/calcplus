import type { CommandDefinition } from './types';
import { cellSetCommand } from './cell-set';
import { clipboardCommands } from './clipboard';
import { dimensionSetCommand } from './dimension';
import { formatCommands } from './format';
import { sheetDeleteCommand, sheetInsertCommand, sheetRenameCommand } from './sheet';
import { workbookDateSystemCommand } from './workbook';

export class CommandRegistry {
  private readonly commands = new Map<string, CommandDefinition<never>>();

  register<Args>(command: CommandDefinition<Args>): void {
    if (this.commands.has(command.id)) {
      throw new Error(`Command id already registered: ${command.id}`);
    }
    this.commands.set(command.id, command as unknown as CommandDefinition<never>);
  }

  get(id: string): CommandDefinition<never> | undefined {
    return this.commands.get(id);
  }

  require(id: string): CommandDefinition<never> {
    const command = this.commands.get(id);
    if (command === undefined) throw new Error(`Unknown command: ${id}`);
    return command;
  }

  list(): CommandDefinition<never>[] {
    return [...this.commands.values()];
  }
}

const undoCommand: CommandDefinition<void> = {
  id: 'edit.undo',
  label: 'Undo',
  category: 'edit',
  keybindings: ['Ctrl+Z'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

const redoCommand: CommandDefinition<void> = {
  id: 'edit.redo',
  label: 'Redo',
  category: 'edit',
  keybindings: ['Ctrl+Y', 'Ctrl+Shift+Z'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

export function createCommandRegistry(): CommandRegistry {
  const registry = new CommandRegistry();
  registry.register(cellSetCommand);
  registry.register(workbookDateSystemCommand);
  registry.register(dimensionSetCommand);
  registry.register(sheetInsertCommand);
  registry.register(sheetDeleteCommand);
  registry.register(sheetRenameCommand);
  for (const command of clipboardCommands()) registry.register(command);
  for (const command of formatCommands()) registry.register(command);
  registry.register(undoCommand);
  registry.register(redoCommand);
  return registry;
}
