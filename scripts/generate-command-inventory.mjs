#!/usr/bin/env node
/**
 * Generate docs/command-inventory.md from the feature matrix and the checked-in
 * command map (src/commands/command-inventory.json).
 *
 * The mapping is bidirectional: every IN matrix row must name at least one
 * command, and every command must name exactly one IN matrix row (the About
 * dialog is sourced from ui-chrome §4.1 under the "UI chrome" pseudo-row).
 * The script also loads the real `createCommandRegistry()` so the catalogue
 * cannot drift from the commands the runtime actually registers.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const matrixPath = path.join(root, '.scratch', 'calcplus', 'spec', 'feature-matrix.md');
const inventoryPath = path.join(root, 'src', 'commands', 'command-inventory.json');
const outputPath = path.join(root, 'docs', 'command-inventory.md');

const RIBBON_TABS = new Set([
  'Home',
  'Insert',
  'Page Layout',
  'Formulas',
  'Data',
  'Review',
  'View',
]);
const CONTROL_KINDS = new Set(['button', 'toggle', 'select', 'split', 'menu', 'launcher']);
const CONTEXT_REGIONS = new Set(['cell', 'rowHeader', 'columnHeader']);
const CONTEXT_REGION_LABELS = {
  cell: 'cell',
  rowHeader: 'row header',
  columnHeader: 'column header',
};
/** The only command allowed to use the ui-chrome pseudo-row instead of the feature matrix. */
const UI_CHROME_COMMANDS = new Set(['help.about']);
/** ui-chrome §1 pins the Quick Access Toolbar to these commands; ticket 42 adds Print. */
const QUICK_ACCESS_COMMANDS = new Set(['file.save', 'file.print', 'edit.undo', 'edit.redo']);
/**
 * The only command allowed inside a quick-access drop-down. Save carries Save
 * As on its split arrow; Save As never gets a quick-access placement of its own.
 */
const QUICK_ACCESS_MENU_COMMANDS = new Set(['file.saveAs']);
const COLOUR_PICKERS = new Set(['font', 'fill']);
/**
 * The reconciliation classification (ticket 36 item 3) every unimplemented
 * entry must carry: `not-a-command` for engine, file or view behaviour that no
 * command will ever hold, `v2-cut` for a command deferred on purpose.
 */
const UNIMPLEMENTED_REASONS = new Set(['not-a-command', 'v2-cut']);
const ERROR_LIMIT = 40;

const errors = [];

function fail(message) {
  if (errors.length < ERROR_LIMIT) errors.push(message);
}

function checkIcon(label, icon, where, iconNames) {
  if (typeof icon !== 'string' || icon.trim() === '') {
    fail(`${label}: ${where} must be a non-empty icon name.`);
    return;
  }
  if (!iconNames.has(icon)) {
    fail(`${label}: unknown icon "${icon}" in ${where}.`);
  }
}

function readMatrix() {
  const text = fs.readFileSync(matrixPath, 'utf8');
  const rows = [];
  let section = null;
  for (const line of text.split(/\r?\n/)) {
    const heading = /^##\s+(.+?)\s*$/.exec(line);
    if (heading !== null) {
      section = heading[1];
      continue;
    }
    if (section === null || !line.startsWith('|')) continue;
    const cells = line
      .split('|')
      .slice(1, -1)
      .map((cell) => cell.trim());
    if (cells.length < 3) continue;
    if (cells[0] === 'Feature' || /^:?-+/.test(cells[0])) continue;
    const statusCell = cells[1];
    const status = statusCell.startsWith('IN')
      ? 'IN'
      : statusCell.startsWith('V2')
        ? 'V2'
        : statusCell.startsWith('NEVER')
          ? 'NEVER'
          : null;
    if (status === null) continue;
    rows.push({
      section,
      feature: cells[0],
      v1: statusCell,
      status,
      order: rows.length,
    });
  }
  return rows;
}

function readInventory() {
  try {
    return JSON.parse(fs.readFileSync(inventoryPath, 'utf8'));
  } catch (error) {
    console.error(`Cannot read ${path.relative(root, inventoryPath)}: ${error.message}`);
    process.exit(1);
  }
}

/** The map may author one context-menu placement or a list (Insert Cells has two). */
function contextMenuPlacements(entry) {
  if (entry.contextMenu === undefined) return [];
  return Array.isArray(entry.contextMenu) ? entry.contextMenu : [entry.contextMenu];
}

/**
 * Load the actual TypeScript registry through Vite's SSR loader rather than
 * scraping `id:` literals from the command files. The returned set is exactly
 * what `createCommandRegistry()` registers in the application. The icon names
 * come from the same loader so the map and the artwork cannot drift.
 */
async function runtimeCatalog() {
  const server = await createServer({
    root,
    configFile: path.join(root, 'vite.config.ts'),
    appType: 'custom',
    logLevel: 'silent',
    server: { middlewareMode: true },
  });
  try {
    const registryModule = await server.ssrLoadModule('/src/commands/registry.ts');
    const createCommandRegistry = registryModule.createCommandRegistry;
    if (typeof createCommandRegistry !== 'function') {
      throw new Error('src/commands/registry.ts does not export createCommandRegistry().');
    }
    const iconModule = await server.ssrLoadModule('/src/chrome/ribbon-icons.ts');
    const iconNames = iconModule.RIBBON_ICON_NAMES;
    if (!Array.isArray(iconNames) || iconNames.length === 0) {
      throw new Error('src/chrome/ribbon-icons.ts does not export RIBBON_ICON_NAMES.');
    }
    const catalogModule = await server.ssrLoadModule('/src/engine/function-catalog.ts');
    const categoryOptions = catalogModule.FUNCTION_CATEGORY_OPTIONS;
    if (!Array.isArray(categoryOptions) || categoryOptions.length === 0) {
      throw new Error('src/engine/function-catalog.ts does not export FUNCTION_CATEGORY_OPTIONS.');
    }
    return {
      registered: new Set(
        createCommandRegistry()
          .list()
          .map((command) => command.id),
      ),
      // A state-dependent label (ticket 43) is only meaningful on a command
      // that reports a checked state.
      checkedCommands: new Set(
        createCommandRegistry()
          .list()
          .filter((command) => typeof command.isChecked === 'function')
          .map((command) => command.id),
      ),
      iconNames: new Set(iconNames),
      // "All" is dialog-only; a ribbon category menu names one real category.
      functionCategories: new Set(categoryOptions.filter((category) => category !== 'All')),
    };
  } finally {
    await server.close();
  }
}

async function main() {
  const matrix = readMatrix();
  const inventory = readInventory();
  if (!Array.isArray(inventory)) {
    console.error('command-inventory.json must contain an array.');
    process.exit(1);
  }
  const { registered, checkedCommands, iconNames, functionCategories } = await runtimeCatalog();
  const matrixByKey = new Map(matrix.map((row) => [`${row.section}::${row.feature}`, row]));
  const byMatrixRow = new Map();
  const seenIds = new Set();
  const contextGroups = new Map();

  for (const entry of inventory) {
    const label = typeof entry.id === 'string' ? entry.id : '<missing id>';
    for (const field of [
      'id',
      'matrixSection',
      'matrixFeature',
      'label',
      'category',
      'enablement',
    ]) {
      if (typeof entry[field] !== 'string' || entry[field].trim() === '') {
        fail(`${label}: missing or empty "${field}".`);
      }
    }
    if (typeof entry.implemented !== 'boolean') fail(`${label}: "implemented" must be boolean.`);
    if (entry.implemented === true) {
      if (entry.unimplementedReason !== undefined) {
        fail(`${label}: an implemented command must not carry "unimplementedReason".`);
      }
    } else if (entry.unimplementedReason === undefined) {
      fail(
        `${label}: an unimplemented entry must classify itself with "unimplementedReason" (${[...UNIMPLEMENTED_REASONS].join(' or ')}).`,
      );
    } else if (!UNIMPLEMENTED_REASONS.has(entry.unimplementedReason)) {
      fail(`${label}: unknown "unimplementedReason": ${entry.unimplementedReason}.`);
    }
    if (!Array.isArray(entry.shortcut)) fail(`${label}: "shortcut" must be an array.`);
    if (typeof entry.id === 'string') {
      if (seenIds.has(entry.id)) fail(`${entry.id}: duplicate command id.`);
      seenIds.add(entry.id);
    }
    const ribbonPlacements = entry.ribbon ?? [];
    if (ribbonPlacements.length > 0) {
      if (entry.icon === undefined) {
        fail(`${label}: a command with a ribbon placement needs an "icon".`);
      } else {
        checkIcon(label, entry.icon, 'command icon', iconNames);
      }
    } else if (entry.icon !== undefined) {
      checkIcon(label, entry.icon, 'command icon', iconNames);
    }
    const key = `${entry.matrixSection}::${entry.matrixFeature}`;
    const row = matrixByKey.get(key);
    if (entry.matrixSection === 'UI chrome') {
      // Only the About dialog is sourced from ui-chrome §4.1; every other
      // command must trace to a feature-matrix row.
      if (!UI_CHROME_COMMANDS.has(entry.id)) {
        fail(`${label}: only help.about may use the "UI chrome" pseudo-row.`);
      }
    } else if (row === undefined) {
      fail(`${label}: matrix row not found: ${entry.matrixSection} / ${entry.matrixFeature}`);
    } else if (row.status !== 'IN') {
      fail(`${label}: matrix row is ${row.status}, not IN: ${entry.matrixFeature}`);
    }
    if (!byMatrixRow.has(key)) byMatrixRow.set(key, []);
    byMatrixRow.get(key).push(entry);

    for (const placement of entry.ribbon ?? []) {
      if (!RIBBON_TABS.has(placement.tab)) fail(`${label}: unknown ribbon tab ${placement.tab}.`);
      if (typeof placement.group !== 'string' || placement.group.trim() === '') {
        fail(`${label}: ribbon placement has no group.`);
      }
      if (!CONTROL_KINDS.has(placement.kind)) {
        fail(`${label}: unknown ribbon control kind ${placement.kind}.`);
      }
      if (placement.iconOnly !== undefined && typeof placement.iconOnly !== 'boolean') {
        fail(`${label}: ribbon placement "iconOnly" must be boolean.`);
      }
      if (placement.kind !== 'launcher' && placement.menuOnly !== true) {
        if (placement.icon === undefined) {
          fail(`${label}: ribbon placement ${placement.tab} → ${placement.group} needs an "icon".`);
        } else {
          checkIcon(label, placement.icon, 'ribbon icon', iconNames);
        }
      } else if (placement.icon !== undefined) {
        checkIcon(label, placement.icon, 'ribbon icon', iconNames);
      }
      if (placement.kind === 'select' && !Array.isArray(placement.options)) {
        fail(`${label}: select control ${placement.controlId ?? placement.label} has no options.`);
      }
      if (placement.row !== undefined && (!Number.isInteger(placement.row) || placement.row < 1)) {
        fail(`${label}: ribbon placement row must be a positive integer.`);
      }
      if (
        placement.rowSpan !== undefined &&
        (!Number.isInteger(placement.rowSpan) || placement.rowSpan < 1)
      ) {
        fail(`${label}: ribbon placement rowSpan must be a positive integer.`);
      }
      if (placement.disabled !== undefined && typeof placement.disabled !== 'boolean') {
        fail(`${label}: ribbon placement "disabled" must be boolean.`);
      }
      if (
        placement.menuHeading !== undefined &&
        (typeof placement.menuHeading !== 'string' || placement.menuHeading.trim() === '')
      ) {
        fail(`${label}: ribbon placement "menuHeading" must be a non-empty string.`);
      }
      if (placement.picker !== undefined) {
        if (!COLOUR_PICKERS.has(placement.picker)) {
          fail(`${label}: unknown colour picker "${placement.picker}".`);
        }
        if (placement.menu !== undefined) {
          // The ruled direction replaces the fixed preset list with the picker.
          fail(`${label}: a colour picker placement must not also author a preset menu.`);
        }
        if (placement.args === undefined) {
          fail(`${label}: a colour picker placement needs args for its main button.`);
        }
      }
      if (placement.functionCategory !== undefined) {
        // A Function Library menu names its catalogue category; the runtime
        // builds the items, so the map never copies a function list.
        if (!functionCategories.has(placement.functionCategory)) {
          fail(`${label}: unknown function category "${placement.functionCategory}".`);
        }
        if (placement.kind !== 'menu') {
          fail(`${label}: a function-category control must be a menu.`);
        }
        if (placement.menu !== undefined) {
          fail(`${label}: a function-category control builds its menu from the catalogue.`);
        }
      }
      if (placement.disabled === true && entry.implemented === true) {
        // The present-but-disabled carve-out (ui-chrome §4.1, ticket 27) is for
        // planned, render-only IN rows only; a live command is never greyed.
        fail(`${label}: an implemented command cannot carry a disabled ribbon placement.`);
      }
      if (placement.disabled === true && (placement.title ?? '').trim() === '') {
        // The carve-out requires a tooltip that names the v2 limitation.
        fail(`${label}: a disabled ribbon placement needs a tooltip (title).`);
      }
      if (placement.activeLabel !== undefined) {
        // The checked-state label (ticket 43): the entry reads one label while
        // the command reports state and another while it does not.
        const defaultLabel = placement.label ?? entry.label;
        if (typeof placement.activeLabel !== 'string' || placement.activeLabel.trim() === '') {
          fail(`${label}: ribbon placement "activeLabel" must be a non-empty string.`);
        } else if (placement.activeLabel === defaultLabel) {
          fail(`${label}: ribbon placement "activeLabel" must differ from its label.`);
        }
        if (!checkedCommands.has(entry.id)) {
          fail(`${label}: an "activeLabel" needs a command with a checked state (isChecked).`);
        }
      }
      for (const option of placement.options ?? []) {
        if (typeof option.label !== 'string' || option.label.trim() === '') {
          fail(`${label}: select option without a label.`);
        }
      }
      for (const menuItem of placement.menu ?? []) {
        if (typeof menuItem.label !== 'string' || typeof menuItem.commandId !== 'string') {
          fail(`${label}: split menu item needs label and commandId.`);
        } else if (!inventory.some((candidate) => candidate.id === menuItem.commandId)) {
          fail(`${label}: split menu item names an unknown command: ${menuItem.commandId}`);
        }
        if (menuItem.shortcut !== undefined && typeof menuItem.shortcut !== 'string') {
          fail(`${label}: menu item "${menuItem.label}" has a non-string shortcut.`);
        }
        if (menuItem.icon !== undefined) {
          checkIcon(label, menuItem.icon, `menu item "${menuItem.label}" icon`, iconNames);
        }
      }
    }
    for (const menu of entry.menu ?? []) {
      if (typeof menu.menu !== 'string' || menu.menu.trim() === '') {
        fail(`${label}: menu placement has no menu name.`);
      }
    }
    for (const placement of contextMenuPlacements(entry)) {
      if (placement.context !== undefined && !CONTEXT_REGIONS.has(placement.context)) {
        fail(`${label}: unknown context-menu region ${placement.context}.`);
      }
      if (
        placement.order !== undefined &&
        (!Number.isInteger(placement.order) || placement.order < 0)
      ) {
        fail(`${label}: context-menu order must be a non-negative integer.`);
      }
      if (
        placement.label !== undefined &&
        (typeof placement.label !== 'string' || placement.label.trim() === '')
      ) {
        fail(`${label}: context-menu label must be a non-empty string.`);
      }
      if (placement.shortcut !== undefined && typeof placement.shortcut !== 'string') {
        fail(`${label}: context-menu shortcut must be a string.`);
      }
      if (placement.group !== undefined) {
        const group = placement.group;
        if (typeof group.id !== 'string' || group.id.trim() === '') {
          fail(`${label}: context-menu group needs a non-empty id.`);
        }
        if (typeof group.label !== 'string' || group.label.trim() === '') {
          fail(`${label}: context-menu group needs a non-empty label.`);
        }
        if (!Number.isInteger(group.order) || group.order < 0) {
          fail(`${label}: context-menu group order must be a non-negative integer.`);
        }
        const seen = contextGroups.get(group.id);
        if (seen !== undefined && (seen.label !== group.label || seen.order !== group.order)) {
          fail(`${label}: context-menu group "${group.id}" disagrees with its first definition.`);
        } else if (seen === undefined) {
          contextGroups.set(group.id, { label: group.label, order: group.order });
        }
      }
    }
    if (entry.compact !== undefined) {
      if (!['button', 'toggle'].includes(entry.compact.kind)) {
        fail(`${label}: compact control must be a button or toggle.`);
      }
      if (entry.compact.iconOnly !== undefined && typeof entry.compact.iconOnly !== 'boolean') {
        fail(`${label}: compact placement "iconOnly" must be boolean.`);
      }
      const compactIcon = entry.compact.icon ?? entry.icon;
      if (compactIcon === undefined) {
        fail(`${label}: compact placement needs an "icon".`);
      } else {
        checkIcon(label, compactIcon, 'compact icon', iconNames);
      }
    }
    if (entry.quickAccess !== undefined) {
      const placement = entry.quickAccess;
      if (typeof placement !== 'object' || Array.isArray(placement)) {
        fail(`${label}: "quickAccess" must be an object.`);
      } else {
        if (!QUICK_ACCESS_COMMANDS.has(entry.id)) {
          fail(`${label}: the Quick Access Toolbar holds Save, Print, Undo and Redo only.`);
        }
        if (
          placement.order !== undefined &&
          (!Number.isInteger(placement.order) || placement.order < 0)
        ) {
          fail(`${label}: quick-access order must be a non-negative integer.`);
        }
        const quickAccessIcon = placement.icon ?? entry.icon;
        if (quickAccessIcon === undefined) {
          fail(`${label}: a quick-access placement needs an "icon".`);
        } else {
          checkIcon(label, quickAccessIcon, 'quick-access icon', iconNames);
        }
        if (placement.menu !== undefined) {
          if (!Array.isArray(placement.menu)) {
            fail(`${label}: "quickAccess.menu" must be an array.`);
          } else {
            for (const menuItem of placement.menu) {
              if (typeof menuItem.label !== 'string' || typeof menuItem.commandId !== 'string') {
                fail(`${label}: quick-access menu item needs label and commandId.`);
                continue;
              }
              if (!inventory.some((candidate) => candidate.id === menuItem.commandId)) {
                fail(
                  `${label}: quick-access menu item names an unknown command: ${menuItem.commandId}`,
                );
              }
              if (!QUICK_ACCESS_MENU_COMMANDS.has(menuItem.commandId)) {
                fail(
                  `${label}: only Save As may sit in a quick-access drop-down: ${menuItem.commandId}`,
                );
              }
              if (menuItem.shortcut !== undefined && typeof menuItem.shortcut !== 'string') {
                fail(
                  `${label}: quick-access menu item "${menuItem.label}" has a non-string shortcut.`,
                );
              }
              if (menuItem.icon !== undefined) {
                checkIcon(
                  label,
                  menuItem.icon,
                  `quick-access menu item "${menuItem.label}" icon`,
                  iconNames,
                );
              }
            }
          }
        }
      }
    }
  }

  for (const id of QUICK_ACCESS_COMMANDS) {
    const entry = inventory.find((candidate) => candidate.id === id);
    if (entry?.quickAccess === undefined) {
      fail(`Quick-access command is missing its "quickAccess" placement: ${id}`);
    }
  }

  for (const row of matrix) {
    if (row.status !== 'IN') continue;
    const key = `${row.section}::${row.feature}`;
    if (!byMatrixRow.has(key)) {
      fail(`IN matrix row has no command: ${row.section} / ${row.feature}`);
    }
  }

  for (const id of registered) {
    if (!seenIds.has(id)) fail(`Runtime command is not in the inventory: ${id}`);
  }
  for (const entry of inventory) {
    const inRegistry = registered.has(entry.id);
    if (entry.implemented === true && !inRegistry) {
      fail(`Inventory says implemented, but the registry has no command: ${entry.id}`);
    } else if (entry.implemented === false && inRegistry) {
      fail(`Registry registers a planned command; inventory says not implemented: ${entry.id}`);
    }
  }

  if (errors.length > 0) {
    console.error(`Command inventory check failed (${errors.length} problem(s)):`);
    for (const error of errors) console.error(`- ${error}`);
    if (errors.length === ERROR_LIMIT) console.error('- ... more problems omitted.');
    process.exit(1);
  }

  const rowOrder = new Map(matrix.map((row) => [`${row.section}::${row.feature}`, row.order]));
  const sorted = [...inventory].sort((left, right) => {
    const leftKey = `${left.matrixSection}::${left.matrixFeature}`;
    const rightKey = `${right.matrixSection}::${right.matrixFeature}`;
    const leftOrder = rowOrder.get(leftKey) ?? Number.MAX_SAFE_INTEGER;
    const rightOrder = rowOrder.get(rightKey) ?? Number.MAX_SAFE_INTEGER;
    return (
      leftOrder - rightOrder ||
      left.matrixSection.localeCompare(right.matrixSection) ||
      left.matrixFeature.localeCompare(right.matrixFeature) ||
      left.id.localeCompare(right.id)
    );
  });

  const implemented = sorted.filter((entry) => entry.implemented).length;
  const cut = sorted.filter((entry) => entry.unimplementedReason === 'v2-cut').length;
  const matrixIn = matrix.filter((row) => row.status === 'IN').length;
  const outRows = matrix.filter((row) => row.status !== 'IN');

  const lines = [];
  lines.push('# calcplus v1 — command inventory');
  lines.push('');
  lines.push(
    'Generated by `npm run inventory` from [the feature matrix](../.scratch/calcplus/spec/feature-matrix.md)',
  );
  lines.push(
    'and [the checked-in command map](../src/commands/command-inventory.json). Do not edit by hand.',
  );
  lines.push('');
  lines.push(
    'Every feature-matrix **IN** row names at least one command; every command names one matrix row.',
  );
  lines.push(
    'The About dialog is the only chrome-contract command without a feature-matrix row; it is sourced',
  );
  lines.push('from `ui-chrome` §4.1 and listed under **UI chrome**.');
  lines.push('');
  lines.push('## Summary');
  lines.push('');
  lines.push(`- Feature-matrix IN rows: ${matrixIn}`);
  lines.push(`- Inventory commands: ${sorted.length}`);
  lines.push(`- Implemented in this build: ${implemented}`);
  lines.push(`- Planned for a later slice: ${cut} (deliberate v2 cuts)`);
  lines.push(
    `- No command by design: ${sorted.length - implemented - cut} (engine, file or view behaviour)`,
  );
  lines.push(`- Out-of-scope rows (V2 or NEVER): ${outRows.length}`);
  lines.push('');
  lines.push(
    'OUT rows are absent from the chrome: no control is rendered and nothing is greyed out to imply',
  );
  lines.push('that the command exists.');
  lines.push('');
  lines.push(
    'One ruled exception (ui-chrome §4.1, ticket 27): a matrix row that is IN as render-only with',
  );
  lines.push(
    'V2 authoring ships its ribbon control present-but-disabled, with a tooltip naming the v2',
  );
  lines.push('limitation. `format.conditionalRender` is the v1 instance.');
  lines.push('');
  lines.push('## Commands');
  lines.push('');
  lines.push(
    'The Icon column names the original SVG artwork in `src/chrome/ribbon-icons.ts`. A command with',
  );
  lines.push(
    'several ribbon controls lists each distinct icon, including the icons of its split-button menu',
  );
  lines.push('items; the command-level icon is the compact and quick-access fallback.');
  lines.push(
    'The Quick access column names the commands pinned above the ribbon (ui-chrome §1, ticket 42).',
  );
  lines.push(
    'The toolbar holds Save, Print, Undo and Redo; Save As stays in the File menu and on the Save',
  );
  lines.push("button's split drop-down.");
  lines.push(
    'A ribbon control whose label changes with the command’s checked state (ticket 43, Freeze',
  );
  lines.push('Panes → Unfreeze Panes) lists the checked label in the Ribbon column.');
  lines.push(
    'The Status column explains every unimplemented entry: **Planned (v2 cut)** is a command',
  );
  lines.push(
    'deferred on purpose; **Not a command (by design)** is engine, file or view behaviour that no',
  );
  lines.push(
    'command will ever carry. The map records the reason per entry (`unimplementedReason`).',
  );
  lines.push('');
  lines.push(
    '| Command ID | Matrix row | Label | Icon | Quick access | Ribbon tab → group | Menu | Shortcut | Enablement | Status |',
  );
  lines.push('|---|---|---|---|---|---|---|---|---|---|');
  for (const entry of sorted) {
    // Several controls of one command can sit in the same band (Font family,
    // size and colour are all `format.font`); the map records each placement,
    // and the register lists the distinct tab → group addresses once. Notes
    // such as the colour picker merge into the address they belong to.
    const ribbonAddresses = new Map();
    for (const placement of entry.ribbon ?? []) {
      const address = `${placement.tab} → ${placement.group}`;
      const notes = ribbonAddresses.get(address) ?? new Set();
      if (placement.menuOnly === true) notes.add('menu');
      if (placement.kind === 'launcher') notes.add('launcher');
      if (placement.picker !== undefined) notes.add(`${placement.picker} colour picker`);
      if (placement.disabled === true) notes.add('present, disabled');
      if (placement.activeLabel !== undefined) {
        notes.add(`checked label: ${placement.activeLabel}`);
      }
      ribbonAddresses.set(address, notes);
    }
    const ribbon =
      [...ribbonAddresses.entries()]
        .map(([address, notes]) =>
          notes.size === 0 ? address : `${address} (${[...notes].join(', ')})`,
        )
        .join('<br>') || '—';
    const menuParts = (entry.menu ?? []).map((placement) => placement.menu);
    for (const placement of contextMenuPlacements(entry)) {
      const region = placement.context ?? 'cell';
      menuParts.push(
        region !== 'cell'
          ? `Context menu (${CONTEXT_REGION_LABELS[region]})`
          : placement.group === undefined
            ? 'Context menu'
            : `Context menu (${placement.group.label})`,
      );
    }
    const menu = [...new Set(menuParts)].join('<br>') || '—';
    const shortcut = entry.shortcut.join(', ') || '—';
    const icons = [
      ...new Set([
        ...(typeof entry.icon === 'string' ? [entry.icon] : []),
        ...(entry.ribbon ?? []).flatMap((placement) => [
          ...(typeof placement.icon === 'string' ? [placement.icon] : []),
          ...(placement.menu ?? []).flatMap((item) =>
            typeof item.icon === 'string' ? [item.icon] : [],
          ),
        ]),
        ...(typeof entry.compact?.icon === 'string' ? [entry.compact.icon] : []),
        ...(typeof entry.quickAccess?.icon === 'string' ? [entry.quickAccess.icon] : []),
        ...(entry.quickAccess?.menu ?? []).flatMap((item) =>
          typeof item.icon === 'string' ? [item.icon] : [],
        ),
      ]),
    ];
    const icon = icons.map((name) => `\`${name}\``).join(', ') || '—';
    const quickAccessMenu = (entry.quickAccess?.menu ?? []).map((item) => item.label);
    const quickAccess =
      entry.quickAccess === undefined
        ? '—'
        : `Yes${entry.quickAccess.order === undefined ? '' : ` (order ${entry.quickAccess.order})`}${
            quickAccessMenu.length === 0 ? '' : `; menu: ${quickAccessMenu.join(', ')}`
          }`;
    const status = entry.implemented
      ? 'Implemented'
      : entry.unimplementedReason === 'v2-cut'
        ? 'Planned (v2 cut)'
        : 'Not a command (by design)';
    lines.push(
      `| \`${entry.id}\` | ${escapeCell(`${entry.matrixSection} — ${entry.matrixFeature}`)} | ${escapeCell(entry.label)} | ${escapeCell(icon)} | ${escapeCell(quickAccess)} | ${escapeCell(ribbon)} | ${escapeCell(menu)} | ${escapeCell(shortcut)} | ${escapeCell(entry.enablement)} | ${status} |`,
    );
  }
  lines.push('');
  lines.push('## Out-of-scope rows');
  lines.push('');
  lines.push(
    'These rows stay hidden in v1. The generator rejects an inventory entry that names one of them.',
  );
  lines.push('');
  lines.push('| Section | Feature | Matrix state | Chrome treatment |');
  lines.push('|---|---|---|---|');
  for (const row of outRows) {
    lines.push(
      `| ${escapeCell(row.section)} | ${escapeCell(row.feature)} | ${row.v1} | Hidden; no control rendered |`,
    );
  }
  lines.push('');

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
  console.log(
    `Wrote ${path.relative(root, outputPath)}: ${sorted.length} commands, ${matrixIn} IN rows, ${outRows.length} hidden OUT rows.`,
  );
}

function escapeCell(value) {
  return String(value).replaceAll('|', '\\|').replace(/\r?\n/g, ' ');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
