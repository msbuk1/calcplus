#!/usr/bin/env node
import { readFileSync } from 'node:fs';

const CONTRACT_ALLOWLIST = new Set([
  'GPL-3.0-only',
  'MIT',
  'ISC',
  'BSD-2-Clause',
  'BSD-3-Clause',
  'Apache-2.0',
  'MPL-2.0',
  'OFL-1.1',
  'PSF-2.0',
  '0BSD',
  'CC0-1.0',
]);

// Reviewed additions: these appear only in build/test tooling and are all
// GPL-3.0-only-compatible. The licensing contract's allowlist is the default;
// this list is deliberately narrow and needs a review note with each change.
const REVIEWED_ADDITIONS = new Map([
  ['Python-2.0', 'argparse (via eslint/js-yaml); GPL-compatible per FSF.'],
  ['BlueOak-1.0.0', 'minimatch 10 (via typescript-eslint); permissive, GPL-compatible.'],
]);

const packageJson = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
if (packageJson.license !== 'GPL-3.0-only') {
  console.error(
    `licence gate: package.json must declare GPL-3.0-only, found ${packageJson.license}`,
  );
  process.exit(1);
}

const lock = JSON.parse(readFileSync(new URL('../package-lock.json', import.meta.url), 'utf8'));
const failures = [];
const reviewedUsed = new Map();
let checked = 0;

for (const [path, entry] of Object.entries(lock.packages ?? {})) {
  if (path === '' || entry === undefined || entry.version === undefined) continue;
  if (!path.startsWith('node_modules/')) continue;
  checked += 1;
  const name = path.slice(path.lastIndexOf('node_modules/') + 'node_modules/'.length);
  const license = normalise(entry.license);
  if (license === null) {
    failures.push(`${name}@${entry.version}: no licence field`);
    continue;
  }
  if (!isAllowed(license)) {
    failures.push(`${name}@${entry.version}: ${license}`);
  }
  for (const [addition, reason] of REVIEWED_ADDITIONS) {
    if (license.includes(addition)) reviewedUsed.set(addition, reason);
  }
}

if (failures.length > 0) {
  console.error('licence gate: disallowed or unknown licences found:');
  for (const failure of failures) console.error(`  - ${failure}`);
  console.error(
    'Review the dependency or extend the allowlist in spec/licensing.md and this script.',
  );
  process.exit(1);
}

console.log(`licence gate: checked ${checked} packages; all licences are GPL-3.0-only compatible.`);
for (const [license, reason] of reviewedUsed) {
  console.log(`  reviewed addition: ${license} — ${reason}`);
}

function normalise(value) {
  if (typeof value === 'string') return value;
  if (value !== null && typeof value === 'object' && typeof value.type === 'string')
    return value.type;
  return null;
}

function isAllowed(expression) {
  // Simple SPDX expression handling for the licence fields npm produces:
  // parentheses, OR (any allowed choice is enough) and AND (all required).
  const tokens = expression.replaceAll('(', ' ').replaceAll(')', ' ').trim().split(/\s+/);
  let index = 0;
  const parseOr = () => {
    let allowed = parseAnd();
    while (tokens[index] === 'OR') {
      index += 1;
      allowed = parseAnd() || allowed;
    }
    return allowed;
  };
  const parseAnd = () => {
    let allowed = parsePrimary();
    while (tokens[index] === 'AND') {
      index += 1;
      allowed = parsePrimary() && allowed;
    }
    return allowed;
  };
  const parsePrimary = () => {
    const token = tokens[index];
    index += 1;
    if (token === undefined) return false;
    return CONTRACT_ALLOWLIST.has(token) || REVIEWED_ADDITIONS.has(token);
  };
  return parseOr() && index >= tokens.length;
}
