import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';

const modelImportMessage =
  'ADR 0007: the grid renderer reads a sheet projection, never the workbook model.';

export default tseslint.config(
  {
    ignores: ['**/dist/**', 'node_modules/**', 'fixtures/**', '.scratch/**', 'docs/**'],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.node,
      },
    },
  },
  {
    files: ['src/grid/**/*.{ts,tsx}'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['**/model/**', '@model/**', '@model'],
              message: modelImportMessage,
            },
          ],
        },
      ],
    },
  },
  {
    files: ['src/**/*.{ts,tsx}', 'tests/**/*.{ts,tsx}'],
    ignores: ['src/ooxml/**'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['**/ooxml/patch/**', '@ooxml/patch/**', '**/ooxml/zip', '@ooxml/zip'],
              message:
                'ADR 0006: only the OOXML layer may patch or write part bytes. Call the open/save facade.',
            },
          ],
        },
      ],
    },
  },
  {
    files: ['scripts/**/*.mjs', '*.config.ts'],
    languageOptions: {
      globals: {
        ...globals.node,
      },
    },
  },
);
