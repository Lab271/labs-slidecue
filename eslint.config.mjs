// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271

// `npm run lint` has never worked (#63): the script existed, but no config ever
// did, so eslint exited non-zero on every invocation and nothing noticed
// because no workflow ran it.
//
// Deliberately *not* built on `@eslint/js`. That package peers a single eslint
// major (`@eslint/js@10` requires `eslint ^10`), so depending on it would
// re-create the exact failure this config is meant to end: dependabot opens a
// solo `eslint` major bump, npm cannot resolve it against the pinned companion,
// and the PR dies at `npm ci` before a single test runs. Every package
// referenced below accepts eslint 9 and 10 alike, so nothing in this tree caps
// eslint's major any more. Keep it that way - prefer restating a core rule here
// over pulling in a major-locked shareable config.
//
// The file is .mjs because package.json has no `"type": "module"`.

import tsParser from '@typescript-eslint/parser';
import tsPlugin from '@typescript-eslint/eslint-plugin';
import reactHooks from 'eslint-plugin-react-hooks';

export default [
  {
    // Build output, dependencies, and the prebuilt web remote that ships as a
    // resource rather than as source. A config object with only `ignores` is
    // the flat-config equivalent of .eslintignore and applies globally.
    ignores: ['out/**', 'dist/**', 'node_modules/**', 'resources/remote/**'],
  },

  {
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
        ecmaFeatures: { jsx: true },
      },
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
    },
    rules: {
      // A subset of eslint's `recommended`, chosen for signal over volume: the
      // rules that catch real mistakes, rather than the ones restating what
      // `tsc --noEmit` (already a CI step) rejects anyway.
      'no-cond-assign': 'error',
      'no-constant-condition': ['error', { checkLoops: false }],
      'no-debugger': 'error',
      'no-dupe-else-if': 'error',
      'no-duplicate-case': 'error',
      'no-fallthrough': 'error',
      'no-irregular-whitespace': 'error',
      'no-self-compare': 'error',
      'no-template-curly-in-string': 'error',
      'no-unmodified-loop-condition': 'error',
      'no-unreachable-loop': 'error',
      'no-unsafe-finally': 'error',
      'no-unused-private-class-members': 'error',
      eqeqeq: ['error', 'always', { null: 'ignore' }],
      'prefer-const': 'error',
      'no-var': 'error',

      // 'warn', not 'error'. This flags seven pre-existing sites in the
      // PowerPoint automation backends where module-level state (currentSlide,
      // currentAnimationStep, localPresentationCopy) is read before an `await`
      // and written after it. They read as genuine interleaving hazards under
      // rapid remote input rather than false positives, but fixing them means
      // reworking the concurrency model of code that can only be exercised
      // against a real PowerPoint install - well outside the scope of turning
      // linting on. Left visible rather than disabled so the debt stays counted.
      'require-atomic-updates': 'warn',

      // Off on purpose, per typescript-eslint's own guidance: the compiler
      // already resolves every identifier, and keeping it on would mean
      // maintaining a hand-written globals list for three runtimes (electron
      // main, preload, browser) that would drift out of date.
      'no-undef': 'off',

      // The TS-aware replacement for core `no-unused-vars`, which mis-reports
      // type-only constructs. Leading underscores remain the opt-out.
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
      '@typescript-eslint/no-empty-object-type': 'error',
      '@typescript-eslint/no-misused-new': 'error',
      '@typescript-eslint/no-namespace': 'error',
      '@typescript-eslint/no-non-null-asserted-optional-chain': 'error',
      '@typescript-eslint/no-unsafe-declaration-merging': 'error',
      '@typescript-eslint/no-unsafe-function-type': 'error',
      '@typescript-eslint/no-wrapper-object-types': 'error',
      '@typescript-eslint/triple-slash-reference': 'error',

      // 'warn': the eleven current hits are all in the COM/winax and PowerShell
      // bridge layers, where the values really are untyped at the boundary.
      // Worth seeing, not worth blocking on.
      '@typescript-eslint/no-explicit-any': 'warn',
    },
  },

  {
    // Hooks rules only where hooks can appear. This pair is the part with
    // teeth: rules-of-hooks catches conditional hook calls, exhaustive-deps
    // catches stale closures in the useEffect-heavy renderer and remote UIs.
    files: ['src/renderer/**/*.{ts,tsx}', 'src/remote/**/*.{ts,tsx}'],
    plugins: { 'react-hooks': reactHooks },
    rules: {
      'react-hooks/rules-of-hooks': 'error',
      'react-hooks/exhaustive-deps': 'warn',
    },
  },

  {
    // Ambient declarations are all types; the value-oriented rules above have
    // nothing to say about them and no-unused-vars misfires on globals.
    files: ['**/*.d.ts'],
    rules: {
      '@typescript-eslint/no-unused-vars': 'off',
      '@typescript-eslint/no-explicit-any': 'off',
    },
  },
];
