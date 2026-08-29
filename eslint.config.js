const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    },
    rules: {
      // `void somePromise()` as a statement marks a deliberately un-awaited call
      // (best-effort run logging, fire-and-forget downloads). Banned as an expression.
      'no-void': ['warn', { allowAsStatement: true }]
    }
  }
];
