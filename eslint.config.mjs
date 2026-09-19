import js from '@eslint/js';
import jsdoc from 'eslint-plugin-jsdoc';

// Apps Script concatenates every src/*.js file into one global scope, so
// cross-file references and Apps Script globals are verified by `tsc`
// (npm run typecheck) rather than by `no-undef` here.
export default [
  js.configs.recommended,
  {
    files: ['src/**/*.js'],
    plugins: { jsdoc },
    languageOptions: { ecmaVersion: 2022, sourceType: 'script' },
    rules: {
      'no-undef': 'off',
      // Top-level functions are consumed from other files, so only flag unused locals.
      'no-unused-vars': ['error', { vars: 'local', args: 'after-used' }],
      'no-var': 'error',
      'prefer-const': 'error',
      'prefer-arrow-callback': 'error',
      eqeqeq: ['error', 'always'],
      'jsdoc/require-jsdoc': ['error', { publicOnly: false, require: { FunctionDeclaration: true } }],
      'jsdoc/require-param': 'error',
      'jsdoc/require-param-type': 'error',
      'jsdoc/check-param-names': 'error',
      'jsdoc/require-returns': 'error',
      'jsdoc/require-returns-type': 'error',
    },
  },
];
