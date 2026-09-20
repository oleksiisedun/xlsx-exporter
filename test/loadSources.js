'use strict';

const fs = require('node:fs');
const path = require('node:path');

const SRC_DIR = path.join(__dirname, '..', 'src');

/**
 * Loads every `src/*.js` file the way Apps Script does: concatenated into one
 * shared scope, with no import system. Returns all top-level functions and
 * constants as one object, so tests can call them directly. Apps Script
 * globals (`SpreadsheetApp`, `Utilities`, ...) are not defined, so only code
 * that doesn't touch them can run — which is the pure logic worth testing.
 * @returns {Record<string, any>}
 */
function loadSources() {
  const source = fs.readdirSync(SRC_DIR)
    .filter((file) => file.endsWith('.js'))
    .sort()
    .map((file) => fs.readFileSync(path.join(SRC_DIR, file), 'utf8'))
    .join('\n');
  const names = [...source.matchAll(/^(?:function|const) (\w+)/gm)].map((match) => match[1]);
  return new Function(`${source}\nreturn { ${names.join(', ')} };`)();
}

module.exports = { loadSources };
