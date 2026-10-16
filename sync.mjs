// The converter's source of truth is the website's src/tools/latex-to-word.js; copy it in before packing.
import { copyFileSync } from 'node:fs';

copyFileSync(new URL('../../src/tools/latex-to-word.js', import.meta.url), new URL('./index.js', import.meta.url));
console.log('index.js synced from src/tools/latex-to-word.js');
