// Builds a .docx with a set of typical AI formulas and checks the equation XML.
// Run in this folder: npm install && npm test   (writes test-output.docx to open in Word)
import { writeFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { strFromU8, unzipSync } from 'fflate';
import { latexToWordMath } from './index.js';

const formulas = [
  String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`,
  String.raw`\sum_{i=1}^{n} i = \frac{n(n+1)}{2}`,
  String.raw`\int_0^1 x^2 \, dx = \frac{1}{3}`,
  String.raw`\lim_{x \to 0} \frac{\sin x}{x} = 1`,
  String.raw`A = P\left(1 + \frac{r}{n}\right)^{nt}`,
  String.raw`\begin{aligned} a &= b + c \\ d &= e \end{aligned}`,
  String.raw`f(x) = \begin{cases} x & x > 0 \\ 0 & \text{otherwise} \end{cases}`,
  String.raw`\alpha \le \beta, \quad \mathbb{R}, \quad \hat{x}, \quad \binom{n}{k}`,
];

const children = [];
for (const tex of formulas) {
  const eqs = latexToWordMath(tex);
  assert.ok(eqs.length >= 1, tex);
  children.push(new Paragraph({ children: [new TextRun({ text: tex, font: 'Consolas', size: 18 })] }));
  for (const eq of eqs) children.push(new Paragraph({ children: [eq] }));
}
const blob = await Packer.toBlob(new Document({ sections: [{ children }] }));
const bytes = new Uint8Array(await blob.arrayBuffer());
const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
assert.equal((xml.match(/<m:oMath>/g) || []).length, formulas.length + 1); // aligned block = 2 equations
assert.ok(xml.includes('<m:f>'), 'fractions');
assert.ok(xml.includes('<m:rad>'), 'roots');
assert.ok(xml.includes('<m:nary>'), 'sums and integrals');
writeFileSync(new URL('./test-output.docx', import.meta.url), bytes);
console.log(`ok: ${formulas.length} formulas -> test-output.docx`);
