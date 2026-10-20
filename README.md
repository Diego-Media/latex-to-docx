# latex-to-docx

[![npm](https://img.shields.io/npm/v/latex-to-docx)](https://www.npmjs.com/package/latex-to-docx) [![license](https://img.shields.io/npm/l/latex-to-docx)](LICENSE)

Convert LaTeX math into **native, editable Word equations** (OMML) for the [`docx`](https://www.npmjs.com/package/docx) library.

Built for the formulas that ChatGPT, Claude and Gemini write — fractions, roots, sums, integrals, limits,
`\left … \right` brackets, Greek letters, `\text{…}`, `aligned` derivations and `cases` — and it never throws:
anything it doesn't know is kept as plain text inside the equation.

It powers the [ChatGPT to Word converter](https://freeforeverconverter.com/chatgpt-to-word/) on Free Forever Converter,
and is explained in [How to copy ChatGPT math to Word](https://freeforeverconverter.com/chatgpt-math-to-word/).

**Tutorial with more examples** (inline math, `$…$` in text, browser download, aligned blocks):
[Convert LaTeX to editable Word equations in JavaScript](https://freeforeverconverter.com/guides/latex-to-word-equations-javascript/).

```bash
npm install latex-to-docx docx
```

## Usage

```js
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { latexToWordMath } from 'latex-to-docx';

const doc = new Document({
  sections: [{
    children: [
      new Paragraph({ children: [new TextRun('The quadratic formula:')] }),
      // One paragraph per equation: aligned blocks return one equation per line.
      ...latexToWordMath(String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`).map((eq) => new Paragraph({ children: [eq] })),
    ],
  }],
});

const blob = await Packer.toBlob(doc); // or Packer.toBuffer(doc) in Node
```

Pass the formula **without** its delimiters (`\( \)`, `\[ \]`, `$ $`, `$$ $$`). Inline equations can be put in a
paragraph next to `TextRun`s; display equations usually go in a paragraph of their own.

## API

### `latexToWordMath(tex: string): Math[]`

Returns an array of `docx` `Math` objects. Usually one; multi-line `aligned`, `gathered`, `align*`, `split` blocks
and a top-level `\\` give one equation per line.

## What is supported

| LaTeX | Word |
|---|---|
| `\frac{a}{b}`, `\dfrac`, `\tfrac` | fraction |
| `x^2`, `x_i`, `x_i^2` | super- / subscripts |
| `\sqrt{x}`, `\sqrt[n]{x}` | radicals |
| `\sum`, `\prod`, `\int`, `\oint` with limits | n-ary operators |
| `\lim_{x \to 0}` | limit with subscript |
| `\left( … \right)`, `[ ]`, `\{ \}`, `\langle \rangle` | growing brackets |
| `\binom{n}{k}` | binomial |
| `\sin`, `\log`, `\exp`, `\operatorname{…}` | upright function names |
| Greek letters, `\le`, `\ne`, `\approx`, `\infty`, `\to`, `\in`, … | Unicode math symbols |
| `\hat`, `\bar`, `\vec`, `\dot`, `\tilde` | accents |
| `\mathbb{R}` | double-struck letters |
| `\text{…}`, `\mathrm{…}`, `\mathbf{…}` | text |
| `aligned`, `gathered` | one equation per line |
| `pmatrix`, `bmatrix`, `cases` | rows kept inside brackets (as one line) |

Not (yet) supported: real 2-D matrices, `\overbrace` / `\underbrace`, colour, and custom macros.

## License

MIT © Diego Callegari — [freeforeverconverter.com](https://freeforeverconverter.com)
