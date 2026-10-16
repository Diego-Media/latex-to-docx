// LaTeX (as written by ChatGPT, Claude, Gemini …) -> native, editable Word equations (OMML via docx).
// Covers what AI answers actually use: fractions, powers/indices, roots, sums/products/integrals/limits,
// \left…\right brackets, functions, Greek letters and common symbols, \text, simple matrices/aligned rows.
// Anything unknown degrades to plain text instead of failing.
import {
  Math as WordMath,
  MathCurlyBrackets,
  MathFraction,
  MathIntegral,
  MathRadical,
  MathRoundBrackets,
  MathRun,
  MathSquareBrackets,
  MathSubScript,
  MathSubSuperScript,
  MathSum,
  MathSuperScript,
  MathAngledBrackets,
} from 'docx';

const SYMBOLS = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ',
  vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ',
  sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'ϕ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  infty: '∞', partial: '∂', nabla: '∇', pm: '±', mp: '∓', times: '×', div: '÷', cdot: '⋅', ast: '∗', star: '⋆',
  circ: '∘', bullet: '∙', leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈', equiv: '≡', sim: '∼',
  simeq: '≃', cong: '≅', propto: '∝', ll: '≪', gg: '≫', in: '∈', notin: '∉', ni: '∋', subset: '⊂', supset: '⊃',
  subseteq: '⊆', supseteq: '⊇', cup: '∪', cap: '∩', setminus: '∖', emptyset: '∅', varnothing: '∅', forall: '∀',
  exists: '∃', neg: '¬', lnot: '¬', land: '∧', wedge: '∧', lor: '∨', vee: '∨', to: '→', rightarrow: '→', leftarrow: '←',
  gets: '←', Rightarrow: '⇒', Leftarrow: '⇐', implies: '⟹', iff: '⟺', leftrightarrow: '↔', Leftrightarrow: '⇔',
  mapsto: '↦', uparrow: '↑', downarrow: '↓', longrightarrow: '⟶', ldots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱',
  dots: '…', prime: '′', degree: '°', angle: '∠', perp: '⊥', parallel: '∥', mid: '∣', hbar: 'ℏ', ell: 'ℓ', Re: 'ℜ',
  Im: 'ℑ', aleph: 'ℵ', therefore: '∴', because: '∵', triangle: '△', square: '□', checkmark: '✓', lfloor: '⌊',
  rfloor: '⌋', lceil: '⌈', rceil: '⌉', langle: '⟨', rangle: '⟩', vert: '|', Vert: '‖', '|': '‖', '{': '{', '}': '}',
  '%': '%', '$': '$', '#': '#', '&': '&', _: '_', ',': ' ', ';': ' ', ':': ' ', '!': '', quad: ' ',
  qquad: '  ', ' ': ' ', '\\': '\n', displaystyle: '', textstyle: '', limits: '', nolimits: '',
};
const FUNCTIONS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh',
  'log', 'ln', 'lg', 'exp', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'deg', 'gcd', 'arg', 'Pr', 'mod', 'bmod']);
const BLACKBOARD = { R: 'ℝ', N: 'ℕ', Z: 'ℤ', Q: 'ℚ', C: 'ℂ', P: 'ℙ', E: '𝔼' };
const ACCENTS = { hat: '̂', bar: '̅', overline: '̅', vec: '⃗', dot: '̇', ddot: '̈', tilde: '̃', widehat: '̂' };
const NARY = { sum: 'sum', prod: 'prod', coprod: 'prod', int: 'int', iint: 'int', iiint: 'int', oint: 'int', bigcup: 'sum', bigcap: 'sum' };
const NARY_CHAR = { prod: '∏', coprod: '∐', iint: '∬', iiint: '∭', oint: '∮', bigcup: '⋃', bigcap: '⋂' };

function tokenize(src) {
  const tokens = [];
  for (let i = 0; i < src.length; ) {
    const c = src[i];
    if (c === '\\') {
      const word = src.slice(i + 1).match(/^[A-Za-z]+/);
      if (word) {
        tokens.push({ t: 'cmd', v: word[0] });
        i += 1 + word[0].length;
      } else {
        tokens.push({ t: 'cmd', v: src[i + 1] ?? '' });
        i += 2;
      }
    } else if (/\s/.test(c)) i++;
    else {
      tokens.push({ t: c === '{' || c === '}' || c === '^' || c === '_' || c === '&' || c === '[' || c === ']' ? c : 'char', v: c });
      i++;
    }
  }
  return tokens;
}

class Parser {
  constructor(src) {
    this.tokens = tokenize(src);
    this.pos = 0;
  }
  peek() {
    return this.tokens[this.pos];
  }
  next() {
    return this.tokens[this.pos++];
  }
  /** Parse until a terminator: '}' , ']' (when asked), \right, \end, & or end of input. */
  expression(stopAt = new Set()) {
    const out = [];
    while (this.pos < this.tokens.length) {
      const tok = this.peek();
      if (tok.t === '}' || (tok.t === ']' && stopAt.has(']')) || tok.t === '&') break;
      if (tok.t === 'cmd' && (tok.v === 'right' || tok.v === 'end' || tok.v === '\\')) break;
      if (tok.t === '^' || tok.t === '_') {
        this.scripts(out);
        continue;
      }
      const atom = this.atom(out);
      if (atom) out.push(...[].concat(atom).filter(Boolean));
    }
    return out;
  }
  group() {
    // A {...} group or a single token as argument.
    const tok = this.peek();
    if (!tok) return [];
    if (tok.t === '{') {
      this.next();
      const inner = this.expression();
      if (this.peek()?.t === '}') this.next();
      return inner;
    }
    const atom = this.atom([]);
    return atom ? [].concat(atom) : [];
  }
  textGroup() {
    // Raw text of a {...} group (for \text, \mathrm …).
    if (this.peek()?.t !== '{') return this.next()?.v ?? '';
    this.next();
    let depth = 1;
    let text = '';
    while (this.pos < this.tokens.length) {
      const tok = this.next();
      if (tok.t === '{') depth++;
      if (tok.t === '}' && --depth === 0) break;
      text += tok.t === 'cmd' ? (SYMBOLS[tok.v] ?? tok.v) : tok.v;
    }
    return text;
  }
  scripts(out) {
    let sub = null;
    let sup = null;
    while (this.peek()?.t === '^' || this.peek()?.t === '_') {
      const kind = this.next().t;
      const arg = this.group();
      if (kind === '^') sup = arg;
      else sub = arg;
    }
    const base = out.length ? [out.pop()] : [new MathRun('')];
    const nary = base[0]?.__nary;
    if (nary) return out.push(...[].concat(this.nary(nary, sub, sup)));
    if (sub && sup) out.push(new MathSubSuperScript({ children: base, subScript: sub, superScript: sup }));
    else if (sup) out.push(new MathSuperScript({ children: base, superScript: sup }));
    else out.push(new MathSubScript({ children: base, subScript: sub }));
  }
  nary(name, sub, sup) {
    // The rest of the current term is the body of the sum/integral.
    const body = this.expression();
    const opts = { children: body.length ? body : [new MathRun('')], ...(sub ? { subScript: sub } : {}), ...(sup ? { superScript: sup } : {}) };
    if (NARY[name] === 'int' && name === 'int') return new MathIntegral(opts);
    if (name === 'sum') return new MathSum(opts);
    // No dedicated docx class: build the operator with limits as a sub/superscript.
    const op = new MathRun(NARY_CHAR[name] || '∑');
    const withLimits = sub && sup ? new MathSubSuperScript({ children: [op], subScript: sub, superScript: sup })
      : sup ? new MathSuperScript({ children: [op], superScript: sup })
      : sub ? new MathSubScript({ children: [op], subScript: sub }) : op;
    return [withLimits, ...body];
  }
  bracket(open, inner, close) {
    const pair = open + close;
    if (pair === '()') return new MathRoundBrackets({ children: inner });
    if (pair === '[]') return new MathSquareBrackets({ children: inner });
    if (pair === '{}') return new MathCurlyBrackets({ children: inner });
    if (pair === '⟨⟩') return new MathAngledBrackets({ children: inner });
    return [new MathRun(open), ...inner, new MathRun(close)];
  }
  delimiter() {
    const tok = this.next();
    if (!tok) return '';
    if (tok.t === 'cmd') return { langle: '⟨', rangle: '⟩', '{': '{', '}': '}', '|': '‖', vert: '|', Vert: '‖', lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉' }[tok.v] ?? '';
    return tok.v === '.' ? '' : tok.v;
  }
  environment(name) {
    // matrix-like and aligned environments: rows separated by \\, cells by &.
    const rows = [[]];
    while (this.pos < this.tokens.length) {
      const tok = this.peek();
      if (tok.t === 'cmd' && tok.v === 'end') {
        this.next();
        this.textGroup();
        break;
      }
      if (tok.t === 'cmd' && tok.v === '\\') {
        this.next();
        rows.push([]);
        continue;
      }
      if (tok.t === '&') {
        this.next();
        rows.at(-1).push(new MathRun(/matrix|array|cases/.test(name) ? ' ' : ''));
        continue;
      }
      rows.at(-1).push(...this.expression());
      if (this.peek()?.t === '}') this.next(); // stray brace
    }
    const lines = rows.filter((r) => r.length);
    const joined = lines.flatMap((r, i) => (i ? [new MathRun(/cases/.test(name) ? '; ' : ', '), ...r] : r));
    if (/^pmatrix/.test(name)) return new MathRoundBrackets({ children: joined });
    if (/^bmatrix/.test(name)) return new MathSquareBrackets({ children: joined });
    if (/^cases/.test(name)) return [new MathRun('{'), ...joined];
    this.rows = lines; // aligned/gathered: caller may split into separate lines
    return joined;
  }
  atom(out) {
    const tok = this.next();
    if (!tok) return null;
    if (tok.t === '{') {
      const inner = this.expression();
      if (this.peek()?.t === '}') this.next();
      return inner;
    }
    if (tok.t === '[' || tok.t === ']' || tok.t === 'char') {
      if (tok.v === "'") return new MathRun('′');
      // Keep runs of letters/digits together so "x" "2" don't become separate italic runs.
      let text = tok.v;
      while (this.peek()?.t === 'char' && /[0-9.]/.test(text.at(-1)) && /[0-9.]/.test(this.peek().v)) text += this.next().v;
      return new MathRun(text);
    }
    if (tok.t !== 'cmd') return new MathRun(tok.v);
    const name = tok.v;
    if (name === 'frac' || name === 'dfrac' || name === 'tfrac' || name === 'cfrac') {
      return new MathFraction({ numerator: this.group(), denominator: this.group() });
    }
    if (name === 'binom' || name === 'dbinom') {
      const n = this.group();
      const k = this.group();
      return new MathRoundBrackets({ children: [new MathFraction({ numerator: n, denominator: k })] });
    }
    if (name === 'sqrt') {
      let degree;
      if (this.peek()?.t === '[') {
        this.next();
        degree = this.expression(new Set([']']));
        if (this.peek()?.t === ']') this.next();
      }
      return new MathRadical({ children: this.group(), ...(degree ? { degree } : {}) });
    }
    if (NARY[name]) {
      if (this.peek()?.t === '^' || this.peek()?.t === '_') {
        const marker = new MathRun('');
        marker.__nary = name;
        return marker; // scripts() picks this up and builds the operator
      }
      return this.nary(name, null, null);
    }
    if (name === 'lim' || name === 'limsup' || name === 'liminf') {
      if (this.peek()?.t === '_') {
        this.next();
        return new MathSubScript({ children: [new MathRun(name)], subScript: this.group() });
      }
      return new MathRun(name);
    }
    if (FUNCTIONS.has(name)) return new MathRun(name);
    if (name === 'left') {
      const open = this.delimiter();
      const inner = this.expression();
      let close = '';
      if (this.peek()?.t === 'cmd' && this.peek().v === 'right') {
        this.next();
        close = this.delimiter();
      }
      return this.bracket(open, inner, close);
    }
    if (/^(big|Big|bigg|Bigg)[lr]?$/.test(name)) return new MathRun(this.delimiter());
    if (name === 'text' || name === 'textrm' || name === 'mathrm' || name === 'operatorname' || name === 'textbf' || name === 'mathbf' || name === 'textit' || name === 'mbox') {
      return new MathRun(this.textGroup());
    }
    if (name === 'mathbb') return new MathRun([...this.textGroup()].map((c) => BLACKBOARD[c] ?? c).join(''));
    if (name === 'mathcal' || name === 'mathit' || name === 'boldsymbol' || name === 'mathsf' || name === 'mathtt') return this.group();
    if (ACCENTS[name]) {
      const base = this.textGroup();
      return new MathRun(base + ACCENTS[name]);
    }
    if (name === 'begin') {
      const env = this.textGroup();
      if (/^array/.test(env) && this.peek()?.t === '{') this.textGroup(); // column spec
      return this.environment(env);
    }
    if (name === 'not') {
      const next = this.atom(out);
      return [].concat(next || []).length ? [new MathRun('¬'), ...[].concat(next)] : new MathRun('¬');
    }
    if (name in SYMBOLS) return SYMBOLS[name] ? new MathRun(SYMBOLS[name]) : null;
    return new MathRun(name); // unknown command: show its name rather than fail
  }
}

/**
 * Convert one LaTeX expression into Word equation components.
 * @param {string} tex
 * @returns {WordMath[]} one equation per line (multi-line aligned/gathered blocks become several)
 */
export function latexToWordMath(tex) {
  const src = tex.trim();
  // aligned/gathered blocks become one equation per line; matrices and cases keep their rows inside one
  // equation (the parser handles their \\ and &); otherwise a top-level \\ also starts a new line.
  const lines = /\\begin\{(aligned|align\*?|gathered|gather\*?|split|eqnarray\*?)\}/.test(src)
    ? src.replace(/\\(begin|end)\{[^}]*\}/g, '').split(/\\\\/)
    : /\\begin\{/.test(src)
      ? [src]
      : src.split(/\\\\/).filter((l) => l.trim());
  return (lines.length ? lines : [src]).map((line) => {
    // In aligned blocks "&" only marks the alignment point; matrices handle their own "&".
    const parser = new Parser(/\\begin\{/.test(line) ? line : line.replace(/&/g, ''));
    const children = [];
    while (parser.pos < parser.tokens.length) {
      children.push(...parser.expression());
      if (parser.pos < parser.tokens.length) parser.next(); // skip a stray } or \right
    }
    return new WordMath({ children: children.length ? children : [new MathRun(line)] });
  });
}
