import type { Math as WordMath } from 'docx';

/**
 * Convert one LaTeX expression (without delimiters such as `\(` `\)` or `$$`) into Word equations.
 * Returns one equation per line: multi-line `aligned` / `gathered` blocks and top-level `\\` give several.
 * Unknown commands degrade to plain text instead of throwing.
 */
export function latexToWordMath(tex: string): WordMath[];
