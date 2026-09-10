// CSP-safe stub for Ajv internal codegen helpers.
//
// Some packages import `ajv/dist/compile/codegen` and expect helpers such as
// `operators`, `str`, `_`, `or`, and `KeywordCxt` to exist. In Chrome/Edge MV3
// service workers we must avoid Ajv's real codegen because it can use
// `new Function()`.

export const operators = {
  LT: '<',
  LTE: '<=',
  GT: '>',
  GTE: '>=',
  EQ: '===',
  NEQ: '!==',
} as const;

function interpolate(strings: TemplateStringsArray, exprs: unknown[]): string {
  let out = '';
  for (let i = 0; i < strings.length; i += 1) {
    out += strings[i] ?? '';
    if (i < exprs.length) out += String(exprs[i]);
  }
  return out;
}

export function _(strings: TemplateStringsArray, ...exprs: unknown[]): string {
  return interpolate(strings, exprs);
}

export function str(strings: TemplateStringsArray, ...exprs: unknown[]): string {
  return interpolate(strings, exprs);
}

export function or(...parts: unknown[]): string {
  return parts.map(part => String(part)).join(' || ');
}

export function and(...parts: unknown[]): string {
  return parts.map(part => String(part)).join(' && ');
}

export function getProperty(prop: unknown): string {
  if (typeof prop === 'string' && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(prop)) return `.${prop}`;

  const safe = String(prop).replace(/[\\'\n\r\u2028\u2029]/g, ch => {
    switch (ch) {
      case '\\':
        return '\\\\';
      case "'":
        return "\\'";
      case '\n':
        return '\\n';
      case '\r':
        return '\\r';
      case '\u2028':
        return '\\u2028';
      case '\u2029':
        return '\\u2029';
      default:
        return ch;
    }
  });

  return `['${safe}']`;
}

export class KeywordCxt {
  public $data = false;
  public schemaCode: string = '""';
  public schema: unknown = undefined;

  constructor(_it: unknown, _def: unknown, _keyword: string) {}
}
