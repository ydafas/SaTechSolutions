// delegation.js — large SharePoint lists. A list marked "large" (more rows than the data row limit)
// gets delegation-safe output and lint warnings for formulas SharePoint cannot run server-side.
// Lists without a size keep the exact output they always had.
import { q } from './core.js';

export const ROW_LIMIT = 2000;
export const SIZES = {
  small: { label: 'Small (under 500 items)', note: 'Default data row limit (500) is enough.' },
  medium: { label: 'Medium (500 to 2,000 items)', note: 'Raise the data row limit to 2000 in Studio settings.' },
  large: { label: 'Large (more than 2,000 items)', note: 'Counts show "2000+", local-only filters run on the newest 2,000 items, and lint flags formulas SharePoint cannot delegate.' },
};
export const largeLists = (project) => new Set((project.lists || []).filter((l) => l.size === 'large').map((l) => l.name));
export const needsRowLimit = (project) => (project.lists || []).some((l) => l.size === 'medium' || l.size === 'large');

const esc = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Regex source matching a reference to list L: Name, 'Name', [@Name], [@'Name']. */
const refSrc = (L) => { const qn = q(L); return `(?:\\[@${esc(qn)}\\]|${esc(qn)}${qn === L ? '\\b' : ''})`; };

/** Index of the quote that closes the string opening at i ("" is an escaped quote). */
function endString(t, i) {
  let j = i + 1;
  while (j < t.length) { if (t[j] === '"') { if (t[j + 1] === '"') { j += 2; continue; } return j; } j++; }
  return t.length;
}
/** Index just past the ")" matching the "(" at i, skipping "strings". */
export function closeParen(fx, i) {
  let depth = 0;
  for (let j = i; j < fx.length; j++) {
    const c = fx[j];
    if (c === '"') { j = endString(fx, j); continue; }
    if (c === '(') depth++;
    else if (c === ')') { depth--; if (depth === 0) return j + 1; }
  }
  return -1;
}
/** Every call fn(...) in fx: [{start, end, args}] where args is the text between the parens. */
export function calls(fx, fn) {
  const out = [], re = new RegExp(`(?<![A-Za-z0-9_.])${fn}\\(`, 'g');
  let m;
  while ((m = re.exec(fx))) {
    const open = m.index + fn.length, end = closeParen(fx, open);
    if (end < 0) break;
    out.push({ start: m.index, end, args: fx.slice(open + 1, end - 1) });
  }
  return out;
}
/** Split top-level arguments of a call's argument text. */
export function splitArgs(args) {
  const out = []; let depth = 0, cur = '';
  for (let i = 0; i < args.length; i++) {
    const c = args[i];
    if (c === '"') { const j = endString(args, i); cur += args.slice(i, j + 1); i = j; continue; }
    if ('({['.includes(c)) depth++; else if (')}]'.includes(c)) depth--;
    if (c === ',' && depth === 0) { out.push(cur.trim()); cur = ''; } else cur += c;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}
const startsWithList = (arg, L) => new RegExp(`^(?:${refSrc(L)}|Filter\\(\\s*${refSrc(L)}\\s*,)`).test(arg.trim()) || new RegExp(`^${refSrc(L)}$`).test(arg.trim());

/** Large lists: Text(CountRows(<list or Filter(list, …)>)) → "2000+" once the row limit is reached. */
export function capCounts(fx, lists) {
  if (!lists.size || !/CountRows\(/.test(fx)) return fx;
  let out = '', last = 0;
  for (const c of calls(fx, 'Text')) {
    if (c.start < last) continue;
    const inner = c.args.trim();
    if (!/^CountRows\(/.test(inner) || closeParen(inner, 'CountRows'.length) !== inner.length) continue;
    const arg = inner.slice('CountRows('.length, -1);
    if (![...lists].some((L) => startsWithList(arg, L))) continue;
    out += fx.slice(last, c.start) + `With({n: CountRows(${arg})}, If(n >= ${ROW_LIMIT}, "${ROW_LIMIT}+", Text(n)))`;
    last = c.end;
  }
  return out + fx.slice(last);
}

const NON_DELEGABLE = [
  [/<>/, 'a "<>" comparison'],
  [/(?<![A-Za-z0-9_])in(?![A-Za-z0-9_])/, 'an "in" test'],
  [/=\s*Blank\(\)/, 'a "= Blank()" test'],
  [/IsBlank\(/, 'IsBlank()'],
  [/\b(?:Len|Lower|Upper|Trim|Left|Right|Mid|Text|Year|Month|Day)\(/, 'a function on a column'],
];
/** Delegation problems with large lists in one formula: [{list, msg}]. */
export function delegationIssues(fx, lists) {
  const out = [];
  if (!lists.size || !fx) return out;
  const capped = (start) => fx.slice(Math.max(0, start - 20), start).includes('With({n: ');
  for (const L of lists) {
    const ref = new RegExp(refSrc(L));
    if (!ref.test(fx)) continue;
    for (const c of calls(fx, 'CountRows')) if (startsWithList(c.args, L) && !capped(c.start)) out.push({ list: L, msg: `CountRows on ${L} stops at the data row limit (${ROW_LIMIT}), so the number is wrong once the list is bigger. Wrap it as Text(CountRows(…)) so the studio shows "${ROW_LIMIT}+", or count with a flow.` });
    for (const c of calls(fx, 'Search')) if (startsWithList(c.args, L)) out.push({ list: L, msg: `Search() on ${L} is not delegable to SharePoint; only the first ${ROW_LIMIT} items are searched. Use StartsWith in a Filter instead.` });
    for (const c of calls(fx, 'Filter')) {
      const [src, ...preds] = splitArgs(c.args);
      if (!new RegExp(`^${refSrc(L)}$`).test(src || '')) continue;
      for (const p of preds) for (const [re, what] of NON_DELEGABLE) if (re.test(p.replace(/"(?:[^"]|"")*"/g, '""'))) {
        out.push({ list: L, msg: `Filter on ${L} uses ${what} (${p.length > 60 ? p.slice(0, 57) + '…' : p}), which SharePoint cannot delegate; only the first ${ROW_LIMIT} items are checked. Use Filter(Filter(${q(L)}, <delegable part>), <the rest>) so the server narrows the list first.` });
        break;
      }
    }
  }
  return out;
}
