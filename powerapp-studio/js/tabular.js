// tabular.js — read CSV and Excel (.xlsx) files in the browser, guess SharePoint column types, and write rows back out as CSV.
// No dependencies beyond the vendored JSZip (passed in), so it also runs under Node for the tests.
import { newList, MAX_ROWS } from './core.js';
import { cleanName, makeColumn } from './edit.js';

// ------------------------------------------------------------------ CSV
/** Parse CSV text into rows of strings. Handles quotes, doubled quotes, commas/newlines inside quotes, CRLF, a BOM,
 *  and semicolon- or tab-separated files (detected from the first line). */
export function parseCSV(text) {
  let s = String(text ?? '');
  if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
  const first = s.slice(0, s.search(/\r?\n/) < 0 ? s.length : s.search(/\r?\n/)).replace(/"[^"]*"/g, '');
  const counts = [',', ';', '\t'].map((d) => [d, first.split(d).length - 1]).sort((a, b) => b[1] - a[1]);
  const D = counts[0][1] > 0 ? counts[0][0] : ',';
  const rows = [];
  let row = [], field = '', q = false, i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (q) {
      if (ch === '"') { if (s[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; }
      field += ch; i++; continue;
    }
    if (ch === '"' && field === '') { q = true; i++; continue; }
    if (ch === D) { row.push(field); field = ''; i++; continue; }
    if (ch === '\r' || ch === '\n') {
      row.push(field); field = ''; rows.push(row); row = [];
      if (ch === '\r' && s[i + 1] === '\n') i++;
      i++; continue;
    }
    field += ch; i++;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  // drop blank lines
  return rows.filter((r) => r.some((c) => String(c).trim() !== ''));
}

/** Write rows (arrays) as CSV, quoting only when needed. */
export function writeCSV(rows) {
  const cell = (v) => { const t = v == null ? '' : String(v); return /[",\r\n]/.test(t) || /^\s|\s$/.test(t) ? `"${t.replace(/"/g, '""')}"` : t; };
  return rows.map((r) => r.map(cell).join(',')).join('\r\n') + '\r\n';
}

// ------------------------------------------------------------------ XLSX
const ent = (t) => String(t).replace(/&(lt|gt|amp|quot|apos|#(\d+)|#x([0-9a-f]+));/gi, (m, n, d, h) => (d ? String.fromCodePoint(+d) : h ? String.fromCodePoint(parseInt(h, 16)) : { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[n.toLowerCase()]));
const attr = (tag, name) => { const m = tag.match(new RegExp(`\\s${name}="([^"]*)"`)); return m ? ent(m[1]) : null; };
const textOf = (xml) => [...xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => ent(m[1])).join('') || (/<t\b[^>]*\/>/.test(xml) ? '' : '');
const BUILTIN_DATE_FMTS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
const isDateFormat = (code) => { const c = String(code).replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); return /[dmy]/i.test(c) && !/^general$/i.test(c); };
/** Excel serial (1900 date system) to YYYY-MM-DD. */
export const excelDate = (n) => new Date(Math.round((Number(n) - 25569) * 864e5)).toISOString().slice(0, 10);
const colIndex = (ref) => { const L = String(ref).match(/^[A-Z]+/i)?.[0].toUpperCase() || 'A'; let n = 0; for (const ch of L) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

/** Read the first worksheet of an .xlsx file into rows of strings (dates as YYYY-MM-DD). */
export async function readXlsx(data, JSZip) {
  const zip = await JSZip.loadAsync(data);
  const read = async (p) => (zip.file(p) ? zip.file(p).async('string') : '');
  // first sheet in the workbook's order
  let sheetPath = 'xl/worksheets/sheet1.xml';
  const wb = await read('xl/workbook.xml'), rels = await read('xl/_rels/workbook.xml.rels');
  const firstSheet = wb.match(/<sheet\b[^>]*>/);
  if (firstSheet && rels) {
    const rid = attr(firstSheet[0], 'r:id');
    const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).find((t) => attr(t, 'Id') === rid);
    const target = rel && attr(rel, 'Target');
    if (target) sheetPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
  }
  const sheet = await read(sheetPath);
  if (!sheet) throw new Error('No worksheet found in this file.');
  const shared = [...(await read('xl/sharedStrings.xml')).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => textOf(m[1]));
  // which cell styles are dates
  const styles = await read('xl/styles.xml');
  const custom = new Map([...styles.matchAll(/<numFmt\b[^>]*>/g)].map((m) => [Number(attr(m[0], 'numFmtId')), attr(m[0], 'formatCode') || '']));
  const xfs = (styles.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/)?.[0] || '').match(/<xf\b[^>]*>/g) || [];
  const dateStyle = xfs.map((t) => { const id = Number(attr(t, 'numFmtId') || 0); return BUILTIN_DATE_FMTS.has(id) || (custom.has(id) && isDateFormat(custom.get(id))); });
  const rows = [];
  for (const rm of sheet.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g)) {
    const r = [];
    for (const cm of (rm[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const head = '<c' + cm[1] + '>', inner = cm[2] || '';
      const t = attr(head, 't'), s = Number(attr(head, 's') || 0), ref = attr(head, 'r');
      const v = inner.match(/<v>([\s\S]*?)<\/v>/)?.[1];
      let val = '';
      if (t === 's') val = shared[Number(v)] ?? '';
      else if (t === 'inlineStr') val = textOf(inner);
      else if (t === 'b') val = v === '1' ? 'TRUE' : 'FALSE';
      else if (t === 'str' || t === 'e') val = v != null ? ent(v) : '';
      else if (v != null) val = dateStyle[s] && Number.isFinite(Number(v)) ? excelDate(v) : ent(v);
      const at = ref ? colIndex(ref) : r.length;
      while (r.length < at) r.push('');
      r[at] = val;
    }
    rows.push(r);
  }
  return rows.filter((r) => r.some((c) => String(c).trim() !== ''));
}

// ------------------------------------------------------------------ type inference
const EMAIL = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;
const YES = /^(yes|no|true|false|y|n)$/i;
const NUM = /^\(?-?\$?\s?-?\d{1,3}(,\d{3})*(\.\d+)?\)?$|^\(?-?\$?\s?-?\d+(\.\d+)?\)?$/;
const MONEY_HEAD = /cost|price|amount|budget|total|fee|salary|spend|\$|usd|revenue/i;
const ISO = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;
const US = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?\s*([AP]M)?)?$/i;
/** A date value as YYYY-MM-DD, or '' when it is not a date. */
export function toIsoDate(v) {
  const t = String(v ?? '').trim();
  let m = t.match(ISO);
  if (m) { const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); return d.getUTCMonth() === +m[2] - 1 ? d.toISOString().slice(0, 10) : ''; }
  m = t.match(US);
  if (m) { const y = m[3].length === 2 ? 2000 + +m[3] : +m[3]; const d = new Date(Date.UTC(y, +m[1] - 1, +m[2])); return d.getUTCMonth() === +m[1] - 1 ? d.toISOString().slice(0, 10) : ''; }
  return '';
}
export const toNumber = (v) => { const t = String(v ?? '').trim(); if (!NUM.test(t)) return null; const neg = /^\(.*\)$/.test(t) || /-/.test(t); const n = Number(t.replace(/[()$,\s-]/g, '')); return Number.isFinite(n) ? (neg ? -n : n) : null; };
export const toBool = (v) => /^(yes|true|y|1)$/i.test(String(v ?? '').trim());

/** Guess a SharePoint type for each column. Returns [{header, type, choices, sample:[...]}]. */
export function inferColumns(header, rows) {
  return header.map((h, i) => {
    const vals = rows.map((r) => String(r[i] ?? '').trim()).filter((v) => v !== '');
    const sample = vals.slice(0, 3);
    const info = { header: String(h || '').trim() || `Column ${i + 1}`, type: 'Text', choices: [], sample, filled: vals.length };
    if (!vals.length) return info;
    const distinct = [...new Set(vals)];
    if (vals.every((v) => YES.test(v))) info.type = 'YesNo';
    else if (vals.every((v) => EMAIL.test(v))) info.type = 'Person';
    else if (vals.every((v) => toNumber(v) !== null) && !vals.some((v) => /^0\d/.test(v) || v.replace(/\D/g, '').length > 15)) {
      const money = vals.some((v) => /\$/.test(v)) || (MONEY_HEAD.test(info.header) && vals.every((v) => /^\D*-?[\d,]+(\.\d{1,2})?\D*$/.test(v)));
      info.type = money ? 'Currency' : 'Number';
    } else if (vals.every((v) => toIsoDate(v))) info.type = 'DateTime';
    else if (vals.some((v) => v.length > 255 || /\n/.test(v))) info.type = 'Note';
    else if (distinct.length <= 12 && distinct.length < vals.length && vals.length >= 4 && distinct.every((v) => v.length <= 60)) { info.type = 'Choice'; info.choices = distinct; }
    return info;
  });
}
/** Convert one cell to the value stored for its column type ('' when blank or unreadable). */
export function convertCell(type, v) {
  const t = String(v ?? '').trim();
  if (t === '') return type === 'YesNo' ? false : '';
  switch (type) {
    case 'Number': case 'Currency': { const n = toNumber(t); return n === null ? '' : n; }
    case 'DateTime': return toIsoDate(t);
    case 'YesNo': return toBool(t);
    case 'Person': return EMAIL.test(t) ? t.toLowerCase() : '';
    case 'Text': case 'Choice': case 'Lookup': return t.slice(0, 255);
    default: return t.slice(0, 4000);
  }
}
/** Rows (arrays) → records keyed by column name. map: [{index, name, type}] */
export function toRecords(rows, map) {
  return rows.map((r) => Object.fromEntries(map.map((m) => [m.name, convertCell(m.type, r[m.index])]).filter(([, v]) => v !== '')));
}
/** A list's rows as CSV, with internal column names as headers (SharePoint "Import from CSV" / Quick Edit paste). */
export function listToCsv(list) {
  const cols = list.columns;
  const out = (c, v) => (v == null || v === '' ? '' : c.type === 'YesNo' ? (v ? 'Yes' : 'No') : String(v));
  return writeCSV([cols.map((c) => c.name), ...(list.rows || []).map((r) => cols.map((c) => out(c, r[c.name])))]);
}

/**
 * Build a SharePoint list (with its rows) from a parsed table.
 * spec: {name, rows (data rows, arrays), cols: [{index, include, name, type, choices}], titleIndex}
 * The Title column takes the values of titleIndex; every other included column becomes a new column.
 * Returns {list, notes, truncated}.
 */
export function listFromTable(project, spec) {
  const notes = [];
  let name = cleanName(spec.name, 'ImportedList'), k = 2;
  name = name[0].toUpperCase() + name.slice(1);
  const base = name; while (project.lists.some((l) => l.name === name)) name = base + k++;
  if (name !== base) notes.push(`A list called ${base} already exists, so this one is ${name}.`);
  const list = newList(name);
  const map = [];
  if (spec.titleIndex != null && spec.titleIndex >= 0) map.push({ index: spec.titleIndex, name: 'Title', type: 'Text' });
  for (const c of spec.cols) {
    if (!c.include || c.index === spec.titleIndex) continue;
    let r = makeColumn(list, { name: c.name, type: c.type, choices: c.choices });
    if (!r.col && /built-in/.test(r.error || '')) { // Created, Modified, ID… are SharePoint's own: keep the data under a prefixed name
      const single = list.name.replace(/ies$/, 'y').replace(/s$/, ''), nm = cleanName(c.name, 'Field');
      r = makeColumn(list, { name: single + nm[0].toUpperCase() + nm.slice(1), type: c.type, choices: c.choices });
      if (r.col) r.note = `${nm} renamed to ${r.col.name} (SharePoint already uses ${nm}).`;
    }
    if (!r.col) { notes.push(r.error); continue; }
    if (r.note) notes.push(r.note);
    list.columns.push(r.col);
    map.push({ index: c.index, name: r.col.name, type: r.col.type });
  }
  const all = spec.rows || [];
  const truncated = all.length > MAX_ROWS;
  if (truncated) notes.push(`The file has ${all.length} rows; the first ${MAX_ROWS} are kept as sample data. Import the full file straight into SharePoint.`);
  list.rows = toRecords(all.slice(0, MAX_ROWS), map);
  // a Choice column gets every value that appears
  for (const c of list.columns) if (c.type === 'Choice') { const seen = new Set(c.choices); for (const r of list.rows) if (r[c.name] && !seen.has(r[c.name])) { seen.add(r[c.name]); c.choices.push(r[c.name]); } }
  const blankTitle = list.rows.filter((r) => !r.Title).length;
  if (blankTitle) notes.push(`${blankTitle} row${blankTitle === 1 ? ' has' : 's have'} no Title; SharePoint requires one.`);
  return { list, notes, truncated };
}
