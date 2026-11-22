// render.js — draws controls in the designer and runs them in Play mode.
// Play mode simulates the generated app against sample data: navigation, galleries,
// search, forms that save, approve/reject, variables, Visible conditions, notifications.
import { cssColor, displayValue, colOf, varSel, walk, sampleRows, PEOPLE, FORM_FACTORS, hasActions } from './core.js';
import { ICONS } from './icons.js';

const PT = 96 / 72; // Power Apps font Size is in points
export const FONT_CSS = {
  'Segoe UI': "'Segoe UI','Segoe UI Web (West European)',system-ui,sans-serif",
  'Open Sans': "'Open Sans',system-ui,sans-serif", Lato: "'Lato',system-ui,sans-serif",
  Arial: 'Arial,Helvetica,sans-serif', Verdana: 'Verdana,sans-serif', Georgia: 'Georgia,serif', 'Courier New': "'Courier New',monospace",
};
const esc = (t) => String(t ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
export const svgIcon = (name, color, stroke = 2) =>
  `<svg viewBox="0 0 24 24" width="100%" height="100%" aria-hidden="true"><path d="${ICONS[name] || ICONS.info}" fill="none" stroke="${color}" stroke-width="${stroke}" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

export function sanitizeHtml(html) {
  const doc = new DOMParser().parseFromString(`<div>${html}</div>`, 'text/html');
  doc.querySelectorAll('script,iframe,object,embed,link,meta,style').forEach((n) => n.remove());
  doc.querySelectorAll('*').forEach((n) => [...n.attributes].forEach((a) => { if (/^on/i.test(a.name) || /javascript:/i.test(a.value)) n.removeAttribute(a.name); }));
  return doc.body.firstChild.innerHTML;
}

// ============================================================ mini Power Fx evaluator
// Enough of the language to preview titles, Visible rules and simple behaviour formulas.
class FxError extends Error {}
function tokenize(src) {
  const t = [], re = /\s+|\/\/[^\n]*|"(?:[^"]|"")*"|'(?:[^']|'')*'|\d+(?:\.\d+)?|[A-Za-z_@][A-Za-z0-9_]*|<>|<=|>=|&&|\|\||[-+*/&=<>!(),.;{}:\[\]]/gy;
  let m;
  while (re.lastIndex < src.length) {
    const i = re.lastIndex; m = re.exec(src);
    if (!m) throw new FxError('token at ' + i);
    const v = m[0];
    if (/^\s|^\/\//.test(v)) continue;
    if (v[0] === '"') t.push({ k: 'str', v: v.slice(1, -1).replace(/""/g, '"') });
    else if (v[0] === "'") t.push({ k: 'id', v: v.slice(1, -1).replace(/''/g, "'") });
    else if (/^\d/.test(v)) t.push({ k: 'num', v: Number(v) });
    else if (/^[A-Za-z_@]/.test(v)) t.push({ k: 'id', v });
    else t.push({ k: 'op', v });
  }
  return t;
}
function parse(src) {
  const T = tokenize(src); let i = 0;
  const peek = (v) => T[i] && T[i].k === 'op' && T[i].v === v;
  const eat = (v) => { if (!peek(v)) throw new FxError('expected ' + v); i++; };
  const BIN = [['||', 'Or'], ['&&', 'And'], ['=', '<>', '<', '>', '<=', '>='], ['&'], ['+', '-'], ['*', '/']];
  const chain = () => { const parts = [expr(0)]; while (peek(';')) { i++; if (i < T.length) parts.push(expr(0)); } return parts.length === 1 ? parts[0] : { k: 'seq', parts }; };
  function expr(lvl) {
    if (lvl >= BIN.length) return unary();
    let l = expr(lvl + 1);
    for (;;) {
      const tk = T[i];
      const isOp = tk && ((tk.k === 'op' && BIN[lvl].includes(tk.v)) || (tk.k === 'id' && BIN[lvl].includes(tk.v)));
      if (!isOp) return l;
      i++; l = { k: 'bin', op: tk.v === 'Or' ? '||' : tk.v === 'And' ? '&&' : tk.v, l, r: expr(lvl + 1) };
    }
  }
  function unary() {
    if (peek('!') || (T[i]?.k === 'id' && T[i].v === 'Not' && !(T[i + 1]?.v === '('))) { i++; return { k: 'not', e: unary() }; }
    if (peek('-')) { i++; return { k: 'neg', e: unary() }; }
    return postfix(primary());
  }
  function primary() {
    const tk = T[i++];
    if (!tk) throw new FxError('end');
    if (tk.k === 'str') return { k: 'lit', v: tk.v };
    if (tk.k === 'num') return { k: 'lit', v: tk.v };
    if (tk.k === 'op' && tk.v === '(') { const e = chain(); eat(')'); return e; }
    if (tk.k === 'op' && tk.v === '{') {
      const fields = {};
      while (!peek('}')) { const key = T[i++].v; eat(':'); fields[key] = expr(0); if (peek(',')) i++; }
      eat('}'); return { k: 'rec', fields };
    }
    if (tk.k === 'op' && tk.v === '[') { const items = []; while (!peek(']')) { items.push(expr(0)); if (peek(',')) i++; } eat(']'); return { k: 'arr', items }; }
    if (tk.k === 'id') {
      if (tk.v === 'true' || tk.v === 'false') return { k: 'lit', v: tk.v === 'true' };
      if (peek('(')) { i++; const args = []; while (!peek(')')) { args.push(chain()); if (peek(',')) i++; } eat(')'); return { k: 'call', f: tk.v, args }; }
      return { k: 'id', v: tk.v };
    }
    throw new FxError('unexpected ' + tk.v);
  }
  function postfix(e) { while (peek('.')) { i++; e = { k: 'dot', e, m: T[i++].v }; } return e; }
  const ast = chain();
  if (i < T.length) throw new FxError('trailing');
  return ast;
}
const blank = (v) => v == null || v === '' || (typeof v === 'object' && !Array.isArray(v) && Object.keys(v).length === 0);
const toText = (v) => (v == null ? '' : typeof v === 'object' ? (v.DisplayName || v.Value || '') : typeof v === 'boolean' ? (v ? 'true' : 'false') : String(v));
const cache = new Map();
export function evalFx(src, env) {
  if (src == null) return undefined;
  let ast = cache.get(src);
  if (!ast) { ast = parse(String(src).replace(/^=/, '')); cache.set(src, ast); }
  return ev(ast, env);
}
export function tryFx(src, env, fallback) { try { return evalFx(src, env); } catch { return fallback; } }
const themeCache = new WeakMap();
function themeRecord(th) {
  if (themeCache.has(th)) return themeCache.get(th);
  const r = { Radius: th.radius, Font: th.font };
  for (const [k, v] of Object.entries(th.colors)) { r[k] = v; r[k + 'Hex'] = v; }
  themeCache.set(th, r); return r;
}
function member(o, m) {
  if (o == null) return undefined;
  if (typeof o === 'string' && (m === 'Value' || m === 'Text')) return o;
  if (typeof o === 'object') return o[m];
  return undefined;
}
function ev(n, env) {
  const rt = env.rt;
  switch (n.k) {
    case 'lit': return n.v;
    case 'seq': { let r; for (const p of n.parts) r = ev(p, env); return r; }
    case 'not': return !ev(n.e, env);
    case 'neg': return -ev(n.e, env);
    case 'rec': { const o = {}; for (const [k, v] of Object.entries(n.fields)) o[k] = ev(v, env); return o; }
    case 'arr': return n.items.map((x) => ({ Value: ev(x, env) }));
    case 'id': {
      const v = n.v;
      if (v === 'ThisItem') return env.row || {};
      if (v === 'appTheme' && env.project) return themeRecord(env.project.theme);
      if (v === 'Self') return {};
      if (rt?.vars && v in rt.vars) return rt.vars[v];
      if (rt?.ctlValues && v in rt.ctlValues) return rt.ctlValues[v];
      if (env.project?.lists.some((l) => l.name === v)) return { __list: v };
      if (env.project?.screens.some((s) => s.name === v)) return { __screen: v };
      if (/^(var|loc|gbl|_)/.test(v)) return undefined;
      throw new FxError('unknown ' + v);
    }
    case 'dot': {
      const o = ev(n.e, env);
      if (o?.__ctl) return o.get(n.m);
      return member(o, n.m);
    }
    case 'bin': {
      if (n.op === '&&') return !!ev(n.l, env) && !!ev(n.r, env);
      if (n.op === '||') return !!ev(n.l, env) || !!ev(n.r, env);
      const a = ev(n.l, env), b = ev(n.r, env);
      switch (n.op) {
        case '&': return toText(a) + toText(b);
        case '=': return toText(a) === toText(b) || (blank(a) && blank(b));
        case '<>': return !(toText(a) === toText(b) || (blank(a) && blank(b)));
        case '<': return a < b; case '>': return a > b; case '<=': return a <= b; case '>=': return a >= b;
        case '+': return Number(a || 0) + Number(b || 0); case '-': return Number(a || 0) - Number(b || 0);
        case '*': return Number(a || 0) * Number(b || 0); case '/': return Number(a || 0) / Number(b || 1);
      }
      return undefined;
    }
    case 'call': return call(n, env);
  }
  throw new FxError('node');
}
function call(n, env) {
  const A = (i) => ev(n.args[i], env), rt = env.rt;
  switch (n.f) {
    case 'If': {
      for (let i = 0; i + 1 < n.args.length; i += 2) if (A(i)) return A(i + 1);
      return n.args.length % 2 ? A(n.args.length - 1) : undefined;
    }
    case 'Switch': {
      const v = toText(A(0));
      for (let i = 1; i + 1 < n.args.length; i += 2) if (toText(A(i)) === v) return A(i + 1);
      return (n.args.length - 1) % 2 ? A(n.args.length - 1) : undefined;
    }
    case 'IsBlank': case 'IsEmpty': return blank(A(0));
    case 'Not': return !A(0);
    case 'And': return n.args.every((_, i) => A(i));
    case 'Or': return n.args.some((_, i) => A(i));
    case 'Coalesce': for (let i = 0; i < n.args.length; i++) { const v = A(i); if (!blank(v)) return v; } return undefined;
    case 'Text': { const v = A(0); const f = n.args[1] ? A(1) : ''; if (/mm\/dd/.test(f)) return displayValue({ type: 'DateTime' }, v); if (/\$/.test(f)) return displayValue({ type: 'Currency' }, v); return toText(v); }
    case 'Value': return Number(A(0)) || 0;
    case 'Len': return toText(A(0)).length;
    case 'Upper': return toText(A(0)).toUpperCase();
    case 'Lower': return toText(A(0)).toLowerCase();
    case 'Proper': return toText(A(0)).replace(/\b\w/g, (c) => c.toUpperCase());
    case 'Trim': return toText(A(0)).trim();
    case 'Concatenate': return n.args.map((_, i) => toText(A(i))).join('');
    case 'Char': return String.fromCharCode(A(0));
    case 'EncodeUrl': return encodeURIComponent(toText(A(0)));
    case 'Sequence': return Array.from({ length: Number(A(0)) || 0 }, (_, i) => ({ Value: i + 1 }));
    case 'FirstN': { const t = A(0); return Array.isArray(t) ? t.slice(0, Number(A(1)) || 0) : t; }
    case 'Today': return new Date().toISOString().slice(0, 10);
    case 'Now': return new Date().toLocaleString();
    case 'User': return { FullName: 'Alex Rivera', Email: 'alex.rivera@contoso.com' };
    case 'Defaults': return {};
    case 'Blank': return undefined;
    case 'CountRows': { const t = A(0); return Array.isArray(t) ? t.length : t?.__list ? (rt?.db[t.__list] || []).length : 0; }
    case 'Filter': {
      const t = A(0); const rows = Array.isArray(t) ? t : rt?.db[t?.__list] || [];
      return rows.filter((row) => n.args.slice(1).every((a) => tryFx2(a, { ...env, row, scope: row })));
    }
    case 'LookUp': { const t = A(0); const rows = Array.isArray(t) ? t : rt?.db[t?.__list] || []; return rows.find((row) => tryFx2(n.args[1], { ...env, row })); }
    case 'Set': { const name = n.args[0].v; const v = A(1); if (rt) rt.vars[name] = v; return true; }
    case 'UpdateContext': { const r = A(0); if (rt) Object.assign(rt.vars, r); return true; }
    case 'Navigate': { const s = A(0); if (rt && s?.__screen) rt.api.navigate(s.__screen); return true; }
    case 'Back': rt?.api.back(); return true;
    case 'Notify': rt?.api.toast(toText(A(0))); return true;
    case 'Launch': rt?.api.launch(toText(A(0))); return true;
    case 'Reset': case 'Refresh': case 'Select': case 'ClearCollect': case 'Collect': case 'Clear': return true;
    case 'ColorValue': case 'RGBA': case 'ColorFade': return undefined;
  }
  throw new FxError('fn ' + n.f);
}
function tryFx2(ast, env) {
  // evaluate a filter predicate with the row's columns in scope as bare names
  const scoped = { ...env, rt: { ...env.rt, vars: { ...(env.rt?.vars || {}), ...(env.row || {}) } } };
  try { return ev(ast, scoped); } catch { return true; }
}

// ============================================================ runtime
export function createRuntime(project, hooks) {
  const db = {};
  for (const l of project.lists) db[l.name] = sampleRows(l, project.lists, 8);
  const rt = { db, vars: {}, inputs: {}, galSel: {}, history: [], screen: null, ctlValues: {}, api: null };
  const start = project.screens.find((s) => s.id === project.app.startScreen) || project.screens[0];
  rt.screen = start?.id;
  for (const l of project.lists) rt.vars[varSel(l.name)] = {};
  const scr = () => project.screens.find((s) => s.id === rt.screen);
  rt.api = {
    navigate(idOrName) {
      const s = project.screens.find((x) => x.id === idOrName || x.name === idOrName);
      if (!s) return;
      rt.history.push(rt.screen); rt.screen = s.id; onVisible(s);
    },
    back() { if (rt.history.length) { rt.screen = rt.history.pop(); onVisible(scr(), true); } },
    toast: (t, kind) => hooks.toast(t, kind),
    launch: (u) => { try { window.open(u, '_blank', 'noopener'); } catch { /* sandboxed */ } hooks.toast('Would open ' + u); },
  };
  function onVisible(s) {
    for (const [c] of walk(s.controls)) delete rt.inputs[c.name];
    if (s.onVisible) tryFx(s.onVisible, { project, rt });
  }
  rt.run = (screen, actions, row) => {
    for (const a of actions || []) if (runAction(project, rt, screen, a, row) === false) break;
    hooks.rerender();
  };
  rt.current = scr;
  return rt;
}

function inputValue(project, rt, ctl) {
  if (ctl.name in rt.inputs) return rt.inputs[ctl.name];
  const f = ctl.field;
  if (f?.list) {
    const rec = rt.vars[varSel(f.list)] || {};
    const v = rec[f.column];
    if (ctl.type === 'person') return v || null;
    if (ctl.type === 'checkbox' || ctl.type === 'toggle') return !!v;
    if (ctl.type === 'datepicker') return v || (colOf(project, f.list, f.column)?.required ? new Date().toISOString().slice(0, 10) : '');
    if (ctl.type === 'dropdown') { const col = colOf(project, f.list, f.column); return v ?? (col?.type === 'Choice' ? col.choices[0] : ''); }
    return v == null ? '' : String(v);
  }
  if (ctl.type === 'checkbox' || ctl.type === 'toggle') return !!ctl.props.defaultOn;
  if (ctl.type === 'dropdown') return String(ctl.props.items || '').split('\n')[0] || '';
  if (ctl.type === 'datepicker') return new Date().toISOString().slice(0, 10);
  if (ctl.type === 'person') return null;
  return ctl.props.defaultText || '';
}

function runAction(project, rt, screen, a, row) {
  const V = a.list ? varSel(a.list) : '';
  const toScreen = (id) => id && rt.api.navigate(id);
  const after = () => { if (a.then === 'back') rt.api.back(); else if (a.then === 'navigate') toScreen(a.screen); };
  switch (a.type) {
    case 'navigate': toScreen(a.screen); break;
    case 'back': rt.api.back(); break;
    case 'newItem': rt.vars[V] = {}; toScreen(a.screen); break;
    case 'openItem': rt.vars[V] = row || {}; if (a.screen) toScreen(a.screen); break;
    case 'save': {
      const rows = rt.db[a.list] || (rt.db[a.list] = []);
      const rec = {};
      const missing = [];
      for (const [c, parent] of walk(screen.controls)) {
        if (parent || c.field?.list !== a.list) continue;
        const col = colOf(project, a.list, c.field.column); if (!col) continue;
        let v = inputValue(project, rt, c);
        if (col.type === 'Number' || col.type === 'Currency') v = v === '' ? null : Number(v);
        if (col.required && (v == null || v === '')) missing.push(col.name);
        rec[col.name] = v;
      }
      if (missing.length) { rt.api.toast('Save failed. Required: ' + missing.join(', '), 'error'); return false; }
      const cur = rt.vars[V];
      if (cur?.ID) { const r = rows.find((x) => x.ID === cur.ID); Object.assign(r, rec, { Modified: new Date().toISOString().slice(0, 10) }); rt.vars[V] = r; }
      else {
        const r = { ID: Math.max(0, ...rows.map((x) => x.ID)) + 1, Created: new Date().toISOString().slice(0, 10), Modified: new Date().toISOString().slice(0, 10), ...rec };
        const st = project.lists.find((l) => l.name === a.list)?.columns.find((c) => c.type === 'Choice' && !(c.name in rec));
        if (st) r[st.name] = st.choices[0];
        rows.unshift(r); rt.vars[V] = a.then === 'stay' ? {} : r;
      }
      rt.api.toast(a.message || 'Saved', 'success');
      if (a.then === 'stay') for (const [c] of walk(screen.controls)) delete rt.inputs[c.name];
      after(); break;
    }
    case 'delete': {
      const cur = rt.vars[V];
      if (!cur?.ID) { rt.api.toast('Delete failed. Nothing selected.', 'error'); return false; }
      rt.db[a.list] = (rt.db[a.list] || []).filter((x) => x.ID !== cur.ID);
      rt.vars[V] = {}; rt.api.toast(a.message || 'Deleted', 'success'); after(); break;
    }
    case 'setField': {
      const cur = rt.vars[V];
      if (!cur?.ID) { rt.api.toast('Update failed. Nothing selected.', 'error'); return false; }
      const r = (rt.db[a.list] || []).find((x) => x.ID === cur.ID);
      const col = colOf(project, a.list, a.field);
      if (r && col) r[a.field] = col.type === 'Person' ? { DisplayName: 'Alex Rivera', Mail: 'alex.rivera@contoso.com' } : col.type === 'YesNo' ? /^(true|yes|1)$/i.test(a.value) : col.type === 'DateTime' ? new Date().toISOString().slice(0, 10) : a.value;
      rt.vars[V] = r || {}; rt.api.toast(a.message || 'Updated', 'success'); after(); break;
    }
    case 'setVar': rt.vars[a.name] = tryFx(a.value || 'true', { project, rt, row }, a.value); break;
    case 'toggleVar': rt.vars[a.name] = !rt.vars[a.name]; break;
    case 'notify': rt.api.toast(a.text || 'Done', (a.kind || '').toLowerCase()); break;
    case 'launch': rt.api.launch(a.url); break;
    case 'refresh': break;
    case 'reset': for (const [c] of walk(screen.controls)) delete rt.inputs[c.name]; break;
    case 'custom': try { evalFx(a.fx || '', { project, rt, row }); } catch { rt.api.toast('Custom formula runs in the real app only'); } break;
  }
  return true;
}

function galleryRowsFor(project, rt, ctl) {
  const d = ctl.data || {};
  if (!d.list && ctl.fx?.Items) { const r = tryFx(ctl.fx.Items, { project, rt }, null); if (Array.isArray(r)) return r; if (r?.__list) return [...(rt.db[r.__list] || [])]; }
  if (!d.list && d.previewList) return [...(rt?.db[d.previewList] || [])];
  let rows = [...(rt?.db[d.list] || [])];
  if (!d.list) return [1, 2, 3, 4, 5].map((n) => ({ Value: n }));
  if (d.searchCtl && d.searchField) {
    const q = String(rt.inputs[d.searchCtl] ?? '').toLowerCase();
    const col = colOf(project, d.list, d.searchField);
    if (q) rows = rows.filter((r) => displayValue(col, r[d.searchField]).toLowerCase().startsWith(q));
  }
  if (d.filterField === 'CreatedByMe') rows = rows.filter((_, i) => i % 2 === 0);
  else if (d.filterField) {
    const col = colOf(project, d.list, d.filterField);
    if (col?.type === 'Person') rows = rows.filter((r) => r[d.filterField]?.DisplayName === PEOPLE[0].DisplayName);
    else if (col?.type === 'YesNo') rows = rows.filter((r) => !!r[d.filterField] === /^(true|yes|1)$/i.test(d.filterValue));
    else rows = rows.filter((r) => String(r[d.filterField] ?? '') === String(d.filterValue ?? ''));
  }
  if (d.sortField) {
    const dir = d.sortDir === 'Ascending' ? 1 : -1;
    rows.sort((a, b) => (a[d.sortField] > b[d.sortField] ? dir : a[d.sortField] < b[d.sortField] ? -dir : (a.ID - b.ID) * dir));
  }
  return rows;
}

// ============================================================ control rendering
export function textOf(ctl, env) {
  const { project, row, rt } = env;
  if (ctl.fx?.Text) {
    const v = tryFx(ctl.fx.Text, { project, row, rt }, undefined);
    return v === undefined ? (env.mode === 'play' ? '' : 'ƒx  ' + ctl.fx.Text) : toText(v);
  }
  const b = ctl.bind;
  if (b) {
    if (b.kind === 'user') return (b.prefix || '') + (b.field === 'Email' ? 'alex.rivera@contoso.com' : 'Alex Rivera');
    if (b.kind === 'count') {
      const rows = rt ? rt.db[b.list] || [] : sampleRows(project.lists.find((l) => l.name === b.list) || { columns: [] }, project.lists, 8);
      const n = b.field ? rows.filter((r) => String(r[b.field]) === String(b.value) || (typeof r[b.field] === 'boolean' && r[b.field] === /^(true|yes|1)$/i.test(b.value))).length : rows.length;
      return (b.prefix || '') + n;
    }
    const col = colOf(project, b.list, b.field);
    let rec = b.src === 'item' ? row : rt ? rt.vars[varSel(b.list)] : null;
    if (!rec && env.mode === 'design') rec = env.sample?.[b.list]?.[0];
    const v = displayValue(col, rec?.[b.field]);
    return (b.prefix || '') + (v || (env.mode === 'design' && !rec ? `{${b.field}}` : ''));
  }
  return ctl.props.text ?? '';
}

function isVisible(ctl, env) {
  if (env.mode !== 'play') return true;
  if (ctl.fx?.Visible) return !!tryFx(ctl.fx.Visible, env, true);
  const v = ctl.visibleWhen;
  if (!v?.var) return true;
  const cur = env.rt.vars[v.var];
  switch (v.op) {
    case 'false': return !cur;
    case 'eq': return toText(cur) === String(v.value ?? '');
    case 'neq': return toText(cur) !== String(v.value ?? '');
    default: return !!cur;
  }
}

/** Build the DOM for one control. env: {project, theme, mode, row, rt, sample, screen, rowIndex} */
export function renderControl(ctl, env) {
  const { theme, mode } = env;
  const C = (v) => cssColor(v, theme);
  const el = document.createElement('div');
  el.className = 'pa-ctl pa-' + ctl.type;
  el.dataset.id = ctl.id;
  el.style.cssText = `left:${ctl.x}px;top:${ctl.y}px;width:${ctl.w}px;height:${ctl.h}px;`;
  const p = ctl.props || {};
  const font = FONT_CSS[theme.font] || FONT_CSS['Segoe UI'];
  const rad = (r) => (r === 'theme' || r == null ? theme.radius : Number(r)) + 'px';
  if (!isVisible(ctl, env)) { el.style.display = 'none'; return el; }
  if (mode === 'design' && (ctl.visibleWhen?.var || ctl.fx?.Visible)) el.classList.add('pa-cond');
  const play = mode === 'play';
  const click = (handler) => { if (play) { el.style.cursor = 'pointer'; el.addEventListener('click', (e) => { e.stopPropagation(); handler(); }); } };
  const runSel = () => click(() => env.rt.run(env.screen, ctl.actions?.OnSelect, env.row));
  switch (ctl.type) {
    case 'label': {
      el.style.cssText += `color:${C(p.color)};background:${C(p.fill)};font:${p.italic ? 'italic ' : ''}${weightCss(p.weight)} ${p.size * PT}px/1.25 ${font};justify-content:${flexAlign(p.align)};text-align:${(p.align || 'Left').toLowerCase()};align-items:${p.valign === 'Top' ? 'flex-start' : p.valign === 'Bottom' ? 'flex-end' : 'center'};`;
      el.innerHTML = `<span>${esc(textOf(ctl, env))}</span>`;
      if (hasActions(ctl)) runSel();
      break;
    }
    case 'button': case 'badge': {
      let fill = C(p.fill);
      const txt = textOf(ctl, env);
      if (ctl.type === 'badge') {
        const m = String(p.colorMap || '').split('\n').map((l) => l.split('=')).find(([k]) => k?.trim() && k.trim() === txt);
        if (m) fill = C(m[1].trim());
      }
      el.style.cssText += `background:${fill};color:${C(p.color)};border:${Number(p.borderThickness) || 0}px solid ${C(p.borderColor)};border-radius:${rad(p.radius)};font:${weightCss(p.weight)} ${p.size * PT}px/1.2 ${font};`;
      el.innerHTML = `<span>${esc(txt)}</span>`;
      if (play && ctl.type === 'button') {
        const save = (ctl.actions?.OnSelect || []).find((a) => a.type === 'save');
        const disabled = save && missingRequired(env.project, env.rt, env.screen, save.list);
        if (disabled) el.classList.add('pa-disabled');
        else runSel();
      }
      if (mode === 'design' && ctl.type === 'button' && !txt && (!p.fill || p.fill === 'transparent')) el.classList.add('pa-ghost');
      break;
    }
    case 'icon':
      el.innerHTML = svgIcon(p.icon, C(p.color), p.stroke || 2);
      if (hasActions(ctl)) runSel();
      break;
    case 'card': {
      const shadow = !!p.shadow;
      const bg = p.fill2 ? `linear-gradient(${p.gradAngle ?? 135}deg,${C(p.fill)},${C(p.fill2)})` : C(p.fill);
      el.innerHTML = `<div class="pa-card-in" style="inset:${shadow ? 6 : 0}px;background:${bg};border-radius:${rad(p.radius)};${Number(p.borderThickness) > 0 ? `border:${p.borderThickness}px solid ${C(p.borderColor)};` : ''}${shadow ? 'box-shadow:0 2px 6px rgba(15,23,42,.14);' : ''}"></div>`;
      if (hasActions(ctl)) runSel();
      break;
    }
    case 'rect':
      el.style.cssText += `background:${C(p.fill)};${Number(p.borderThickness) > 0 ? `box-shadow:inset 0 0 0 ${p.borderThickness}px ${C(p.borderColor)};` : ''}`;
      if (hasActions(ctl)) runSel();
      break;
    case 'image': {
      const fxSrc = ctl.fx?.Image ? tryFx(ctl.fx.Image, env, '') : '';
      if (fxSrc && /^(data:|https?:)/.test(fxSrc)) { el.innerHTML = `<img src="${esc(fxSrc)}" alt="" style="object-fit:contain">`; if (hasActions(ctl)) runSel(); break; }
      el.innerHTML = p.src ? `<img src="${esc(p.src)}" alt="" style="object-fit:${{ Fit: 'contain', Fill: 'cover', Stretch: 'fill', Center: 'none', Tile: 'none' }[p.fit] || 'contain'}">` : `<div class="pa-img-ph">${svgIcon('file', '#94a3b8', 1.5)}<span>Image</span></div>`;
      if (hasActions(ctl)) runSel();
      break;
    }
    case 'html':
      { const src = ctl.fx?.HtmlText ? tryFx(ctl.fx.HtmlText, env, null) : p.html;
        el.innerHTML = `<div class="pa-html" style="font-family:${font};color:${C('@Text')}">${src == null ? '<code style="opacity:.6">ƒx HtmlText</code>' : sanitizeHtml(String(src))}</div>`; }
      if (hasActions(ctl)) runSel();
      break;
    case 'textinput': case 'person': {
      const base = `background:${C('@InputFill')};border:1px solid ${C('@InputBorder')};color:${C('@Text')};font:${p.size * PT}px ${font};`;
      el.style.cssText += base;
      const multiline = ctl.field ? colOf(env.project, ctl.field.list, ctl.field.column)?.type === 'Note' : p.multiline;
      if (play) {
        const val = inputValue(env.project, env.rt, ctl);
        const shown = ctl.type === 'person' ? (val?.DisplayName ?? '') : val;
        const inp = document.createElement(multiline ? 'textarea' : 'input');
        inp.value = ctl.type === 'person' && ctl.name + '__q' in env.rt.inputs ? env.rt.inputs[ctl.name + '__q'] : shown;
        inp.placeholder = p.hint || '';
        inp.style.cssText = `font:inherit;color:inherit;`;
        if (p.numeric || ['Number', 'Currency'].includes(colOf(env.project, ctl.field?.list, ctl.field?.column)?.type)) inp.inputMode = 'decimal';
        inp.addEventListener('input', () => {
          if (ctl.type === 'person') { env.rt.inputs[ctl.name + '__q'] = inp.value; env.rt.inputs[ctl.name] = null; }
          else env.rt.inputs[ctl.name] = inp.value;
          env.rt.ctlValues[ctl.name] = { __ctl: true, get: (m) => (m === 'Text' ? inp.value : undefined) };
          env.rerenderSoon?.(ctl.name);
        });
        el.appendChild(inp);
        if (ctl.type === 'person') {
          const qv = env.rt.inputs[ctl.name + '__q'];
          if (qv && qv.length > 1 && !env.rt.inputs[ctl.name]) {
            const box = document.createElement('div');
            box.className = 'pa-people';
            box.style.cssText = `background:${C('@Surface')};border:1px solid ${C('@Border')};color:${C('@Text')};`;
            PEOPLE.filter((x) => x.DisplayName.toLowerCase().includes(qv.toLowerCase())).slice(0, 5).forEach((person) => {
              const r = document.createElement('div'); r.textContent = `${person.DisplayName}  -  ${person.Mail}`;
              r.addEventListener('mousedown', (e) => { e.preventDefault(); env.rt.inputs[ctl.name] = person; delete env.rt.inputs[ctl.name + '__q']; env.rt.hooks?.rerender(); });
              box.appendChild(r);
            });
            el.appendChild(box); el.style.overflow = 'visible'; el.style.zIndex = 50;
          }
        }
      } else {
        const v = ctl.field ? '' : p.defaultText;
        el.innerHTML = `<span class="${v ? '' : 'pa-hint'}">${esc(v || p.hint || '')}</span>${ctl.type === 'person' ? `<i class="pa-pp">${svgIcon('user', C('@TextMuted'), 2)}</i>` : ''}`;
        if (multiline) el.style.alignItems = 'flex-start';
      }
      break;
    }
    case 'dropdown': {
      const col = ctl.field ? colOf(env.project, ctl.field.list, ctl.field.column) : null;
      const items = col?.type === 'Choice' ? col.choices : col?.type === 'Lookup' ? ((env.rt?.db[col.lookupList] || env.sample?.[col.lookupList] || []).map((r) => r.Title)) : String(p.items || '').split('\n').filter(Boolean);
      el.style.cssText += `background:${C('@InputFill')};border:1px solid ${C('@InputBorder')};color:${C('@Text')};font:${p.size * PT}px ${font};`;
      const chev = `<i class="pa-chev" style="background:${C('@Primary')}">${svgIcon('chevron', C('@OnPrimary'), 2.4)}</i>`;
      if (play) {
        const sel = document.createElement('select');
        const val = inputValue(env.project, env.rt, ctl);
        items.forEach((it) => { const o = document.createElement('option'); o.textContent = it; if (it === val) o.selected = true; sel.appendChild(o); });
        sel.addEventListener('change', () => { env.rt.inputs[ctl.name] = sel.value; env.rt.run(env.screen, ctl.actions?.OnChange, env.row); });
        el.appendChild(sel); el.insertAdjacentHTML('beforeend', chev);
      } else el.innerHTML = `<span>${esc(items[0] || '')}</span>${chev}`;
      break;
    }
    case 'datepicker': {
      el.style.cssText += `background:${C('@InputFill')};border:1px solid ${C('@InputBorder')};color:${C('@Text')};font:${p.size * PT}px ${font};`;
      const ico = `<i class="pa-chev" style="background:${C('@Primary')}">${svgIcon('calendar', C('@OnPrimary'), 2)}</i>`;
      if (play) {
        const d = document.createElement('input'); d.type = 'date'; d.value = inputValue(env.project, env.rt, ctl) || '';
        d.addEventListener('change', () => { env.rt.inputs[ctl.name] = d.value; env.rt.hooks?.rerender(); });
        el.appendChild(d); el.insertAdjacentHTML('beforeend', ico);
      } else el.innerHTML = `<span>${displayValue({ type: 'DateTime' }, new Date().toISOString().slice(0, 10))}</span>${ico}`;
      break;
    }
    case 'checkbox': case 'toggle': {
      const on = play ? inputValue(env.project, env.rt, ctl) : !!p.defaultOn;
      if (ctl.type === 'checkbox') {
        el.style.cssText += `color:${C(p.color)};font:${p.size * PT}px ${font};`;
        el.innerHTML = `<i class="pa-box" style="border-color:${C('@InputBorder')};background:${on ? C('@Primary') : C('@InputFill')}">${on ? svgIcon('check', C('@OnPrimary'), 3) : ''}</i><span>${esc(p.text)}</span>`;
      } else {
        el.innerHTML = `<i class="pa-tgl" style="background:${on ? C('@Primary') : '#c8c8c8'}"><b style="left:${on ? 'calc(100% - 22px)' : '2px'}"></b></i><span style="color:${C('@Text')};font:${12 * PT}px ${font}">${on ? 'On' : 'Off'}</span>`;
      }
      if (play) click(() => { env.rt.inputs[ctl.name] = !on; env.rt.run(env.screen, ctl.actions?.[on ? 'OnUncheck' : 'OnCheck'], env.row); });
      break;
    }
    case 'gallery': {
      const rowH = Math.max(8, Number(p.templateSize) || 60);
      const rows = play ? galleryRowsFor(env.project, env.rt, ctl) : (env.sample?.[ctl.data?.list || ctl.data?.previewList] || [{}, {}, {}, {}]);
      const inner = document.createElement('div'); inner.className = 'pa-gal-in';
      const n = play ? rows.length : Math.max(1, Math.min(rows.length, Math.ceil(ctl.h / rowH)));
      for (let i = 0; i < n; i++) {
        const row = rows[i] || {};
        const r = document.createElement('div');
        r.className = 'pa-gal-row' + (!play && i > 0 ? ' pa-ghost-row' : '') + (!play && i === 0 ? ' pa-tpl' : '');
        const selected = play ? env.rt.galSel[ctl.name] === row.ID : false;
        r.style.cssText = `height:${rowH}px;background:${C(selected ? p.selectedFill : p.templateFill)};`;
        for (const ch of ctl.children || []) r.appendChild(renderControl(ch, { ...env, row, rowIndex: i, inGallery: ctl }));
        if (play && hasActions(ctl)) {
          r.style.cursor = 'pointer';
          r.addEventListener('click', () => { env.rt.galSel[ctl.name] = row.ID; env.rt.run(env.screen, ctl.actions.OnSelect, row); });
        }
        inner.appendChild(r);
      }
      if (play && !rows.length) inner.innerHTML = `<div class="pa-empty" style="color:${C('@TextMuted')}">No items</div>`;
      el.appendChild(inner);
      if (!play) el.classList.toggle('pa-gal-empty', !(ctl.children || []).length);
      break;
    }
    case 'timer':
      if (play) { el.style.display = 'none'; break; }
      el.innerHTML = `<i>${svgIcon('clock', '#64748b', 2)}</i><span>Timer</span>`;
      break;
    case 'raw':
      el.innerHTML = `<span>${esc(ctl.paControl)}</span><small>${esc(ctl.name)}</small>`;
      break;
  }
  return el;
}
const weightCss = (w) => ({ Lighter: 300, Normal: 400, Semibold: 600, Bold: 700 }[w] || 400);
const flexAlign = (a) => ({ Left: 'flex-start', Center: 'center', Right: 'flex-end', Justify: 'flex-start' }[a] || 'flex-start');

export function missingRequired(project, rt, screen, list) {
  for (const [c, parent] of walk(screen.controls)) {
    if (parent || c.field?.list !== list) continue;
    const col = colOf(project, list, c.field.column);
    if (!col?.required || c.type === 'checkbox' || c.type === 'toggle') continue;
    const v = inputValue(project, rt, c);
    if (v == null || (typeof v === 'string' && !v.trim())) return true;
  }
  return false;
}

/** Render a full screen into `host` (sized to the form factor). */
export function renderScreen(host, screen, env) {
  const ff = FORM_FACTORS[env.project.app.formFactor || 'tablet'];
  host.innerHTML = '';
  host.style.width = ff.w + 'px'; host.style.height = ff.h + 'px';
  host.style.background = cssColor(screen.fill || '@Bg', env.theme);
  host.style.fontFamily = FONT_CSS[env.theme.font] || FONT_CSS['Segoe UI'];
  for (const ctl of screen.controls) host.appendChild(renderControl(ctl, { ...env, screen }));
}
