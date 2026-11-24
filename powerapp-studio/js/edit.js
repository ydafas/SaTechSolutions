// edit.js — targeted edits to an existing project. Used by the Data tab (renames) and by
// "Ask for a change" (Claude edit operations). Every function changes only what it must:
// screens it does not touch keep byte-identical YAML.
import { walk, slug, varSel, q, newColumn, SP_RESERVED_COLS, COLTYPES, IDENT, FORM_FACTORS } from './core.js';
import { inputFor, sectionOf, sectionCard, layoutCtx, galleryRows, stepperFx, summaryFx, statusCol } from './templates.js';
import { missingFx } from './emit.js';

const pretty = (n) => String(n).replace(/([a-z])([A-Z])/g, '$1 $2');
const escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
export const cleanName = (n, fallback) => String(n || '').replace(/[^A-Za-z0-9]+(.)?/g, (_, c) => (c ? c.toUpperCase() : '')).replace(/^[^A-Za-z]+/, '') || fallback;
const BUILTIN_COLS = new Set(['Title', 'ID', 'Id', 'Created', 'Modified', 'Author', 'Editor']);

// ------------------------------------------------------------------ formula rewriting
/** Apply fn to every code part (outside "strings" and // comments) of a Power Fx formula. */
export function mapCode(f, fn) {
  const src = String(f ?? '');
  let out = '', i = 0, start = 0;
  const flush = (end) => { if (end > start) out += fn(src.slice(start, end)); };
  while (i < src.length) {
    if (src[i] === '"') {
      flush(i); let j = i + 1;
      while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += src.slice(i, j + 1); i = j + 1; start = i; continue;
    }
    if (src.startsWith('//', i)) { flush(i); let j = src.indexOf('\n', i); j = j < 0 ? src.length : j; out += src.slice(i, j); i = j; start = i; continue; }
    i++;
  }
  flush(src.length);
  return out;
}
/** Visit every formula-bearing field of the project. fn(text) returns the new text. */
export function rewriteFormulas(project, fn) {
  const fix = (v) => (typeof v === 'string' && v ? fn(v) : v);
  for (const s of project.screens) {
    s.onVisible = fix(s.onVisible);
    if (s.fillFx) s.fillFx = fix(s.fillFx);
    for (const [c] of walk(s.controls)) {
      for (const k of Object.keys(c.fx || {})) c.fx[k] = fix(c.fx[k]);
      if (c.data?.extraFilter) c.data.extraFilter = fix(c.data.extraFilter);
      for (const acts of Object.values(c.actions || {})) for (const a of acts || []) {
        if (a.type === 'custom') a.fx = fix(a.fx);
        if (a.type === 'setVar') { a.value = fix(a.value); a.name = fix(a.name); }
        if (a.type === 'toggleVar') a.name = fix(a.name);
      }
      if (c.visibleWhen?.var) c.visibleWhen.var = fix(c.visibleWhen.var);
      if (c.data?.scopeVar) c.data.scopeVar = fix(c.data.scopeVar);
      if (c.type === 'raw') for (const k of Object.keys(c.rawProps || {})) c.rawProps[k] = fix(c.rawProps[k]);
    }
  }
  project.app.onStart = fix(project.app.onStart);
  project.app.formulas = fix(project.app.formulas);
}

// ------------------------------------------------------------------ renames
/** Rename a list everywhere: model, bindings, actions, flows, and the formulas that were generated with its name. */
export function renameListIn(project, oldN, newN) {
  newN = String(newN || '').trim();
  if (!newN || newN === oldN) return 'Same name.';
  if (project.lists.some((l) => l.name === newN)) return 'A list with that name exists.';
  const l = project.lists.find((x) => x.name === oldN); if (!l) return `No list called ${oldN}.`;
  l.name = newN;
  for (const x of project.lists) for (const c of x.columns) if (c.lookupList === oldN) c.lookupList = newN;
  for (const s of project.screens) for (const [c] of walk(s.controls)) {
    if (c.field?.list === oldN) c.field.list = newN; if (c.bind?.list === oldN) c.bind.list = newN; if (c.data?.list === oldN) c.data.list = newN;
    if (c.data?.previewList === oldN) c.data.previewList = newN;
    Object.values(c.actions || {}).flat().forEach((a) => { if (a.list === oldN) a.list = newN; });
  }
  for (const f of project.flows || []) if (f.list === oldN) f.list = newN;
  if (project.app.deepLink?.list === oldN) project.app.deepLink.list = newN;
  if (project.app.approvalsCfg?.list === oldN) project.app.approvalsCfg.list = newN;
  if (project.app.chain?.list === oldN) project.app.chain.list = newN;
  // generated formulas: the table name, [@List], 'List' and the per-list variables (varSelRequests, varScopeRequests…)
  const so = slug(oldN), sn = slug(newN);
  const bare = IDENT.test(oldN) ? new RegExp(`(^|[^.\\w'])${escRe(oldN)}(?![\\w'])`, 'g') : null;
  const quoted = new RegExp(`'${escRe(oldN.replace(/'/g, "''"))}'`, 'g');
  const vars = new RegExp(`\\b((?:var|loc|col)[A-Za-z]*?)${escRe(so)}\\b`, 'g');
  rewriteFormulas(project, (f) => mapCode(f, (code) => {
    let c = code.replace(vars, (m, p) => p + sn).replace(quoted, q(newN));
    if (bare) c = c.replace(bare, (m, p) => p + q(newN));
    return c;
  }));
  return null;
}
/** Rename a column everywhere: model, bindings, flows, approvals and the generated formulas (ThisItem.Old, varSel.Old, Old.Value). */
/** Columns the approval chain generates (per-stage decision/date/by/comments, CurrentStage, ApprovalHistory): renamed only through the stage. */
const chainManaged = (project, list, col) => project.app.chain?.list === list && !!col.chain;
export function renameColumnIn(project, list, oldN, newN) {
  newN = String(newN || '').trim();
  const l = project.lists.find((x) => x.name === list); if (!l) return `No list called ${list}.`;
  const col = l.columns.find((c) => c.name === oldN); if (!col) return `No column ${oldN} in ${list}.`;
  if (col.builtin) return 'The Title column cannot be renamed.';
  if (chainManaged(project, list, col)) return `${oldN} belongs to the approval chain; its name follows the stage name.`;
  if (!newN || newN === oldN) return 'Same name.';
  if (l.columns.some((c) => c.name === newN)) return 'That column exists.';
  col.name = newN;
  const uses = new Set();
  for (const s of project.screens) {
    let touches = false;
    for (const [c] of walk(s.controls)) {
      if (c.field?.list === list && c.field.column === oldN) { c.field.column = newN; touches = true; }
      if (c.bind?.list === list && c.bind.field === oldN) { c.bind.field = newN; touches = true; }
      if (c.data?.list === list) {
        for (const k of ['searchField', 'filterField', 'sortField', 'statusField']) if (c.data[k] === oldN) c.data[k] = newN;
        if (c.data.metaCols) c.data.metaCols = c.data.metaCols.map((x) => (x === oldN ? newN : x));
        touches = true;
      }
      Object.values(c.actions || {}).flat().forEach((a) => { if (a.list === list && a.field === oldN) { a.field = newN; touches = true; } });
    }
    if (touches) uses.add(s.id);
  }
  for (const f of project.flows || []) {
    if (f.list !== list) continue;
    if (f.statusField === oldN) f.statusField = newN;
    for (const r of f.recipients) if (r.kind === 'column' && r.value === oldN) r.value = newN;
    const tok = new RegExp(`\\{${escRe(oldN)}(\\.Email)?\\}`, 'g');
    f.subject = f.subject.replace(tok, (m, e) => `{${newN}${e || ''}}`); f.body = f.body.replace(tok, (m, e) => `{${newN}${e || ''}}`);
  }
  if (project.app.approvalsCfg?.list === list && project.app.approvalsCfg.field === oldN) project.app.approvalsCfg.field = newN;
  const ch = project.app.chain;
  if (ch?.list === list) { if (ch.field === oldN) ch.field = newN; for (const st of ch.stages) if (st.approver === 'column' && st.column === oldN) st.column = newN; }
  // formulas on screens that work with this list: member access (.Old) and bare column names inside Filter predicates
  if (IDENT.test(oldN)) {
    const member = new RegExp(`\\.${escRe(oldN)}\\b`, 'g'), bare = new RegExp(`(^|[^.\\w])${escRe(oldN)}\\b(?=\\s*(?:\\.|=|<>|,|\\)|&&|\\|\\|))`, 'g');
    const onList = new Set(project.screens.filter((s) => uses.has(s.id) || [...walk(s.controls)].some(([c]) => c.data?.list === list || c.field?.list === list || c.bind?.list === list)).map((s) => s.id));
    for (const s of project.screens) {
      if (!onList.has(s.id)) continue;
      const fx = (v) => (typeof v === 'string' ? mapCode(v, (code) => code.replace(member, '.' + newN).replace(bare, (m, p) => p + newN)) : v);
      s.onVisible = fx(s.onVisible);
      for (const [c] of walk(s.controls)) {
        for (const k of Object.keys(c.fx || {})) c.fx[k] = fx(c.fx[k]);
        if (c.data?.extraFilter) c.data.extraFilter = fx(c.data.extraFilter);
        for (const acts of Object.values(c.actions || {})) for (const a of acts || []) if (a.type === 'custom') a.fx = fx(a.fx);
      }
    }
  }
  // captions the generator wrote ("Priority", "Priority *") and generated cards that print the column name
  const pOld = prettyName(oldN), pNew = prettyName(newN);
  for (const s of project.screens) {
    if (!uses.has(s.id) && ![...walk(s.controls)].some(([c]) => c.data?.list === list)) continue;
    for (const [c] of walk(s.controls)) if (c.type === 'label' && (c.props?.text === pOld || c.props?.text === pOld + ' *') && !c.fx?.Text) c.props.text = c.props.text.replace(pOld, pNew);
    for (const c of s.controls) {
      if (c.type === 'gallery' && c.data?.list === list && (c.children || []).some((ch) => /^htmlRow/.test(ch.name) && (ch.fx?.HtmlText || '').includes('.' + newN))) rebuildRows(project, s, c, l);
      if (/^htmlSummary/.test(c.name) && s.name.startsWith('scrDetail' + slug(list)) && (c.fx?.HtmlText || '').includes('.' + newN)) c.fx.HtmlText = summaryFx(project, l, c.w, c.h);
    }
  }
  return null;
}

// ------------------------------------------------------------------ locate generated screens
/** Screens of a kind ('edit' | 'detail' | 'browse' | 'approvals') that work with a list. */
export function screensFor(project, list, kind) {
  const pre = { edit: 'scrEdit', detail: 'scrDetail', browse: 'scrBrowse', approvals: 'scrApprovals' }[kind];
  const named = project.screens.filter((s) => s.name.startsWith(pre + slug(list)));
  if (named.length) return named;
  return project.screens.filter((s) => {
    const cs = [...walk(s.controls)].map(([c]) => c);
    if (kind === 'edit') return cs.some((c) => (c.actions?.OnSelect || []).some((a) => a.type === 'save' && a.list === list));
    if (kind === 'detail') return cs.some((c) => c.bind?.src === 'sel' && c.bind.list === list) && !cs.some((c) => c.field?.list === list);
    if (kind === 'browse') return cs.some((c) => c.type === 'gallery' && c.data?.list === list);
    return false;
  });
}
export function screenKind(s) {
  if (s.role === 'loading') return 'loading';
  const m = /^scr(Home|Browse|Detail|Edit|Approvals|Loading)/.exec(s.name);
  return m ? m[1].toLowerCase() : 'custom';
}

// ------------------------------------------------------------------ layout helpers
const prettyName = (n) => String(n).replace(/([a-z])([A-Z])/g, '$1 $2');
/** The form's "Still needed: …" strip (not a column caption such as lblNeededBy…). */
const isNeededStrip = (c) => /^lblNeeded/.test(c.name) && /varTriedSave/.test(c.fx?.Visible || '');
const isAnchored = (c, barTop) => c.y >= barTop || /^(rectBar|btnRun|btnFab)/.test(c.name) || isNeededStrip(c) || /^varDone/.test(c.visibleWhen?.var || '');
/** Move controls below `fromY` (not anchored to the bottom bar) down so they start at least at `minY`. */
function pushDown(screen, fromY, minY, barTop, skip) {
  const below = screen.controls.filter((c) => !skip.has(c) && c.y >= fromY && !isAnchored(c, barTop));
  if (!below.length) return 0;
  const top = Math.min(...below.map((c) => c.y));
  const d = Math.max(0, minY - top);
  if (d) below.forEach((c) => { c.y += d; });
  return d;
}
const inside = (c, card) => c.x >= card.x - 1 && c.x < card.x + card.w && c.y >= card.y && c.y < card.y + card.h;
/** Put newly added controls just before the bottom bar / overlays so overlays stay in front. */
function moveBeforeAnchors(screen, added, barTop) {
  const first = screen.controls.findIndex((c) => !added.includes(c) && isAnchored(c, barTop));
  if (first < 0) return;
  for (const a of added) screen.controls.splice(screen.controls.indexOf(a), 1);
  screen.controls.splice(first, 0, ...added);
}

// ------------------------------------------------------------------ add a column to generated screens
/** Add a label + bound input for a column to a list's form (edit) screen, in the right section, growing it. */
export function addFieldToForm(project, listName, colName, screenName = '') {
  const l = project.lists.find((x) => x.name === listName), col = l?.columns.find((c) => c.name === colName);
  if (!col) return { error: `No column ${colName} in ${listName}.` };
  const scr = screenName ? project.screens.find((s) => s.name === screenName) : screensFor(project, listName, 'edit')[0];
  if (!scr) return { error: `No form screen for ${listName}. Add one first (addScreenFromTemplate edit).` };
  if ([...walk(scr.controls)].some(([c]) => c.field?.list === listName && c.field.column === colName)) return { error: `${colName} is already on ${scr.name}.` };
  const H = FORM_FACTORS[project.app.formFactor || 'tablet'].h;
  const bar = scr.controls.find((c) => /^rectBar/.test(c.name) && !/Line/.test(c.name));
  const barTop = bar ? bar.y : H;
  const group = sectionOf(col);
  const cards = scr.controls.filter((c) => c.type === 'card' && /^cardSec/.test(c.name));
  let card = cards.find((c) => c.name.startsWith('cardSec' + slug(group)));
  const ctx = layoutCtx(project, scr);
  const before = scr.controls.length;
  let x, y, w;
  if (card) {
    const inputs = scr.controls.filter((c) => c.field?.list && inside(c, card));
    const labels = scr.controls.filter((c) => c.type === 'label' && inside(c, card) && !/^lblSec/.test(c.name) && inputs.some((i) => Math.abs(i.x - c.x) < 2 && i.y - c.y > 0 && i.y - c.y < 30));
    const xs = [...new Set(labels.map((c) => c.x))].sort((a, b) => a - b);
    const n = group === 'Notes' ? 1 : Math.max(1, xs.length);
    const rows = [...new Set(labels.map((c) => c.y))].sort((a, b) => a - b);
    w = labels.length ? Math.min(...labels.map((c) => c.w)) : card.w - 40;
    const lastRow = rows[rows.length - 1];
    const inLast = labels.filter((c) => c.y === lastRow).length;
    if (labels.length && inLast < n && group !== 'Notes') { x = xs[inLast]; y = lastRow; }
    else {
      const bottom = inputs.length ? Math.max(...inputs.map((c) => c.y + c.h)) : card.y + 34;
      x = xs[0] ?? card.x + 20; y = bottom + 12;
    }
    if (group === 'Notes') w = card.w - 40;
  } else {
    // new section card after the last section
    const last = cards.reduce((m, c) => (c.y + c.h > m ? c.y + c.h : m), (ctx.HH || 72) + 2);
    const before2 = scr.controls.length;
    card = sectionCard(ctx.add, 'Sec' + slug(group) + scr.name.replace(/^scr/, ''), cards[0]?.x ?? ctx.pad, last + 14, cards[0]?.w ?? (ctx.W - ctx.pad * 2), group);
    moveBeforeAnchors(scr, scr.controls.slice(before2), barTop);
    x = card.x + 20; y = card.y + 48; w = group === 'Notes' || ctx.phone ? card.w - 40 : Math.floor((card.w - 40 - 20) / 2);
  }
  const oldBottom = card.y + card.h;
  const r = inputFor(project, scr, l, col, x, y, w);
  const need = r.bottom + 20;
  const skip = new Set([card, r.lab, r.inp]);
  if (need > oldBottom) { card.h = need - card.y; pushDown(scr, oldBottom, card.y + card.h + 14, barTop, skip); }
  moveBeforeAnchors(scr, scr.controls.slice(before), barTop);
  // the "Still needed" strip lists required inputs; rebuild it so it knows about the new one
  const needed = scr.controls.find(isNeededStrip);
  if (needed) { const miss = missingFx(project, scr, listName); if (miss) { needed.fx.Text = `"  Still needed: " & ${miss.list}`; needed.fx.Visible = `varTriedSave${slug(listName)} && (${miss.any})`; } }
  const lowest = () => Math.max(...scr.controls.filter((c) => !isAnchored(c, barTop)).map((c) => c.y + c.h));
  // like the generator: when two columns no longer fit on a tablet, lay the sections out in three
  if (lowest() > barTop - 12 && !ctx.phone) reflowForm(scr, barTop, 3);
  return { screen: scr, warning: lowest() > barTop ? `${scr.name} no longer fits above the Save bar; move fields or remove one.` : '' };
}
/** Re-grid a generated form's section cards (label + input pairs) into n columns, stacking the cards again. */
function reflowForm(scr, barTop, n) {
  const cards = scr.controls.filter((c) => c.type === 'card' && /^cardSec/.test(c.name)).sort((a, b) => a.y - b.y);
  if (!cards.length) return;
  const oldBottom = Math.max(...cards.map((c) => c.y + c.h));
  let y = cards[0].y;
  for (const card of cards) {
    const members = scr.controls.filter((c) => c !== card && inside(c, card));
    const head = members.find((c) => /^lblSec/.test(c.name));
    const inputs = members.filter((c) => c.field?.list).sort((a, b) => a.y - b.y || a.x - b.x);
    const pairs = inputs.map((inp) => [members.find((c) => c.type === 'label' && c !== head && Math.abs(c.x - inp.x) < 2 && inp.y - c.y === 24), inp]);
    const strays = members.filter((c) => c !== head && !pairs.some(([l, i]) => c === l || c === i));
    const dy = y - card.y;
    card.y = y; if (head) head.y = y + 14;
    strays.forEach((c) => { c.y += dy; });
    const notes = /^cardSecNotes/.test(card.name), k = notes ? 1 : n, gap = 20, inner = card.w - 40, cw = Math.floor((inner - gap * (k - 1)) / k);
    let fy = y + 48, rowBottom = fy, i = 0;
    for (const [lab, inp] of pairs) {
      const x = card.x + 20 + (i % k) * (cw + gap);
      if (lab) { lab.x = x; lab.y = fy; lab.w = cw; }
      inp.x = x; inp.y = fy + 24; if (inp.type !== 'toggle') inp.w = cw;
      rowBottom = Math.max(rowBottom, inp.y + inp.h);
      i++; if (i % k === 0) fy = rowBottom + 12;
    }
    card.h = Math.max(rowBottom + 20 - y, 60);
    y += card.h + 14;
  }
  const d = y - 14 - oldBottom;
  if (d) scr.controls.filter((c) => c.y >= oldBottom && !isAnchored(c, barTop) && !cards.includes(c)).forEach((c) => { c.y += d; });
}

/** Add a caption + value for a column to a list's details screen. */
export function addFieldToDetail(project, listName, colName, screenName = '') {
  const l = project.lists.find((x) => x.name === listName), col = l?.columns.find((c) => c.name === colName);
  if (!col) return { error: `No column ${colName} in ${listName}.` };
  const scr = screenName ? project.screens.find((s) => s.name === screenName) : screensFor(project, listName, 'detail')[0];
  if (!scr) return { error: `No details screen for ${listName}. Add one first (addScreenFromTemplate detail).` };
  if ([...walk(scr.controls)].some(([c]) => c.bind?.src === 'sel' && c.bind.list === listName && c.bind.field === colName)) return { error: `${colName} is already on ${scr.name}.` };
  const H = FORM_FACTORS[project.app.formFactor || 'tablet'].h;
  const bar = scr.controls.find((c) => /^rectBar/.test(c.name) && !/Line/.test(c.name));
  const barTop = bar ? bar.y : H;
  const group = sectionOf(col);
  const ctx = layoutCtx(project, scr), T = scr.name.replace(/^scr/, '');
  const cards = scr.controls.filter((c) => c.type === 'card' && /^cardSec/.test(c.name));
  let card = cards.find((c) => c.name.startsWith('cardSec' + slug(group)));
  const before = scr.controls.length;
  const vis = { Visible: `!IsBlank(${varSel(listName)}.ID)` };
  let x, y, w;
  if (card) {
    const caps = scr.controls.filter((c) => /^lblCap/.test(c.name) && inside(c, card));
    const vals = scr.controls.filter((c) => /^lblVal/.test(c.name) && inside(c, card));
    const xs = [...new Set(caps.map((c) => c.x))].sort((a, b) => a - b), n = Math.max(1, xs.length);
    const rows = [...new Set(caps.map((c) => c.y))].sort((a, b) => a - b), lastRow = rows[rows.length - 1];
    const inLast = caps.filter((c) => c.y === lastRow).length;
    w = caps.length ? Math.min(...caps.map((c) => c.w)) : card.w - 40;
    if (caps.length && inLast < n && group !== 'Notes') { x = xs[inLast]; y = lastRow; }
    else { x = xs[0] ?? card.x + 20; y = (vals.length ? Math.max(...vals.map((c) => c.y + c.h)) : card.y + 36) + 10; }
    if (group === 'Notes') w = card.w - 40;
  } else {
    const last = cards.reduce((m, c) => (c.y + c.h > m ? c.y + c.h : m), (ctx.HH || 72) + 140);
    card = sectionCard(ctx.add, 'Sec' + slug(group) + T, cards[0]?.x ?? ctx.pad, last + 12, ctx.W - ctx.pad * 2, group);
    card.fx.Visible = vis.Visible;
    x = card.x + 20; y = card.y + 46; w = group === 'Notes' ? card.w - 40 : Math.floor((card.w - 40 - 32) / 3);
  }
  const oldBottom = card.y + card.h;
  const noteH = col.type === 'Note' ? 72 : 22;
  const cap = ctx.add('label', { name: 'lblCap' + slug(col.name) + T, x, y, w, h: 18, props: { text: pretty(col.name), size: 10, weight: 'Semibold', color: '@TextMuted' }, fx: { ...vis } });
  const val = ctx.add('label', { name: 'lblVal' + slug(col.name) + T, x, y: y + 18, w, h: noteH, props: { text: '-', size: 12, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: listName, field: col.name }, fx: { ...vis } });
  const need = val.y + val.h + 16;
  if (need > oldBottom) { card.h = need - card.y; pushDown(scr, oldBottom, card.y + card.h + 12, barTop, new Set([card, cap, val])); }
  moveBeforeAnchors(scr, scr.controls.slice(before), barTop);
  const bottom = Math.max(...scr.controls.filter((c) => !isAnchored(c, barTop) && (c.fx?.Visible || '').includes('!IsBlank')).map((c) => c.y + c.h));
  return { screen: scr, warning: bottom > barTop ? `${scr.name} runs past the bottom of the screen; move or remove a field.` : '' };
}

/** Show a column in the meta line of the list's generated gallery rows. */
export function addFieldToGallery(project, listName, colName, screenName = '') {
  const l = project.lists.find((x) => x.name === listName), col = l?.columns.find((c) => c.name === colName);
  if (!col) return { error: `No column ${colName} in ${listName}.` };
  const targets = [];
  for (const s of project.screens) {
    if (screenName && s.name !== screenName) continue;
    for (const c of s.controls) if (c.type === 'gallery' && c.data?.list === listName && (c.children || []).some((ch) => /^htmlRow/.test(ch.name))) targets.push([s, c]);
  }
  if (!targets.length) return { error: `No generated gallery shows ${listName}${screenName ? ' on ' + screenName : ''}.` };
  for (const [s, g] of targets) {
    g.data.metaCols = [...new Set([...(g.data.metaCols || []), colName])];
    rebuildRows(project, s, g, l);
  }
  return { screens: targets.map(([s]) => s) };
}
/** Regenerate a generated gallery's row card (keeps control names and the row tap). */
export function rebuildRows(project, screen, gal, list) {
  const compact = Number(gal.props.templateSize) === 72;
  const keep = gal.children.filter((c) => !/^(htmlRow|icoRowGo)/.test(c.name));
  gal.children = [];
  galleryRows(project, screen, gal, list, { compact });
  gal.children.push(...keep);
}
/** After a Choice column's values change: rebuild row cards, the details summary pill and the stepper of that list. */
export function refreshStatusVisuals(project, listName) {
  const l = project.lists.find((x) => x.name === listName); if (!l) return [];
  const touched = new Set();
  for (const s of project.screens) {
    for (const c of s.controls) {
      if (c.type === 'gallery' && c.data?.list === listName && (c.children || []).some((ch) => /^htmlRow/.test(ch.name))) { rebuildRows(project, s, c, l); touched.add(s); }
      if (/^htmlSummary/.test(c.name) && s.name.startsWith('scrDetail' + slug(listName))) { c.fx.HtmlText = summaryFx(project, l, c.w, c.h); touched.add(s); }
      if (/^htmlStepper/.test(c.name) && s.name.startsWith('scrDetail' + slug(listName))) { const st = statusCol(l); if (st) { c.fx.HtmlText = stepperFx(l, st, varSel(listName))(c.h); touched.add(s); } }
    }
  }
  return [...touched];
}

/** Remove a column and the controls generated for it. Returns {removed, warnings}. */
export function removeColumnIn(project, listName, colName) {
  const l = project.lists.find((x) => x.name === listName), col = l?.columns.find((c) => c.name === colName);
  if (!col) return { error: `No column ${colName} in ${listName}.` };
  if (col.builtin) return { error: 'The Title column cannot be removed.' };
  if (chainManaged(project, listName, col) || (project.app.chain?.list === listName && (project.app.chain.field === colName || project.app.chain.stages.some((st) => st.column === colName)))) return { error: `${colName} is used by the approval chain.` };
  l.columns.splice(l.columns.indexOf(col), 1);
  const removed = [], warnings = [], s0 = slug(colName);
  for (const s of project.screens) {
    const kill = new Set();
    for (const [c, parent] of walk(s.controls)) {
      if ((c.field?.list === listName && c.field.column === colName) || (c.bind?.list === listName && c.bind.field === colName && c.bind.kind !== 'count')) kill.add(c);
    }
    for (const c of s.controls) {
      if (!kill.size) break;
      if (c.type === 'label' && !c.bind && (new RegExp(`^lbl(Cap)?${s0}`).test(c.name)) && [...kill].some((k) => Math.abs(k.x - c.x) < 2 && k.y > c.y && k.y - c.y < 30)) kill.add(c);
    }
    const prune = (arr) => arr.filter((c) => { if (kill.has(c)) { removed.push(`${s.name}.${c.name}`); return false; } if (c.children) c.children = prune(c.children); return true; });
    s.controls = prune(s.controls);
    for (const c of s.controls) {
      if (c.type === 'gallery' && c.data?.list === listName) {
        for (const k of ['searchField', 'filterField', 'sortField', 'statusField']) if (c.data[k] === colName) { c.data[k] = k === 'sortField' ? 'Modified' : k === 'searchField' ? 'Title' : ''; }
        if (c.data.metaCols) c.data.metaCols = c.data.metaCols.filter((x) => x !== colName);
        if ((c.children || []).some((ch) => /^htmlRow/.test(ch.name))) rebuildRows(project, s, c, l);
      }
      if (isNeededStrip(c)) { const miss = missingFx(project, s, listName); if (miss) { c.fx.Text = `"  Still needed: " & ${miss.list}`; c.fx.Visible = `varTriedSave${slug(listName)} && (${miss.any})`; } }
    }
    const re = new RegExp(`\\.${colName}\\b`);
    for (const [c] of walk(s.controls)) for (const [k, v] of Object.entries(c.fx || {})) if (re.test(mapCode(v, (x) => x))) {
      if (/^htmlSummary/.test(c.name)) { c.fx.HtmlText = summaryFx(project, l, c.w, c.h); continue; }
      warnings.push(`${s.name}.${c.name} ${k} still refers to ${colName}; edit it under Formulas.`);
    }
  }
  for (const f of project.flows || []) {
    if (f.list !== listName) continue;
    if (f.statusField === colName) warnings.push(`Flow "${f.name}" watched ${colName}; pick another column on the Flows tab.`);
    if (f.recipients.some((r) => r.kind === 'column' && r.value === colName)) { f.recipients = f.recipients.filter((r) => !(r.kind === 'column' && r.value === colName)); warnings.push(`Flow "${f.name}" no longer emails ${colName}.`); }
    if (new RegExp(`\\{${colName}(\\.Email)?\\}`).test(f.subject + f.body)) warnings.push(`Flow "${f.name}" email still uses {${colName}}.`);
  }
  if (project.app.approvalsCfg?.list === listName && project.app.approvalsCfg.field === colName) warnings.push(`Approvals used ${colName}; the approvals screen and buttons will not work until you add it back.`);
  return { removed, warnings };
}

/** Add or remove admin-only gating on generated delete / approve / reject buttons and the approvals tile. */
export function setAdminGating(project, on) {
  const touched = new Set();
  for (const s of project.screens) for (const [c] of walk(s.controls)) {
    const gated = /^btn(Delete|Approve|Reject)/.test(c.name) || (/^btnTile/.test(c.name) && (c.actions?.OnSelect || []).some((a) => a.type === 'navigate' && project.screens.find((x) => x.id === a.screen)?.name.startsWith('scrApprovals')));
    if (!gated) continue;
    const v = c.fx.Visible || '';
    if (on && !/varIsAdmin/.test(v)) { c.fx.Visible = v ? `${v} && varIsAdmin` : 'varIsAdmin'; touched.add(s); }
    if (!on && /varIsAdmin/.test(v)) { const nv = v.replace(/\s*&&\s*varIsAdmin\b/g, '').replace(/^varIsAdmin$/, ''); if (nv) c.fx.Visible = nv; else delete c.fx.Visible; touched.add(s); }
  }
  return [...touched];
}

/** A new column from loose input: clean name, SharePoint-safe, unique. Returns {col, note}. */
export function makeColumn(list, spec) {
  let name = cleanName(spec.name, 'Field'), note = '';
  if (BUILTIN_COLS.has(name)) return { error: `${name} is a built-in column.` };
  const single = list.name.replace(/ies$/, 'y').replace(/s$/, '');
  if (SP_RESERVED_COLS.has(name)) { note = `${name} renamed to ${single + name} (SharePoint already uses ${name}).`; name = single + name; }
  let k = 2; const base = name; while (list.columns.some((c) => c.name === name)) name = base + k++;
  let type = COLTYPES[spec.type] ? spec.type : 'Text';
  let choices = [...new Set((spec.choices || []).map((x) => String(x).trim()).filter(Boolean))];
  if (type === 'Choice' && !choices.length) { type = 'Text'; note = (note ? note + ' ' : '') + `${name} had no choices, so it is a text column.`; }
  if (type !== 'Choice') choices = [];
  return { col: newColumn(name, type, { required: type === 'YesNo' ? false : !!spec.required, choices, lookupList: type === 'Lookup' ? String(spec.lookupList || '') : '' }), note };
}
