// ui.js — Power App Studio: the designer shell.
import * as core from './core.js';
import { CATALOG, CATS, COLTYPES, THEME_COLORS, THEME_PRESETS, FONTS, FORM_FACTORS, clone, uid, walk, findCtl, newControl, newScreen, newList, newColumn, allNames, IDENT, RESERVED, sampleRows, varSel, companions, INPUT_TYPES, colOf, cssColor, isHex, s as fxString, renameCopy, normalizeProject } from './core.js';
import { screenYaml, themeFormulas, onStartFx, startScreenName, startScreenFx, checklist, provisioningScript, lint, ACTION_TYPES, controlYaml, usedLists } from './emit.js';
import { flowDefinition, flowFileName, flowFxName, FLOW_KINDS } from './flows.js';
import { initDescribeUI, describeDialog } from './describeui.js';
import { initChangeUI, changeDialog } from './changeui.js';
import { initTableUI, importTableDialog } from './tableui.js';
import { initDriftUI, driftDialog } from './driftui.js';
import { listToCsv } from './tabular.js';
import { CONTRAST_PAIRS, contrast, fixContrast, AA } from './a11y.js';
import { renameListIn, renameColumnIn, attachNoRoom, syncAttachSlots } from './edit.js';
import { initAssistUI, setupDialog, setupCardHtml, bindSetupCard, openSetupCount, deployView, exportGuide, startTour, tourDone, conditionalCount } from './assistui.js';
import { SETUP_CODES } from './setup.js';
import { initFlowsUI, panelFlows, addFlowDialog, flowDialog, downloadFlow, duplicateFlow, deleteFlow, loadFlowSample, flowZipBlob } from './flowsui.js';
import { renderControl, renderScreen, createRuntime, svgIcon, FONT_CSS, tryFx } from './render.js';
import { buildScenario, scenarioIncludes, addScenarioFlows } from './scenarios.js';
import { SCENARIOS, generateApp, SCREEN_TEMPLATES, genBlank, genHome, genBrowse, genDetail, genEdit, genApprovals, genLoading, BLOCKS, insertBlock, galleryRows, buildDemo } from './templates.js';
import { importYaml } from './import.js';
import { MAX_STAGES, isGuid, stageCols } from './chain.js';
import { addChainFlows } from './ai.js';
import { encodeShare, decodeShare, isShareHash, WARN_CHARS, MAX_CHARS } from './share.js';
import { loadBlocks, saveBlocks, makeBlock, insertSavedBlock, exportBlocksJson, importBlocksJson } from './blocks.js';
import { ICONS, ICON_NAMES } from './icons.js';
import { SIZES } from './delegation.js';
import { readVersions, writeVersions, versionsSig } from './versionstore.js';
import { initFeedbackUI, feedbackDialog } from './feedbackui.js';
import { initHistoryUI, historyDialog } from './historyui.js';
import { takeSnapshot, mergeSnapshots } from './history.js';
import { bundleZip, readBundle, mergeProjects } from './bundle.js';

// ------------------------------------------------------------------ small helpers
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (t) => String(t ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const UI_PATHS = {
  bolt: 'M13 2L4 14h7l-1 8 9-12h-7l1-8z', undo: 'M9 7L4 12l5 5M4 12h11a5 5 0 0 1 0 10h-3', redo: 'M15 7l5 5-5 5M20 12H9a5 5 0 0 0 0 10h3',
  layers: 'M12 3l9 5-9 5-9-5 9-5zM3 13l9 5 9-5M3 17.5l9 5 9-5', data: 'M4 6c0-1.7 3.6-3 8-3s8 1.3 8 3-3.6 3-8 3-8-1.3-8-3zM4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3',
  palette: 'M12 3a9 9 0 1 0 0 18c1.2 0 1.8-.8 1.8-1.7 0-1.3-1.2-1.6-1.2-2.7 0-.9.7-1.6 1.6-1.6H17a4 4 0 0 0 4-4C21 6.6 17 3 12 3zM7.5 11h.01M10 7h.01M15 7h.01',
  copy: 'M8 8h12v12H8zM4 16V4h12', up: 'M12 19V5M6 11l6-6 6 6', down: 'M12 5v14M6 13l6 6 6-6', front: 'M4 14h6v6H4zM9 4h11v11H9z',
  label: 'M5 6h14M12 6v13', button: 'M3 8h18v8H3zM8 12h8', input: 'M3 7h18v10H3zM7 10v4', dropdown: 'M3 7h18v10H3zM14 11l2 2 2-2',
  checkbox: 'M4 4h16v16H4zM8 12l3 3 5-6', toggle: 'M2 12a5 5 0 0 1 5-5h10a5 5 0 0 1 0 10H7a5 5 0 0 1-5-5zM17 9.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z',
  rect: 'M4 6h16v12H4z', card: 'M7 5h10a3 3 0 0 1 3 3v8a3 3 0 0 1-3 3H7a3 3 0 0 1-3-3V8a3 3 0 0 1 3-3z', html: 'M8 7l-5 5 5 5M16 7l5 5-5 5', image: 'M4 5h16v14H4zM4 15l5-5 4 4 3-3 4 4M15 9h.01',
  gallery: 'M4 4h16v5H4zM4 11h16v5H4zM4 18h16', timer: 'M12 5a8 8 0 1 0 0 16 8 8 0 0 0 0-16zM12 9v4l2.5 2M9 2h6', badge: 'M5 8h14a4 4 0 0 1 0 8H5a4 4 0 0 1 0-8z',
  icon: ICONS.star, person: ICONS.user, datepicker: ICONS.calendar, raw: 'M8 4H5v16h3M16 4h3v16h-3', screen: 'M4 5h16v12H4zM9 21h6M12 17v4',
};
const ui = (name, color = 'currentColor', w = 2) => `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="${UI_PATHS[name] || ICONS[name] || ICONS.info}" fill="none" stroke="${color}" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const pretty = (n) => String(n).replace(/([a-z])([A-Z])/g, '$1 $2');
function toast(msg, err) { const t = document.createElement('div'); t.className = 'toast' + (err ? ' err' : ''); t.textContent = msg; $('#toastRoot').appendChild(t); setTimeout(() => t.remove(), 2600); }
async function copyText(text) {
  try { await navigator.clipboard.writeText(text); toast('Copied to clipboard'); return true; } catch { /* fall back */ }
  const ta = document.createElement('textarea'); ta.value = text; ta.style.cssText = 'position:fixed;opacity:0'; document.body.appendChild(ta); ta.select();
  let ok = false; try { ok = document.execCommand('copy'); } catch { /* ignore */ }
  ta.remove(); toast(ok ? 'Copied to clipboard' : 'Copy was blocked. Select the text and press Ctrl+C.', !ok); return ok;
}
function download(name, data, type = 'text/plain') {
  try {
    const blob = data instanceof Blob ? data : new Blob([data], { type });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    toast('Downloading ' + name);
  } catch { toast('Download blocked here. Use Copy instead.', true); }
}

// ------------------------------------------------------------------ state
const KEY = 'pas.project.v1';
let P = null;
const S = { tab: 'insert', screenId: null, sel: [], zoom: 'fit', undo: [], redo: [], last: '', clip: null, actEvent: 'OnSelect', blockList: '', codeTab: 'deploy', showCond: false };
function loadSaved() { try { const t = localStorage.getItem(KEY); return t ? JSON.parse(t) : null; } catch { return null; } }
/** Save to this browser. Returns false (and says so once) when storage is full, e.g. after many saved versions. */
// Version history is stored apart from the project (versionstore.js). Until the first successful write of the
// versions, the project keeps them inside (the old format), so nothing is lost if this tab closes mid-migration.
const VS = { ready: false, sig: null, chain: Promise.resolve(), warned: false };
function syncVersions() {
  const sig = versionsSig(P.snapshots);
  if (sig === VS.sig) return;
  VS.sig = sig;
  const list = P.snapshots || [];
  VS.chain = VS.chain.then(() => writeVersions(list)).then((where) => {
    if (!where && !VS.warned) { VS.warned = true; toast('Versions could not be saved in this browser. Save a project file to keep them.', true); }
  });
}
/** After loading the saved project: bring in the stored versions (and move old in-project ones out of it). */
function attachVersions(hadSaved = true) {
  VS.ready = false;
  // no saved project (first run, or storage cleared): versions left in IndexedDB belong to nothing, so drop them
  return (hadSaved ? readVersions() : Promise.resolve([])).then(async (stored) => {
    const merged = mergeSnapshots(P.snapshots || [], stored);
    if (merged.length) P.snapshots = merged; else delete P.snapshots;
    const where = await writeVersions(merged);
    if (where) { VS.ready = true; VS.sig = versionsSig(P.snapshots); persist(); }
  });
}
function persist() {
  if (S.previewing) return;
  if (S.viewer) return true;
  if (VS.ready) syncVersions();
  try { localStorage.setItem(KEY, VS.ready ? JSON.stringify({ ...P, snapshots: undefined }) : JSON.stringify(P)); S.storeWarned = false; return true; }
  catch {
    if (!S.storeWarned) toast('This browser\'s storage is full: changes are not being kept. Delete old versions (Project → Version history) or save a project file.', true);
    S.storeWarned = true; return false;
  }
}
/** Version history, the tenant findings log and deploy records are not part of Undo: they describe the tenant and the past, not the app. */
const META = ['snapshots', 'findings', 'deployed'];
function ser(p) { const o = {}; for (const k of Object.keys(p)) if (!META.includes(k)) o[k] = p[k]; return JSON.stringify(o); }
function fromUndo(text) { const q = JSON.parse(text); for (const k of META) { if (P[k]) q[k] = P[k]; else delete q[k]; } return q; }
const screen = () => P.screens.find((s) => s.id === S.screenId) || P.screens[0];
const selCtls = () => S.sel.map((id) => findCtl(P, id)).filter(Boolean);
const one = () => (S.sel.length === 1 ? findCtl(P, S.sel[0]) : null);
const ff = () => FORM_FACTORS[P.app.formFactor || 'tablet'];

function setProject(p, keepUndo = false) {
  P = p;
  migrate(P);
  S.screenId = P.screens[0]?.id; S.sel = [];
  if (!keepUndo) { S.undo = []; S.redo = []; }
  S.last = ser(P); persist(); renderAll(); zoomFit();
}
/** Bring any saved or shared project up to the current model. Safe to run on untrusted files. */
function migrate(p) {
  normalizeProject(p);
  if (!p.screens.length) p.screens.push(newScreen(p, 'scrHome'));
  p.version = 1;
}
let commitTimer = null;
function commit() {
  clearTimeout(commitTimer);
  if (S.previewing) return;
  const cur = ser(P);
  if (cur !== S.last) { S.undo.push(S.last); if (S.undo.length > 120) S.undo.shift(); S.redo = []; S.last = cur; }
  persist();
  updateTop(); lintSoon();
}
const commitSoon = () => { clearTimeout(commitTimer); commitTimer = setTimeout(commit, 350); };
function undo() { commit(); if (!S.undo.length) return; S.redo.push(S.last); S.last = S.undo.pop(); P = fromUndo(S.last); afterHistory(); }
function redo() { if (!S.redo.length) return; S.undo.push(S.last); S.last = S.redo.pop(); P = fromUndo(S.last); afterHistory(); }
const isBlank = (p) => !p.lists.length && p.screens.every((s) => !s.controls.length);
/** Swap in a whole new project as ONE undo step (Describe builds, Ask for a change, imports, restores).
 *  `label` saves the current app as an automatic version first. Versions, the findings log and switched-off
 *  properties carry over unless the new project brings its own. `mergeVersions` (bundle import) joins both version
 *  lists and leaves deploy records with the app they belong to. */
function replaceProject(p, label = '', { mergeVersions = false } = {}) {
  commit();
  const prev = S.last;
  if (mergeVersions) p.snapshots = mergeSnapshots(P.snapshots, p.snapshots);
  for (const k of META) if (!p[k] && P[k] && !(mergeVersions && k === 'deployed')) p[k] = clone(P[k]); // another app's deploy records never carry into a bundle
  if (P.app.suppressProps && p.app && !p.app.suppressProps) p.app.suppressProps = clone(P.app.suppressProps);
  if (label && !isBlank(P)) takeSnapshot(p, label, { auto: true, source: P });
  if (!p.snapshots?.length) delete p.snapshots;
  P = p; migrate(P);
  S.undo.push(prev); if (S.undo.length > 120) S.undo.shift(); S.redo = [];
  S.last = ser(P);
  if (!P.screens.find((s) => s.id === S.screenId)) S.screenId = P.screens[0]?.id;
  S.sel = S.sel.filter((id) => findCtl(P, id));
  persist(); renderAll();
}
function afterHistory() { if (!P.screens.find((s) => s.id === S.screenId)) S.screenId = P.screens[0]?.id; S.sel = S.sel.filter((id) => findCtl(P, id)); persist(); renderAll(); }

// ------------------------------------------------------------------ render all
function renderAll() { updateTop(); renderRail(); renderPanel(); renderTabs(); renderCanvas(); renderInspector(); lintSoon(); }
function updateTop() {
  const inp = $('#appName'); if (document.activeElement !== inp) inp.value = P.app.name;
  $('#btnUndo').disabled = !S.undo.length && ser(P) === S.last; $('#btnRedo').disabled = !S.redo.length;
  updateSetupBadge();
}
function updateSetupBadge() {
  let n = 0; try { n = openSetupCount(); } catch { n = 0; }
  $('#btnSetup').hidden = !n; $('#setupBadge').textContent = n;
}
let lintTimer = null;
function lintSoon() {
  clearTimeout(lintTimer);
  lintTimer = setTimeout(() => {
    const f = lint(P), e = f.filter((x) => x.level === 'error').length, w = f.filter((x) => x.level === 'warn').length;
    const b = $('#lintBadge'); b.hidden = !(e || w); b.textContent = e || w; b.classList.toggle('warn', !e);
    b.title = `${e} errors, ${w} warnings`;
  }, 300);
}
function renderRail() {
  const icons = { insert: 'plus', screens: 'screen', layers: 'layers', data: 'data', flows: 'bolt', theme: 'palette', app: 'settings' };
  $$('.rail button').forEach((b) => { b.classList.toggle('on', b.dataset.tab === S.tab); b.querySelector('i').innerHTML = ui(icons[b.dataset.tab]); });
}
function renderTabs() {
  const start = P.app.startScreen || P.screens[0]?.id;
  $('#screenTabs').innerHTML = P.screens.map((s) => `<button data-scr="${s.id}" class="${s.id === screen().id ? 'on' : ''}">${s.id === start ? '<span class="star" title="Start screen">★</span>' : ''}${esc(s.name)}</button>`).join('') +
    `<button class="add" data-cmd="newScreen" title="Add a screen">+ Screen</button>`;
}

// ------------------------------------------------------------------ canvas
let SAMPLE = {}, sampleKey = '';
/** Sample rows only change when the data model changes, so typing in the inspector stays cheap. */
function refreshSample() {
  const key = JSON.stringify(P.lists);
  if (key === sampleKey) return;
  sampleKey = key; SAMPLE = {};
  for (const l of P.lists) SAMPLE[l.name] = sampleRows(l, P.lists, 6);
}
function renderCanvas() {
  const scr = screen(); if (!scr) return;
  refreshSample();
  const host = $('#screen');
  host.classList.add('design');
  renderScreen(host, scr, designEnv());
  host.classList.toggle('show-cond', !!S.showCond);
  const n = conditionalCount(scr), chip = $('#condChip');
  chip.hidden = !n; chip.textContent = `${n} conditional`; chip.setAttribute('aria-pressed', String(!!S.showCond)); chip.classList.toggle('on', !!S.showCond);
  applyZoom();
  drawOverlay();
}
/** HTML of one screen as the designer draws it, for guide thumbnails. */
function renderThumb(scr) {
  refreshSample();
  const host = document.createElement('div');
  renderScreen(host, scr, { ...designEnv(), showHidden: false, selIds: new Set() });
  return host.innerHTML;
}
function openDemo() {
  setProject(buildDemo(core.newProject), true); commit();
  toast('Demo loaded: a finished Request Tracker with flows. Press Play to try it.');
  if (!tourDone()) setTimeout(startTour, 350);
}
function afterBuild() { if (openSetupCount()) setTimeout(setupDialog, 350); }
/** Design view previews formulas against the first sample row of each list. */
function designEnv() {
  const vars = { varIsAdmin: true, locInited: true, locStep: 2 }; for (const l of P.lists) vars[varSel(l.name)] = SAMPLE[l.name]?.[0] || {};
  for (const st of P.app.chain?.stages || []) if (st.approver === 'group') vars[stageCols(st).groupVar] = true;
  const noop = () => {};
  return { project: P, theme: P.theme, mode: 'design', sample: SAMPLE, showHidden: S.showHidden, selIds: new Set(S.sel), rt: { db: SAMPLE, vars, inputs: {}, galSel: {}, ctlValues: {}, api: { navigate: noop, back: noop, toast: noop, launch: noop } } };
}
function zoomValue() {
  if (S.zoom !== 'fit') return S.zoom;
  const vp = $('#viewport'), f = ff();
  return Math.max(0.2, Math.min(1.5, Math.min((vp.clientWidth - 64) / f.w, (vp.clientHeight - 64) / f.h)));
}
function applyZoom() {
  const z = zoomValue(), f = ff();
  $('#board').style.width = f.w * z + 'px'; $('#board').style.height = f.h * z + 'px';
  $('#scaler').style.transform = `scale(${z})`;
  $('#overlay').style.width = f.w + 'px'; $('#overlay').style.height = f.h + 'px';
  $('#zoomVal').textContent = Math.round(z * 100) + '%';
}
function zoomFit() { S.zoom = 'fit'; applyZoom(); drawOverlay(); }
function absRect(found) {
  const { ctl, parent } = found;
  if (!parent) return { x: ctl.x, y: ctl.y, w: ctl.w, h: ctl.h };
  return { x: parent.x + ctl.x, y: parent.y + ctl.y, w: ctl.w, h: ctl.h };
}
let lastSelKey = '';
function drawOverlay(guides = []) {
  // selecting or deselecting a control that is conditionally hidden re-renders so it shows while selected
  const key = S.sel.join(',');
  if (key !== lastSelKey) {
    const was = new Set(lastSelKey ? lastSelKey.split(',') : []);
    lastSelKey = key;
    const cond = (id) => { const f = findCtl(P, id); return f && (f.ctl.visibleWhen?.var || f.ctl.fx?.Visible); };
    if ([...was, ...S.sel].some(cond) && !S.showHidden) { renderCanvasQuick(); return; }
  }
  const ov = $('#overlay'); ov.innerHTML = '';
  const z = zoomValue(), items = selCtls();
  for (const f of items) {
    const r = absRect(f);
    const box = document.createElement('div');
    box.className = 'sel-box' + (items.length > 1 ? ' multi' : '') + (r.y < 24 ? ' below' : '');
    const bw = 2 / z;
    box.style.cssText = `left:${r.x}px;top:${r.y}px;width:${r.w}px;height:${r.h}px;border-width:${bw}px;`;
    if (items.length === 1) {
      box.innerHTML = `<span class="tag" style="transform:scale(${1 / z});transform-origin:${r.y < 24 ? '0 0' : '0 100%'}">${esc(f.ctl.name)}</span>`;
      for (const hd of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
        const h = document.createElement('div'); h.className = 'handle'; h.dataset.h = hd;
        const s = 10 / z, half = s / 2;
        const x = hd.includes('w') ? -half : hd.includes('e') ? r.w - half : r.w / 2 - half;
        const y = hd.includes('n') ? -half : hd.includes('s') ? r.h - half : r.h / 2 - half;
        h.style.cssText = `left:${x - bw}px;top:${y - bw}px;width:${s}px;height:${s}px;border-width:${bw}px;cursor:${hd}-resize`;
        box.appendChild(h);
      }
    }
    ov.appendChild(box);
  }
  for (const g of guides) {
    const d = document.createElement('div'); d.className = 'guide ' + g.d;
    if (g.d === 'v') d.style.cssText = `left:${g.v}px;width:${1 / z}px`; else d.style.cssText = `top:${g.v}px;height:${1 / z}px`;
    ov.appendChild(d);
  }
}
function screenPoint(e) {
  const r = $('#screen').getBoundingClientRect(), z = zoomValue();
  return { x: (e.clientX - r.left) / z, y: (e.clientY - r.top) / z };
}
function galleryAt(pt, exceptId) {
  const scr = screen();
  for (let i = scr.controls.length - 1; i >= 0; i--) {
    const g = scr.controls[i];
    if (g.type !== 'gallery' || g.id === exceptId) continue;
    const rowH = Number(g.props.templateSize) || 60;
    if (pt.x >= g.x && pt.x <= g.x + g.w && pt.y >= g.y && pt.y <= g.y + Math.min(rowH, g.h)) return g;
  }
  return null;
}

// drag state
let drag = null;
function onScreenDown(e) {
  if (e.button !== 0) return;
  const handle = e.target.closest('.handle');
  const pt = screenPoint(e);
  if (handle) { startResize(e, handle.dataset.h, pt); return; }
  const el = e.target.closest('.pa-ctl');
  if (!el || !findCtl(P, el.dataset.id)) {
    if (!e.shiftKey) S.sel = [];
    drag = { mode: 'marquee', start: pt, base: [...S.sel] };
    renderInspector(); drawOverlay(); renderPanelIf('layers');
    return;
  }
  const id = el.dataset.id;
  if (e.shiftKey || e.ctrlKey || e.metaKey) S.sel = S.sel.includes(id) ? S.sel.filter((x) => x !== id) : [...S.sel, id];
  else if (!S.sel.includes(id)) S.sel = [id];
  renderInspector(); drawOverlay(); renderPanelIf('layers');
  const items = selCtls();
  drag = { mode: 'move', start: pt, items: items.map((f) => ({ f, x: f.ctl.x, y: f.ctl.y })), moved: false };
  e.preventDefault();
}
function startResize(e, h, pt) {
  const f = one(); if (!f) return;
  drag = { mode: 'resize', h, start: pt, f, o: { x: f.ctl.x, y: f.ctl.y, w: f.ctl.w, h: f.ctl.h }, moved: false };
  e.preventDefault(); e.stopPropagation();
}
function snapTargets(f) {
  const xs = [], ys = [];
  const sib = f.parent ? f.parent.children : screen().controls;
  const W = f.parent ? f.parent.w : ff().w, H = f.parent ? Number(f.parent.props.templateSize) : ff().h;
  xs.push(0, W / 2, W); ys.push(0, H / 2, H);
  for (const c of sib) if (!S.sel.includes(c.id)) { xs.push(c.x, c.x + c.w / 2, c.x + c.w); ys.push(c.y, c.y + c.h / 2, c.y + c.h); }
  return { xs, ys, ox: f.parent ? f.parent.x : 0, oy: f.parent ? f.parent.y : 0 };
}
function onMove(e) {
  if (!drag) return;
  const pt = screenPoint(e);
  const dx = pt.x - drag.start.x, dy = pt.y - drag.start.y;
  if (!drag.moved && Math.abs(dx) < 2 && Math.abs(dy) < 2) return;
  drag.moved = true;
  const grid = e.altKey ? 1 : 8, guides = [];
  if (drag.mode === 'marquee') {
    const x = Math.min(pt.x, drag.start.x), y = Math.min(pt.y, drag.start.y), w = Math.abs(dx), h = Math.abs(dy);
    const hit = screen().controls.filter((c) => c.x < x + w && c.x + c.w > x && c.y < y + h && c.y + c.h > y).map((c) => c.id);
    S.sel = [...new Set([...drag.base, ...hit])];
    drawOverlay();
    const m = document.createElement('div'); m.className = 'marquee'; m.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${h}px;border-width:${1 / zoomValue()}px`; $('#overlay').appendChild(m);
    return;
  }
  if (drag.mode === 'move') {
    const lead = drag.items[0];
    let nx = Math.round((lead.x + dx) / grid) * grid, ny = Math.round((lead.y + dy) / grid) * grid;
    if (!e.altKey && drag.items.length === 1) {
      const t = snapTargets(lead.f), c = lead.f.ctl, th = 6 / zoomValue();
      let best = null;
      for (const [off] of [[0], [c.w / 2], [c.w]]) for (const tx of t.xs) { const d = Math.abs(lead.x + dx + off - tx); if (d < th && (!best || d < best.d)) best = { d, v: tx - off, g: tx }; }
      if (best) { nx = Math.round(best.v); guides.push({ d: 'v', v: best.g + t.ox }); }
      best = null;
      for (const off of [0, c.h / 2, c.h]) for (const ty of t.ys) { const d = Math.abs(lead.y + dy + off - ty); if (d < th && (!best || d < best.d)) best = { d, v: ty - off, g: ty }; }
      if (best) { ny = Math.round(best.v); guides.push({ d: 'h', v: best.g + t.oy }); }
    }
    const ddx = nx - lead.x, ddy = ny - lead.y;
    for (const it of drag.items) { it.f.ctl.x = it.x + ddx; it.f.ctl.y = it.y + ddy; }
    moveDom(); drawOverlay(guides);
    // drop target highlight for reparenting into a gallery
    if (drag.items.length === 1 && !lead.f.parent && lead.f.ctl.type !== 'gallery' && lead.f.ctl.type !== 'person') {
      const c = lead.f.ctl, g = galleryAt({ x: c.x + c.w / 2, y: c.y + c.h / 2 }, c.id);
      if (g) { const d = document.createElement('div'); d.className = 'drop-hl'; d.style.cssText = `left:${g.x}px;top:${g.y}px;width:${g.w}px;height:${g.props.templateSize}px`; $('#overlay').appendChild(d); }
    }
    return;
  }
  if (drag.mode === 'resize') {
    const o = drag.o, c = drag.f.ctl, h = drag.h, sn = (v) => Math.round(v / grid) * grid;
    let { x, y, w, h: hh } = o;
    if (h.includes('e')) w = Math.max(4, sn(o.w + dx));
    if (h.includes('s')) hh = Math.max(4, sn(o.h + dy));
    if (h.includes('w')) { const nx = sn(o.x + dx); w = Math.max(4, o.w + (o.x - nx)); x = o.x + o.w - w; }
    if (h.includes('n')) { const ny = sn(o.y + dy); hh = Math.max(4, o.h + (o.y - ny)); y = o.y + o.h - hh; }
    Object.assign(c, { x, y, w, h: hh });
    if (c.type === 'gallery' && c.props.templateSize > hh) c.props.templateSize = hh;
    renderCanvasQuick(); drawOverlay();
  }
}
function moveDom() {
  for (const f of selCtls()) {
    const els = $$(`#screen [data-id="${f.ctl.id}"]`);
    els.forEach((el) => { el.style.left = f.ctl.x + 'px'; el.style.top = f.ctl.y + 'px'; });
  }
}
let quickRaf = 0;
function renderCanvasQuick() { cancelAnimationFrame(quickRaf); quickRaf = requestAnimationFrame(() => { const scr = screen(); renderScreen($('#screen'), scr, designEnv()); drawOverlay(); }); }
function onUp() {
  if (!drag) return;
  const d = drag; drag = null;
  if (d.mode === 'marquee') { drawOverlay(); renderInspector(); renderPanelIf('layers'); return; }
  if (!d.moved) return;
  if (d.mode === 'move' && d.items.length === 1) {
    const f = d.items[0].f, c = f.ctl, scr = screen();
    if (!f.parent && c.type !== 'gallery' && c.type !== 'person') {
      const g = galleryAt({ x: c.x + c.w / 2, y: c.y + c.h / 2 }, c.id);
      if (g) { scr.controls.splice(scr.controls.indexOf(c), 1); c.x -= g.x; c.y -= g.y; g.children.push(c); toast(`Added to ${g.name}: repeats on every row`); }
    } else if (f.parent) {
      const g = f.parent, cx = c.x + c.w / 2, cy = c.y + c.h / 2;
      if (cx < 0 || cy < 0 || cx > g.w || cy > g.h) { g.children.splice(g.children.indexOf(c), 1); c.x += g.x; c.y += g.y; scr.controls.push(c); toast('Moved out of the gallery'); }
    }
  }
  renderCanvas(); renderInspector(); renderPanelIf('layers'); commit();
}

// palette drag & drop
function insertControl(type, pt) {
  const scr = screen();
  const ctl = newControl(P, type);
  const f = ff();
  pt = pt || { x: Math.round(f.w / 2 - ctl.w / 2), y: Math.round(f.h / 3) };
  ctl.x = Math.max(0, Math.round((pt.x - ctl.w / 2) / 8) * 8); ctl.y = Math.max(0, Math.round((pt.y - ctl.h / 2) / 8) * 8);
  if (type === 'button' && !ctl.actions.OnSelect) ctl.actions.OnSelect = [];
  const g = type !== 'gallery' && type !== 'person' ? galleryAt(pt) : null;
  if (g) { ctl.x = Math.max(0, Math.min(g.w - ctl.w, Math.round(pt.x - g.x - ctl.w / 2))); ctl.y = Math.max(0, Math.min(g.props.templateSize - ctl.h, Math.round(pt.y - g.y - ctl.h / 2))); g.children.push(ctl); }
  else scr.controls.push(ctl);
  if (type === 'gallery' && P.lists[0]) { ctl.data.list = P.lists[0].name; ctl.data.searchField = P.lists[0].columns[0].name; galleryRows(P, scr, ctl, P.lists[0]); ctl.actions.OnSelect = [{ type: 'openItem', list: P.lists[0].name, screen: '' }]; }
  S.sel = [ctl.id];
  renderCanvas(); renderInspector(); renderPanelIf('layers'); commit();
}
function onDragOver(e) {
  if (!e.dataTransfer.types.includes('text/pas')) return;
  e.preventDefault(); e.dataTransfer.dropEffect = 'copy';
  const pt = screenPoint(e), g = galleryAt(pt);
  drawOverlay();
  if (g) { const d = document.createElement('div'); d.className = 'drop-hl'; d.style.cssText = `left:${g.x}px;top:${g.y}px;width:${g.w}px;height:${g.props.templateSize}px`; $('#overlay').appendChild(d); }
}
function onDrop(e) {
  const v = e.dataTransfer.getData('text/pas'); if (!v) return;
  e.preventDefault();
  const pt = screenPoint(e);
  if (v.startsWith('block:')) return doBlock(v.slice(6), pt);
  if (v.startsWith('myblock:')) return doMyBlock(v.slice(8), pt);
  insertControl(v, pt);
}
function doBlock(key, pt) {
  const b = BLOCKS[key];
  const list = P.lists.find((l) => l.name === S.blockList) || P.lists[0];
  if (b.needsList && !list) { toast('Add a list on the Data tab first', true); return; }
  const added = insertBlock(P, screen(), key, list, pt ? { x: Math.max(0, Math.round(pt.x / 8) * 8), y: Math.max(0, Math.round(pt.y / 8) * 8) } : { x: 40, y: 96 });
  S.sel = added.map((c) => c.id);
  renderCanvas(); renderInspector(); renderPanelIf('layers'); commit();
}

// ------------------------------------------------------------------ Play preview links (read-only, the project rides in the #fragment)
function shareDialog() {
  const m = modal(`<div class="modal-h"><h2>Share a Play preview</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b"><p style="margin-top:0">Anyone with this link can click through <b>${esc(P.app.name)}</b> with sample data in their browser. They cannot change your copy; they can open their own copy in the Studio. The app travels inside the link itself; nothing is uploaded.</p>
    <label class="chk-line"><input type="checkbox" id="shLinks"> Include the SharePoint site URL and app link</label>
    <p class="note">Never included: API keys, flow test samples and setup drafts. Group IDs and fixed email addresses are replaced, and imported sample rows are capped.</p>
    <div id="shOut" class="share-out"><p class="note">Building the link…</p></div></div>
    <div class="modal-f"><span class="grow" id="shSize"></span><button class="btn" data-close>Close</button><button class="btn pri" id="shCopy" disabled>${ui('copy')}Copy link</button></div>`, 'md');
  let url = '';
  const build = async () => {
    const r = await encodeShare(P, { includeLinks: m.el.querySelector('#shLinks').checked });
    const out = m.el.querySelector('#shOut'), btn = m.el.querySelector('#shCopy');
    const kb = (r.chars / 1024).toFixed(1);
    if (r.tooLong) {
      url = ''; btn.disabled = true;
      out.innerHTML = `<p class="err-box">This app is too big for a link (${kb} KB; the limit is about ${Math.round(MAX_CHARS / 1000)} KB). Large embedded images are the usual cause: remove them or add them as app media, or send the project file instead.</p>`;
    } else {
      url = location.origin + location.pathname + r.hash; btn.disabled = false;
      out.innerHTML = `<textarea class="inp" rows="4" readonly id="shUrl" aria-label="Preview link">${esc(url)}</textarea>${r.warn ? `<p class="warn-box">This link is long (${kb} KB). Some mail and chat apps cut links this long; if it does not open, send the project file instead.</p>` : ''}`;
    }
    m.el.querySelector('#shSize').textContent = `${kb} KB${r.compressed ? ', compressed' : ''}`;
  };
  m.el.querySelector('#shLinks').onchange = build;
  m.el.querySelector('#shCopy').onclick = () => url && copyText(url);
  build().catch((e) => { m.el.querySelector('#shOut').innerHTML = `<p class="err-box">${esc(e.message)}</p>`; });
}
/** Opened from a #play=… link: Play only, nothing saved over the visitor's own project. */
async function openShared(hash) {
  S.viewer = true; document.body.classList.add('viewer');
  let p;
  try { p = await decodeShare(hash); } catch (e) {
    const m = modal(`<div class="modal-h"><h2>Cannot open this preview</h2></div><div class="modal-b"><p style="margin:0">${esc(e.message)}</p></div><div class="modal-f"><button class="btn pri" data-x>Open the Studio</button></div>`, 'sm');
    m.el.querySelector('[data-x]').onclick = () => { m.close(); leaveViewer(); };
    return;
  }
  P = p; migrate(P); S.screenId = P.screens[0]?.id; S.last = ser(P); renderAll();
  play({ viewer: true });
}
function leaveViewer(copy = null) {
  S.viewer = false; document.body.classList.remove('viewer');
  try { history.replaceState(null, '', location.pathname + location.search); } catch { /* sandboxed */ }
  if (copy) { copy.app.name = (copy.app.name || 'App') + ' (copy)'; VS.ready = true; VS.sig = null; setProject(copy); toast('Opened a copy. It is saved in this browser.'); return; }
  const saved = loadSaved(); setProject(saved || core.newProject()); attachVersions(!!saved); if (!saved) wizard(true);
}

// ------------------------------------------------------------------ my blocks (saved in this browser; export/import as .json)
function myBlocksHtml() {
  const bl = loadBlocks();
  return `<div class="sub">My blocks <small class="muted">${bl.length ? 'saved in this browser' : ''}</small></div>
    ${bl.length ? bl.map((b) => `<div class="myblock"><button class="rowbtn" draggable="true" data-myblock="${esc(b.id)}" title="${esc(`${b.controls.length} control${b.controls.length > 1 ? 's' : ''}, ${b.w}×${b.h}${b.lists.length ? ' · uses ' + b.lists.join(', ') : ''}`)}"><i>${ui('layers')}</i><span>${esc(b.name)}</span></button><button class="mini" data-mbcmd="rename" data-id="${esc(b.id)}" title="Rename" aria-label="Rename ${esc(b.name)}">${ui('edit')}</button><button class="mini danger" data-mbcmd="del" data-id="${esc(b.id)}" title="Delete" aria-label="Delete ${esc(b.name)}">${ui('trash')}</button></div>`).join('')
    : '<p class="note">Select controls on a screen and choose <b>Save as block</b> to reuse them on any screen or app.</p>'}
    <div class="btns"><button class="btn sm" data-mbcmd="import">${ui('upload')}Import</button>${bl.length ? `<button class="btn sm" data-mbcmd="export">${ui('download')}Export all</button>` : ''}</div>`;
}
async function saveSelAsBlock() {
  const items = selCtls().filter((f) => !f.parent);
  if (!items.length) { toast('Select controls on the screen (not inside a gallery row) to save them as a block', true); return; }
  const name = await askText('Save as block', 'Block name', items.length === 1 ? items[0].ctl.name.replace(/[A-Z][a-z]+$/, '') : 'My block');
  if (name == null) return;
  const b = makeBlock(items.map((f) => f.ctl), name);
  const all = loadBlocks(); all.unshift(b);
  if (!saveBlocks(all)) { toast('This browser would not keep the block (private window or storage full). Use Export to keep a copy.', true); return; }
  toast(`Saved "${b.name}" to My blocks`); renderPanelIf('insert');
}
function doMyBlock(id, pt) {
  const b = loadBlocks().find((x) => x.id === id); if (!b) return;
  const at = pt ? { x: Math.max(0, Math.round(pt.x / 8) * 8), y: Math.max(0, Math.round(pt.y / 8) * 8) } : { x: 40, y: 96 };
  const r = insertSavedBlock(P, screen(), b, at);
  S.sel = r.added.map((c) => c.id);
  renderCanvas(); renderInspector(); renderPanelIf('layers'); commit();
  r.warnings.forEach((w) => toast(w, true));
}
async function myBlockCmd(cmd, id) {
  const all = loadBlocks(), b = all.find((x) => x.id === id);
  if (cmd === 'rename' && b) { const n = await askText('Rename block', 'Block name', b.name); if (n == null || !n.trim()) return; b.name = n.trim().slice(0, 60); saveBlocks(all); }
  if (cmd === 'del' && b) { if (!(await ask('Delete block?', `"${b.name}" is removed from My blocks in this browser. Screens that use it keep their controls.`, 'Delete'))) return; saveBlocks(all.filter((x) => x !== b)); }
  if (cmd === 'export') download('my-blocks.json', exportBlocksJson(all), 'application/json');
  if (cmd === 'import') { const fi = $('#fileBlocks'); fi.value = ''; fi.click(); return; }
  renderPanelIf('insert');
}
async function importBlocksFile(f) {
  try {
    const { blocks, skipped } = importBlocksJson(await f.text());
    const all = loadBlocks();
    for (const b of blocks) { let n = b.name, i = 2; while (all.some((x) => x.name === n)) n = `${b.name} ${i++}`; b.name = n; all.push(b); }
    if (!saveBlocks(all)) { toast('This browser would not keep the blocks (private window or storage full).', true); return; }
    toast(`Imported ${blocks.length} block${blocks.length === 1 ? '' : 's'}${skipped ? `; ${skipped} could not be read` : ''}`);
    renderPanelIf('insert');
  } catch (err) { toast(err.message, true); }
}

// ------------------------------------------------------------------ edit ops
function deleteSel() {
  const items = selCtls(); if (!items.length) return;
  for (const f of items) { const arr = f.parent ? f.parent.children : f.screen.controls; const i = arr.indexOf(f.ctl); if (i >= 0) arr.splice(i, 1); }
  S.sel = []; renderCanvas(); renderInspector(); renderPanelIf('layers'); commit();
}
function pasteCtls(list, offset = 16) {
  const scr = screen(), ids = [], used = allNames(P);
  for (const raw of list) {
    const c = renameCopy(P, clone(raw), used); c.x += offset; c.y += offset;
    if (raw.__parent) { const g = findCtl(P, raw.__parent)?.ctl; delete c.__parent; if (g?.children && screen().controls.includes(g)) { g.children.push(c); ids.push(c.id); continue; } }
    delete c.__parent; scr.controls.push(c); ids.push(c.id);
  }
  S.sel = ids; renderCanvas(); renderInspector(); renderPanelIf('layers'); commit();
}
const copySel = () => { S.clip = selCtls().map((f) => ({ ...clone(f.ctl), __parent: f.parent?.id })); if (S.clip.length) toast(`Copied ${S.clip.length} control${S.clip.length > 1 ? 's' : ''}`); };
function zMove(dir) {
  for (const f of selCtls()) {
    const arr = f.parent ? f.parent.children : f.screen.controls, i = arr.indexOf(f.ctl);
    arr.splice(i, 1);
    const j = dir === 'front' ? arr.length : dir === 'back' ? 0 : dir === 'up' ? Math.min(arr.length, i + 1) : Math.max(0, i - 1);
    arr.splice(j, 0, f.ctl);
  }
  renderCanvas(); renderPanelIf('layers'); commit();
}
function nudge(dx, dy) { for (const f of selCtls()) { f.ctl.x += dx; f.ctl.y += dy; } renderCanvas(); renderInspector(); commitSoon(); }
function alignSel(how) {
  const items = selCtls(); if (items.length < 2) return;
  const xs = items.map((f) => f.ctl), L = Math.min(...xs.map((c) => c.x)), R = Math.max(...xs.map((c) => c.x + c.w)), T = Math.min(...xs.map((c) => c.y)), B = Math.max(...xs.map((c) => c.y + c.h));
  for (const c of xs) {
    if (how === 'left') c.x = L; if (how === 'right') c.x = R - c.w; if (how === 'center') c.x = Math.round((L + R) / 2 - c.w / 2);
    if (how === 'top') c.y = T; if (how === 'bottom') c.y = B - c.h; if (how === 'middle') c.y = Math.round((T + B) / 2 - c.h / 2);
  }
  if (how === 'hdist' || how === 'vdist') {
    const k = how === 'hdist' ? 'x' : 'y', s = how === 'hdist' ? 'w' : 'h';
    const sorted = [...xs].sort((a, b) => a[k] - b[k]), total = sorted.reduce((a, c) => a + c[s], 0);
    const span = (how === 'hdist' ? R - L : B - T) - total, gap = span / (sorted.length - 1);
    let pos = how === 'hdist' ? L : T; for (const c of sorted) { c[k] = Math.round(pos); pos += c[s] + gap; }
  }
  renderCanvas(); renderInspector(); commit();
}

// ------------------------------------------------------------------ left panel
function renderPanelIf(tab) { if (S.tab === tab) renderPanel(); }
function renderPanel() {
  const el = $('#panel');
  el.classList.toggle('wide', S.tab === 'data');
  const f = { insert: panelInsert, screens: panelScreens, layers: panelLayers, data: panelData, flows: panelFlows, theme: panelTheme, app: panelApp }[S.tab];
  el.innerHTML = f();
  if (S.tab === 'data') bindDataEditor(el, P.lists, () => { renderCanvas(); renderInspector(); commitSoon(); });
}
function listOptions(sel, withBlank = true, blankText = '(choose a list)') {
  return (withBlank ? `<option value="">${blankText}</option>` : '') + P.lists.map((l) => `<option ${l.name === sel ? 'selected' : ''}>${esc(l.name)}</option>`).join('');
}
function panelInsert() {
  const cats = CATS.map((cat) => {
    const items = Object.entries(CATALOG).filter(([, d]) => d.cat === cat);
    return `<div class="sub">${cat}</div><div class="palette">${items.map(([k, d]) => `<button class="pal" draggable="true" data-ins="${k}" title="Drag onto the screen, or click to add">${`<i>${ui(k)}</i>`}${esc(d.label)}</button>`).join('')}</div>`;
  }).join('');
  if (!S.blockList && P.lists[0]) S.blockList = P.lists[0].name;
  const blocks = Object.entries(BLOCKS).map(([k, b]) => `<button class="rowbtn" draggable="true" data-block="${k}"><i>${ui(b.icon)}</i><span>${esc(b.label)}</span></button>`).join('');
  return `<div class="ph"><h2>Insert</h2></div><div class="pb">${cats}
    <div class="sub">Ready-made blocks</div>
    <label class="f"><span>Blocks use this list</span><select data-pcmd="blockList">${listOptions(S.blockList, !P.lists.length, '(add a list on the Data tab)')}</select></label>
    ${blocks}
    ${myBlocksHtml()}
    <div class="sub">Screens</div>
    <button class="btn block" data-cmd="newScreen">${ui('plus')}New screen from a template</button>
  </div>`;
}
function panelScreens() {
  const start = P.app.startScreen || P.screens[0]?.id;
  return `<div class="ph"><h2>Screens</h2><button class="btn sm" data-cmd="newScreen">${ui('plus')}Add</button></div><div class="pb">
    ${P.screens.map((s, i) => `<div class="scrrow ${s.id === screen().id ? 'on' : ''}" data-scr="${s.id}">
      <span class="nm">${esc(s.name)}</span>${s.id === start ? '<span class="start">START</span>' : ''}
      <button class="mini" data-cmd="scrUp" data-i="${i}" title="Move up">${ui('up')}</button>
      <button class="mini" data-cmd="scrDup" data-i="${i}" title="Duplicate">${ui('copy')}</button>
      <button class="mini danger" data-cmd="scrDel" data-i="${i}" title="Delete">${ui('trash')}</button></div>`).join('')}
    <p class="note">Each screen becomes one YAML block you paste into Power Apps Studio. Select a screen tab, then click an empty spot on the canvas to edit the screen's own settings.</p></div>`;
}
function panelLayers() {
  const scr = screen();
  const row = (c, depth) => `<div class="layer ${S.sel.includes(c.id) ? 'on' : ''}" data-sel="${c.id}" style="padding-left:${6 + depth * 16}px"><i>${ui(c.type)}</i><span class="nm">${esc(c.name)}</span>
    <span class="acts"><button class="mini" data-lcmd="up" data-id="${c.id}" title="Bring forward">${ui('up')}</button><button class="mini" data-lcmd="down" data-id="${c.id}" title="Send backward">${ui('down')}</button></span></div>` +
    (c.children ? [...c.children].reverse().map((ch) => row(ch, depth + 1)).join('') : '');
  return `<div class="ph"><h2>Layers</h2></div><div class="pb">
    <p class="note">Top of the list is in front. In Power Apps, later controls draw on top and receive clicks first.</p>
    ${[...scr.controls].reverse().map((c) => row(c, 0)).join('') || '<p class="note">This screen is empty.</p>'}</div>`;
}
function panelData() {
  return `<div class="ph"><h2>Data</h2><button class="btn sm" data-cmd="addList">${ui('plus')}List</button></div><div class="pb">
    <p class="note">Design the SharePoint lists your app uses. Play mode fills them with sample rows. Get YAML → Provisioning script creates them on your site.</p>
    ${dataEditorHtml(P.lists)}
    <div class="btns"><button class="btn" data-cmd="importTable">${ui('upload')}Import CSV or Excel</button><button class="btn" data-cmd="importSchema">${ui('upload')}Import schema</button><button class="btn" data-cmd="showScript">${ui('file')}Provisioning script</button></div></div>`;
}
function dataEditorHtml(lists) {
  if (!lists.length) return '<p class="note"><b>No lists yet.</b> Add one, or run New app to start from a scenario.</p>';
  return lists.map((l, li) => `<div class="listcard" data-li="${li}">
    <div class="listcard-h"><i style="width:16px;height:16px;color:var(--cyan)">${ui('data')}</i><input value="${esc(l.name)}" data-dl="name" aria-label="List name"><button class="mini danger" data-dcmd="delList" title="Delete list">${ui('trash')}</button></div>
    <div class="lsize"><label for="lsz${li}">Expected size</label><select class="inp" id="lsz${li}" data-dl="size" title="${esc(SIZES[l.size]?.note || 'Not set: formulas are generated for lists under 500 items.')}"><option value="">Not set</option>${Object.entries(SIZES).map(([k, v]) => `<option value="${k}" ${k === l.size ? 'selected' : ''}>${v.label}</option>`).join('')}</select></div>
    <label class="lattach" title="Each item can carry files. Attachments cannot be pasted: the screens get a dashed box where you insert one Edit form by hand (the Paste assistant has the exact steps)."><input type="checkbox" data-dl="attachments" ${l.attachments ? 'checked' : ''}>Attachments <small>(form added by hand after the paste)</small></label>
    ${l.columns.map((c, ci) => `<div class="col" data-ci="${ci}">
      <div class="col-top"><input class="inp" value="${esc(c.name)}" data-dc="name" ${c.builtin ? 'disabled' : ''} aria-label="Column name" title="${esc(c.name)}">
        ${c.builtin ? '' : `<button class="mini" data-dcmd="colUp" title="Move up" aria-label="Move ${esc(c.name)} up">${ui('up')}</button><button class="mini danger" data-dcmd="delCol" title="Delete column" aria-label="Delete ${esc(c.name)}">${ui('trash')}</button>`}</div>
      <div class="col-more"><select class="inp col-type" data-dc="type" ${c.builtin ? 'disabled' : ''} aria-label="Type of ${esc(c.name)}">${Object.entries(COLTYPES).map(([k, v]) => `<option value="${k}" ${k === c.type ? 'selected' : ''}>${v.label}</option>`).join('')}</select>
        <label class="chk" style="margin:0"><input type="checkbox" data-dc="required" ${c.required ? 'checked' : ''} ${c.type === 'YesNo' ? 'disabled' : ''}>Required</label>${c.builtin ? '<span class="lock">Built-in Title column</span>' : ''}
        ${c.type === 'Choice' ? `<textarea class="inp" data-dc="choices" placeholder="One choice per line">${esc(c.choices.join('\n'))}</textarea>` : ''}
        ${c.type === 'Lookup' ? `<select class="inp" data-dc="lookupList"><option value="">(lookup target list)</option>${lists.filter((x) => x !== l).map((x) => `<option ${x.name === c.lookupList ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select>` : ''}</div>
    </div>`).join('')}
    ${l.rows?.length ? `<div class="rowsbar"><span><b>${l.rows.length}</b> imported row${l.rows.length === 1 ? '' : 's'} (Play uses them)</span><button class="btn sm" data-dcmd="exportRows">${ui('download')}Export CSV</button><button class="btn sm" data-dcmd="clearRows">Clear</button></div>` : ''}
    <div style="padding:8px 10px"><button class="btn sm" data-dcmd="addCol">${ui('plus')}Column</button></div></div>`).join('');
}
function bindDataEditor(root, lists, onChange, rerender) {
  const redraw = rerender || (() => { renderPanel(); });
  root.oninput = root.onchange = (e) => {
    const t = e.target, card = t.closest('[data-li]'); if (!card) return;
    const l = lists[+card.dataset.li];
    if (t.dataset.dl === 'size') {
      if (e.type !== 'change') return;
      if (t.value) l.size = t.value; else delete l.size;
      t.title = SIZES[l.size]?.note || '';
      onChange();
      if (lists === P.lists && l.size) toast(SIZES[l.size].note);
      return;
    }
    if (t.dataset.dl === 'attachments') {
      if (e.type !== 'change') return;
      if (t.checked) l.attachments = true; else delete l.attachments;
      if (lists !== P.lists) return;
      (async () => {
        const tight = t.checked ? attachNoRoom(P, l.name) : [];
        const rebuild = tight.length ? await ask('Make room for attachments', `${tight.join(' and ')} ${tight.length > 1 ? 'have' : 'has'} no free space for the attachments box. Rebuild ${tight.length > 1 ? 'them' : 'it'} from the data model (same names, so links keep working)? Changes you made by hand on ${tight.length > 1 ? 'those screens' : 'that screen'} are replaced. Cancel puts the box below the form for you to move.`, 'Rebuild') : false;
        const r = syncAttachSlots(P, l.name, { rebuild });
        onChange();
        const n = r.added.length + r.rebuilt.length;
        if (t.checked) toast(n ? `Attachments box on ${[...r.added, ...r.rebuilt].join(', ')}. After pasting, follow the Paste assistant's Attachments step.` : `Attachments on for ${l.name}. Its edit and details screens get the box when you add them.`);
        else toast(r.removed.length ? `Attachments box removed from ${r.removed.join(', ')}. Delete the hand-added form in Studio too.` : `Attachments off for ${l.name}.`);
      })();
      return;
    }
    if (t.dataset.dl === 'name') {
      const nv = t.value.trim();
      if (e.type === 'change' && nv && nv !== l.name) { if (lists === P.lists) renameList(l.name, nv); else l.name = nv; onChange(); if (lists === P.lists) renderPanel(); }
      return;
    }
    const colEl = t.closest('[data-ci]'); if (!colEl || !t.dataset.dc) return;
    const c = l.columns[+colEl.dataset.ci], k = t.dataset.dc;
    if (k === 'name') { if (e.type !== 'change') return; const nv = t.value.trim().replace(/\s+/g, ''); if (!nv || nv === c.name) return; if (lists === P.lists) renameColumn(l.name, c.name, nv); else c.name = nv; onChange(); redraw(); return; }
    if (k === 'type') { c.type = t.value; if (c.type === 'Choice' && !c.choices.length) c.choices = ['Option A', 'Option B']; if (c.type === 'YesNo') c.required = false; onChange(); redraw(); return; }
    if (k === 'required') c.required = t.checked;
    if (k === 'choices') c.choices = t.value.split('\n').map((x) => x.trim()).filter(Boolean);
    if (k === 'lookupList') c.lookupList = t.value;
    onChange();
  };
  root.onclick = (e) => {
    const b = e.target.closest('[data-dcmd]'); if (!b) return;
    const card = b.closest('[data-li]'), l = lists[+card.dataset.li], colEl = b.closest('[data-ci]'), ci = colEl ? +colEl.dataset.ci : -1;
    const cmd = b.dataset.dcmd;
    if (cmd === 'addCol') { let n = 'NewColumn', i = 2; while (l.columns.some((c) => c.name === n)) n = 'NewColumn' + i++; l.columns.push(newColumn(n, 'Text')); }
    if (cmd === 'delCol') l.columns.splice(ci, 1);
    if (cmd === 'exportRows') { download(`${core.slug(l.name) || 'list'}.csv`, listToCsv(l), 'text/csv'); return; }
    if (cmd === 'clearRows') delete l.rows;
    if (cmd === 'colUp' && ci > 1) l.columns.splice(ci - 1, 0, l.columns.splice(ci, 1)[0]);
    if (cmd === 'delList') {
      if (lists === P.lists && usedLists(P).includes(l.name)) { ask('Delete list?', `"${l.name}" is used by controls in your screens. They will show errors until you rebind them.`, 'Delete').then((ok) => { if (ok) { lists.splice(lists.indexOf(l), 1); onChange(); redraw(); } }); return; }
      lists.splice(lists.indexOf(l), 1);
    }
    onChange(); redraw();
  };
}
// renames also rewrite the formulas the generator baked into screens (HtmlText, variables, filters)
function renameList(oldN, newN) {
  const err = renameListIn(P, oldN, newN);
  if (err) { if (err !== 'Same name.') toast(err, true); return; }
  if (S.blockList === oldN) S.blockList = newN;
}
function renameColumn(list, oldN, newN) {
  const err = renameColumnIn(P, list, oldN, newN);
  if (err && err !== 'Same name.') toast(err, true);
}
/** WCAG AA check for the text/background pairs the generated screens use. */
function refreshContrast() { const b = $('#cxBox'); if (b) b.innerHTML = contrastHtml(P.theme); }
function contrastHtml(t) {
  const rows = CONTRAST_PAIRS.map(([f, b, where]) => {
    const r = contrast(t.colors[f], t.colors[b]), ok = r >= AA;
    return `<div class="cx ${ok ? 'ok' : 'bad'}"><span class="cx-sw" style="background:${esc(t.colors[b])};color:${esc(t.colors[f])}">Aa</span>
      <span class="cx-t"><b>${f} on ${b}</b><small>${esc(where)}</small></span><span class="cx-r">${r.toFixed(2)}:1</span>
      ${ok ? '<span class="cx-ok" title="Passes WCAG AA (4.5:1)">AA</span>' : `<button class="btn sm" data-cfix="${f}|${b}" title="Nudge the color just enough to reach 4.5:1">Fix</button>`}</div>`;
  }).join('');
  return `<div class="sub">Contrast <small>(WCAG AA, 4.5:1 for text)</small></div><div class="cxs">${rows}</div>`;
}
function panelTheme() {
  const t = P.theme;
  return `<div class="ph"><h2>Theme</h2></div><div class="pb">
    <div class="sub">Presets</div>
    <div class="presets">${THEME_PRESETS.map((p, i) => `<button class="preset ${p.name === t.name ? 'on' : ''}" data-preset="${i}"><div class="bar">${['HeaderBg', 'Primary', 'Accent', 'Bg', 'Surface'].map((k) => `<i style="background:${p.colors[k]}"></i>`).join('')}</div><span>${esc(p.name)}</span></button>`).join('')}</div>
    <div class="sub">Type and shape</div>
    <label class="f"><span>Font</span><select data-th="font">${FONTS.map((f) => `<option ${f === t.font ? 'selected' : ''}>${f}</option>`).join('')}</select></label>
    <label class="f"><span>Corner radius: <b id="radVal">${t.radius}px</b></span><input type="range" min="0" max="24" value="${t.radius}" data-th="radius"></label>
    <label class="f"><span>Header style for new screens</span><select data-th="header"><option value="gradient" ${t.header === 'gradient' ? 'selected' : ''}>Gradient (HeaderBg → Primary)</option><option value="solid" ${t.header === 'solid' ? 'selected' : ''}>Solid</option></select></label>
    <label class="chk"><input type="checkbox" data-th="dark" ${t.dark ? 'checked' : ''}>Dark theme (hover lightens instead of darkens)</label>
    <div id="cxBox">${contrastHtml(t)}</div>
    <div class="sub">Colors</div>
    ${THEME_COLORS.map(([k, d]) => `<div class="tokrow"><input type="color" value="${(t.colors[k] || '#000000').slice(0, 7)}" data-tc="${k}" aria-label="${k}"><span class="tn">${k}<small>${d}</small></span><input class="hex" value="${esc(t.colors[k])}" data-tch="${k}" aria-label="${k} hex"></div>`).join('')}
    <p class="note">Every control reads these from one <code>appTheme</code> record in App.Formulas, so you can also restyle the app later in Studio by editing that record.</p></div>`;
}
function panelApp() {
  const start = P.app.startScreen || P.screens[0]?.id;
  return `<div class="ph"><h2>App</h2></div><div class="pb">
    <label class="f"><span>App name</span><input data-app="name" value="${esc(P.app.name)}"></label>
    <label class="f"><span>Format</span><select data-app="formFactor">${Object.entries(FORM_FACTORS).map(([k, v]) => `<option value="${k}" ${k === P.app.formFactor ? 'selected' : ''}>${v.label}</option>`).join('')}</select></label>
    <label class="f"><span>Start screen</span><select data-app="startScreen">${P.screens.map((s) => `<option value="${s.id}" ${s.id === start ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select></label>
    <div class="sub">Email links</div>
    <label class="f"><span>App web link <small>(Apps → Details → Web link)</small></span><input data-app="appLink" value="${esc(P.app.appLink)}" placeholder="https://apps.powerapps.com/play/e/…?tenantId=…"></label>
    <div class="grid2"><label class="f"><span>Open links on</span><select data-app="deepLinkScreen">${sel([['', '(start screen)'], ...P.screens.map((s) => [s.id, s.name])], P.app.deepLink?.screen)}</select></label>
    <label class="f"><span>Item from list</span><select data-app="deepLinkList">${sel([['', '(list)'], ...P.lists.map((l) => l.name)], P.app.deepLink?.list)}</select></label></div>
    <p class="note">A link ending in <code>&amp;ID=12</code> opens this screen with item 12 selected. Flow emails use it for {Link}.</p>
    <label class="f"><span>SharePoint site URL <small>(for flows)</small></span><input data-app="siteUrl" value="${esc(P.app.siteUrl)}" placeholder="https://contoso.sharepoint.com/sites/Team"></label>
    <div class="sub">Advanced</div>
    <label class="f"><span>Extra App.OnStart (Power Fx)</span><textarea data-app="onStart" rows="4" placeholder="Set(varIsAdmin, false)">${esc(P.app.onStart)}</textarea></label>
    <label class="f"><span>Extra App.Formulas (named formulas)</span><textarea data-app="formulas" rows="4" placeholder="AppVersion = &quot;1.0&quot;;">${esc(P.app.formulas)}</textarea></label>
    <label class="chk"><input type="checkbox" data-appft="a11y" ${P.app.features?.a11y ? 'checked' : ''}>Add accessibility labels (verify in your tenant)</label>
    <p class="note" style="margin-top:-4px">Adds <code>AccessibleLabel</code> to inputs and to the invisible tap buttons over icons, images and cards, so screen readers announce them. Not yet field-tested in a GCC-High paste: paste one screen first and check it.</p>
    <div class="sub">Build</div>
    <button class="rowbtn" data-cmd="describe"><i>${ui('edit')}</i><span>Describe your app<small>Write the requirements; Claude plans the lists, screens and flows</small></span></button>
    <button class="rowbtn" data-cmd="wizard"><i>${ui('star')}</i><span>New app from requirements<small>Pick a scenario, shape the lists, generate every screen</small></span></button>
    <button class="rowbtn" data-cmd="import"><i>${ui('upload')}</i><span>Import screen YAML<small>Bring a screen from Power Apps Studio back in to edit</small></span></button>
    <button class="rowbtn" data-cmd="code"><i>${ui('html')}</i><span>Get YAML<small>App.Formulas, OnStart and every screen, ready to paste</small></span></button>
  </div>`;
}

// ------------------------------------------------------------------ inspector
const sel = (opts, v) => opts.map((o) => { const [val, label] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(val)}" ${String(val) === String(v ?? '') ? 'selected' : ''}>${esc(label)}</option>`; }).join('');
const F = (label, inner, cls = '') => `<label class="f ${cls}"><span>${label}</span>${inner}</label>`;
const inp = (path, v, t = 'str', extra = '') => `<input ${t === 'num' ? 'type="number"' : 'type="text"'} data-p="${path}" data-t="${t}" value="${esc(v ?? '')}" ${extra}>`;
const ta = (path, v, cls = 'plain', rows = 2, ph = '') => `<textarea class="${cls}" rows="${rows}" data-p="${path}" placeholder="${esc(ph)}">${esc(v ?? '')}</textarea>`;
const pick = (path, opts, v, re = false) => `<select data-p="${path}" ${re ? 'data-re="1"' : ''}>${sel(opts, v)}</select>`;
const chk = (path, v, label, re = false) => `<label class="chk"><input type="checkbox" data-p="${path}" data-t="bool" ${v ? 'checked' : ''} ${re ? 'data-re="1"' : ''}>${label}</label>`;
function colorField(path, v, label) {
  const css = cssColor(v, P.theme);
  const name = !v || v === 'transparent' ? 'Transparent' : v[0] === '@' ? 'Theme · ' + v.slice(1) : v;
  return F(label, `<div class="colorf"><button type="button" class="sw" data-color="${path}" aria-label="Pick color"><b style="background:${css}"></b></button><input type="text" data-p="${path}" value="${esc(v)}" title="${esc(name)}" placeholder="@Primary, #RRGGBB or transparent"></div>`);
}
const ALIGNS = ['Left', 'Center', 'Right'], WEIGHTS = ['Lighter', 'Normal', 'Semibold', 'Bold'];
function fieldOpts(listName, types) {
  const l = P.lists.find((x) => x.name === listName);
  return l ? l.columns.filter((c) => !types || types.includes(c.type)).map((c) => c.name) : [];
}
const INPUT_COLTYPES = { textinput: ['Text', 'Note', 'Number', 'Currency'], dropdown: ['Choice', 'Lookup'], datepicker: ['DateTime'], checkbox: ['YesNo'], toggle: ['YesNo'], person: ['Person'] };

/** Re-render the inspector, keeping focus (and the caret) on the same field if it still exists. */
function renderInspector() {
  const el = $('#inspector');
  const a = document.activeElement;
  const keep = a && el.contains(a) ? { sel: ['data-p', 'data-o', 'data-ev', 'data-i', 'data-bindkind', 'data-vis'].filter((k) => a.hasAttribute(k)).map((k) => `[${k}="${CSS.escape(a.getAttribute(k))}"]`).join(''), s: a.selectionStart, e: a.selectionEnd, top: el.scrollTop } : null;
  const items = selCtls();
  if (items.length > 1) el.innerHTML = inspMulti(items);
  else if (items.length === 1) el.innerHTML = inspControl(items[0]);
  else el.innerHTML = inspScreen(screen());
  if (keep) {
    el.scrollTop = keep.top;
    const n = keep.sel && el.querySelector(keep.sel);
    if (n) { n.focus({ preventScroll: true }); try { if (keep.s != null) n.setSelectionRange(keep.s, keep.e); } catch { /* not a text field */ } }
  }
}
function inspMulti(items) {
  return `<div class="insp-head"><i>${ui('layers')}</i><div class="t"><b>Selection</b><span style="font-weight:700">${items.length} controls</span></div></div>
  <div class="sec"><h3>Align</h3><div class="btns">${[['left', 'Left'], ['center', 'Center'], ['right', 'Right'], ['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom'], ['hdist', 'Space across'], ['vdist', 'Space down']].map(([k, l]) => `<button class="btn sm" data-align="${k}">${l}</button>`).join('')}</div></div>
  <div class="sec"><h3>Arrange</h3><div class="btns"><button class="btn sm" data-cmd="z" data-z="front">To front</button><button class="btn sm" data-cmd="z" data-z="back">To back</button><button class="btn sm" data-cmd="dup">Duplicate</button><button class="btn sm" data-cmd="saveBlock">Save as block</button><button class="btn sm danger" data-cmd="del">Delete</button></div></div>`;
}
function inspScreen(scr) {
  const start = (P.app.startScreen || P.screens[0]?.id) === scr.id;
  return `<div class="insp-head"><i>${ui('screen')}</i><div class="t"><b>Screen</b><input data-o="screen" data-p="name" value="${esc(scr.name)}" aria-label="Screen name"></div></div>
  <div class="sec"><h3>Screen</h3>
    ${colorField('fill', scr.fill || '@Bg', 'Background')}
    ${start ? '<p class="note">★ This is the start screen.</p>' : '<button class="btn sm" data-cmd="makeStart">Make start screen</button>'}
  </div>
  <div class="sec"><h3>OnVisible <small>Power Fx</small></h3><label class="f">${ta('onVisible', scr.onVisible, '', 3, 'Runs each time the screen opens. Inputs bound to a list are reset automatically.').replace('data-p=', 'data-o="screen" data-p=')}</label>
    <p class="note">Navigate() is not allowed here. Hide the button that leads to a screen instead of bouncing users out.</p></div>
  <div class="sec"><h3>Screen YAML preview</h3><pre class="code small">${esc(screenYaml(P, scr).split('\n').slice(0, 40).join('\n'))}${screenYaml(P, scr).split('\n').length > 40 ? '\n…' : ''}</pre>
    <div class="btns" style="margin-top:8px"><button class="btn sm" data-cmd="copyScreen">${ui('copy')}Copy this screen's YAML</button></div></div>
  <div class="empty-insp">${ui('plus')}<p>Select a control to edit it, or drag one in from Insert.</p></div>`.replace(/data-p="fill"/g, 'data-o="screen" data-p="fill"').replace(/data-color="fill"/, 'data-color="screen:fill"');
}
function inspControl(f) {
  const c = f.ctl, p = c.props, d = CATALOG[c.type];
  const parts = [];
  parts.push(`<div class="insp-head"><i>${ui(c.type)}</i><div class="t"><b>${esc(d.label)}${f.parent ? ' · in ' + esc(f.parent.name) : ''}</b><input data-p="name" data-t="name" value="${esc(c.name)}" spellcheck="false" aria-label="Control name"></div>
    <button class="mini" data-cmd="dup" title="Duplicate (Ctrl+D)">${ui('copy')}</button>${f.parent ? '' : `<button class="mini" data-cmd="saveBlock" title="Save as a reusable block (My blocks)" aria-label="Save as block">${ui('download')}</button>`}<button class="mini danger" data-cmd="del" title="Delete">${ui('trash')}</button></div>`);
  parts.push(`<div class="sec"><h3>Position and size${f.parent ? ' <small>relative to the row</small>' : ''}</h3><div class="grid4">${F('X', inp('x', c.x, 'num'))}${F('Y', inp('y', c.y, 'num'))}${F('W', inp('w', c.w, 'num'))}${F('H', inp('h', c.h, 'num'))}</div>
    <div class="btns"><button class="btn sm" data-cmd="z" data-z="front">To front</button><button class="btn sm" data-cmd="z" data-z="back">To back</button>${!f.parent ? `<button class="btn sm" data-cmd="fullw">Full width</button>` : ''}</div></div>`);
  // content
  const content = [];
  const textish = ['label', 'button', 'badge'].includes(c.type);
  if (textish) {
    content.push(dataSourceForText(f));
    if (!c.bind && !c.fx.Text) content.push(F('Text', ta('props.text', p.text, 'plain', c.type === 'label' ? 2 : 1)));
    content.push(`<div class="grid2">${F('Size', inp('props.size', p.size, 'num'))}${F('Weight', pick('props.weight', WEIGHTS, p.weight))}</div>`);
    if (c.type === 'label') content.push(`<div class="grid2">${F('Align', pick('props.align', ALIGNS, p.align))}${F('Vertical', pick('props.valign', ['Top', 'Middle', 'Bottom'], p.valign))}</div>`, chk('props.italic', p.italic, 'Italic'));
    content.push(colorField('props.color', p.color, 'Text color'), colorField('props.fill', p.fill, 'Fill'));
    if (c.type === 'button') content.push(colorField('props.borderColor', p.borderColor, 'Border color'), `<div class="grid2">${F('Border', inp('props.borderThickness', p.borderThickness, 'num'))}${F('Radius', inp('props.radius', p.radius, 'str', 'placeholder="theme"'))}</div>`);
    if (c.type === 'badge') content.push(F('Color per value <small>(value=color, one per line)</small>', ta('props.colorMap', p.colorMap, '', 4)), F('Radius', inp('props.radius', p.radius)));
  }
  if (c.type === 'icon') {
    content.push(`<div class="f"><span>Icon</span><div class="icon-grid">${ICON_NAMES.map((n) => `<button type="button" class="${n === p.icon ? 'on' : ''}" data-icon="${n}" title="${n}">${svgIcon(n, 'currentColor', 2)}</button>`).join('')}</div></div>`);
    content.push(colorField('props.color', p.color, 'Color'), F('Line width', inp('props.stroke', p.stroke, 'num', 'step="0.2" min="0.6" max="4"')));
  }
  if (c.type === 'card') {
    content.push(colorField('props.fill', p.fill, 'Fill'), colorField('props.fill2', p.fill2, 'Gradient to (optional)'));
    if (p.fill2) content.push(F('Gradient angle', inp('props.gradAngle', p.gradAngle, 'num')));
    content.push(colorField('props.borderColor', p.borderColor, 'Border color'), `<div class="grid2">${F('Border', inp('props.borderThickness', p.borderThickness, 'num'))}${F('Radius', inp('props.radius', p.radius, 'str', 'placeholder="theme"'))}</div>`, chk('props.shadow', p.shadow, 'Drop shadow'));
    content.push('<p class="note">Cards are HTML text controls: the only way to get rounded corners, gradients and shadows on a canvas. Rectangles have no radius in Power Apps.</p>');
  }
  if (c.type === 'rect') content.push(colorField('props.fill', p.fill, 'Fill'), colorField('props.borderColor', p.borderColor, 'Border color'), F('Border', inp('props.borderThickness', p.borderThickness, 'num')));
  if (c.type === 'image') content.push(`<div class="f"><span>Image</span><div class="btns"><button class="btn sm" data-cmd="pickImg">${ui('upload')}Upload</button>${p.src ? '<button class="btn sm danger" data-cmd="clearImg">Remove</button>' : ''}</div></div>`, F('Or image URL / data URI', inp('props.src', p.src && p.src.startsWith('data:') ? '(embedded image)' : p.src, 'str', p.src?.startsWith('data:') ? 'disabled' : '')), F('Fit', pick('props.fit', ['Fit', 'Fill', 'Stretch', 'Center', 'Tile'], p.fit)));
  if (c.type === 'html') content.push(F('HTML <small>(inline styles only)</small>', ta('props.html', p.html, '', 6)), chk('props.scroll', p.scroll, 'Allow scrolling (otherwise content is clipped to the box, so no scrollbar can appear)'));
  if (['textinput', 'person'].includes(c.type)) {
    content.push(F('Hint text', inp('props.hint', p.hint)));
    if (c.type === 'textinput' && !c.field) content.push(F('Default text', inp('props.defaultText', p.defaultText)), chk('props.multiline', p.multiline, 'Multiple lines'), chk('props.numeric', p.numeric, 'Numbers only'));
    content.push(F('Size', inp('props.size', p.size, 'num')));
  }
  if (c.type === 'dropdown') { if (!c.field) content.push(F('Choices <small>(one per line)</small>', ta('props.items', p.items, '', 4))); content.push(F('Size', inp('props.size', p.size, 'num'))); }
  if (c.type === 'datepicker') content.push(F('Size', inp('props.size', p.size, 'num')));
  if (c.type === 'checkbox') { content.push(F('Text', inp('props.text', p.text)), F('Size', inp('props.size', p.size, 'num')), colorField('props.color', p.color, 'Text color')); if (!c.field) content.push(chk('props.defaultOn', p.defaultOn, 'Checked by default')); }
  if (c.type === 'toggle' && !c.field) content.push(chk('props.defaultOn', p.defaultOn, 'On by default'));
  if (c.type === 'gallery') content.push(`<div class="grid2">${F('Row height', inp('props.templateSize', p.templateSize, 'num'))}<span></span></div>`, colorField('props.templateFill', p.templateFill, 'Row fill'), colorField('props.selectedFill', p.selectedFill, 'Selected row fill'));
  if (c.type === 'timer') content.push(`<div class="grid2">${F('Duration (ms)', inp('props.duration', p.duration, 'num'))}<span></span></div>`, chk('props.repeat', p.repeat, 'Repeat'), chk('props.autoStart', p.autoStart, 'Start automatically'));
  if (c.type === 'raw') content.push(`<p class="note">Imported <b>${esc(c.paControl)}</b>. Its properties are kept exactly as pasted; edit them under Formulas. Move and resize work normally.</p>`);
  if (content.length) parts.push(`<div class="sec"><h3>${c.type === 'gallery' ? 'Rows' : 'Content and style'}</h3>${content.join('')}</div>`);
  if (INPUT_TYPES.has(c.type)) parts.push(inspField(f));
  if (c.type === 'gallery') parts.push(inspGalleryData(f));
  parts.push(inspVisible(c));
  parts.push(inspActions(f));
  parts.push(inspFx(f));
  return parts.join('');
}
function dataSourceForText(f) {
  const c = f.ctl, b = c.bind || null, inGal = !!f.parent || c.type === 'gallery';
  const kind = c.fx.Text ? 'fx' : !b ? 'text' : b.kind === 'count' ? 'count' : b.kind === 'user' ? 'user' : b.src === 'item' ? 'item' : 'sel';
  const opts = [['text', 'Fixed text'], ...(inGal ? [['item', 'A column of this gallery row']] : []), ['sel', 'A column of the selected item'], ['user', 'Signed-in user'], ['count', 'Count of items in a list'], ['fx', 'Power Fx formula']];
  let html = F('Shows', `<select data-bindkind="1">${sel(opts, kind)}</select>`);
  if (kind === 'item' || kind === 'sel') {
    html += `<div class="grid2">${F('List', pick('bind.list', [['', '(list)'], ...P.lists.map((l) => l.name)], b.list, true))}${F('Column', pick('bind.field', [['', '(column)'], ...fieldOpts(b.list), 'ID', 'Created', 'Modified'], b.field))}</div>`;
    html += F('Text before it', inp('bind.prefix', b.prefix, 'str', 'placeholder="e.g. Due: "'));
  }
  if (kind === 'user') html += `<div class="grid2">${F('Field', pick('bind.field', [['FullName', 'Full name'], ['Email', 'Email']], b.field))}${F('Text before it', inp('bind.prefix', b.prefix))}</div>`;
  if (kind === 'count') {
    const col = colOf(P, b.list, b.field);
    html += `<div class="grid2">${F('List', pick('bind.list', [['', '(list)'], ...P.lists.map((l) => l.name)], b.list, true))}${F('Where column', pick('bind.field', [['', '(all items)'], ...fieldOpts(b.list, ['Choice', 'YesNo', 'Text'])], b.field, true))}</div>`;
    if (b.field) html += F('Equals', col?.type === 'Choice' ? pick('bind.value', col.choices, b.value) : col?.type === 'YesNo' ? pick('bind.value', ['true', 'false'], b.value) : inp('bind.value', b.value));
  }
  if (kind === 'fx') html += F('Text formula', ta('fx.Text', c.fx.Text, '', 3));
  return html;
}
function inspField(f) {
  const c = f.ctl, fl = c.field || {};
  const types = INPUT_COLTYPES[c.type];
  const ok = P.lists.length;
  return `<div class="sec"><h3>Saves to</h3>${ok ? `<div class="grid2">${F('List', pick('field.list', [['', '(not saved)'], ...P.lists.map((l) => l.name)], fl.list, true))}${F('Column', pick('field.column', [['', '(column)'], ...fieldOpts(fl.list, types)], fl.column, true))}</div>
    <p class="note">${fl.list ? `A Save action on this screen writes this input to <b>${esc(fl.list)}.${esc(fl.column || '?')}</b>, and it loads the selected item's value when the screen opens.` : `Bind it to a ${types.join(' / ')} column so a Save action writes it.`}${c.type === 'person' ? ' People pickers need the Office 365 Users data source.' : ''}</p>` : '<p class="note">Add a list on the Data tab to bind this input.</p>'}</div>`;
}
function inspGalleryData(f) {
  const c = f.ctl, d = c.data;
  const inputs = screen().controls.filter((x) => x.type === 'textinput').map((x) => x.name);
  const col = colOf(P, d.list, d.filterField);
  return `<div class="sec"><h3>Data</h3>
    ${F('Items from list', pick('data.list', [['', '(sample numbers)'], ...P.lists.map((l) => l.name)], d.list, true))}
    ${d.list ? `<div class="grid2">${F('Search box', pick('data.searchCtl', [['', '(none)'], ...inputs], d.searchCtl))}${F('Searches column', pick('data.searchField', fieldOpts(d.list, ['Text']), d.searchField))}</div>
    <div class="grid2">${F('Only show where', pick('data.filterField', [['', '(show all)'], ['CreatedByMe', 'I created it'], ...fieldOpts(d.list, ['Choice', 'YesNo', 'Person', 'Text', 'Number'])], d.filterField, true))}${d.filterField && d.filterField !== 'CreatedByMe' ? F('Equals', col?.type === 'Choice' ? pick('data.filterValue', col.choices, d.filterValue) : col?.type === 'Person' ? '<input disabled value="the signed-in user">' : col?.type === 'YesNo' ? pick('data.filterValue', ['true', 'false'], d.filterValue) : inp('data.filterValue', d.filterValue)) : '<span></span>'}</div>
    <div class="grid2">${F('Sort by', pick('data.sortField', [['', '(none)'], 'Modified', 'Created', 'ID', ...fieldOpts(d.list, ['Text', 'Number', 'Currency', 'DateTime'])], d.sortField))}${F('Order', pick('data.sortDir', ['Descending', 'Ascending'], d.sortDir))}</div>
    ${F('Extra filter (Power Fx, optional)', ta('data.extraFilter', d.extraFilter, '', 2, 'Priority.Value = "High"'))}
    ${(c.children || []).length ? '' : '<button class="btn sm" data-cmd="galRows">Add a row layout</button>'}` : ''}
    <p class="note">Search uses StartsWith, which SharePoint can run server-side (delegable).</p></div>`;
}
function inspVisible(c) {
  const v = c.visibleWhen || {};
  const vars = knownVars();
  return `<div class="sec"><h3>Visible</h3>
    ${F('Show this control', `<select data-vis="mode">${sel([['always', 'Always'], ['var', 'Only when a variable…'], ['fx', 'Power Fx condition']], c.fx.Visible ? 'fx' : v.var ? 'var' : 'always')}</select>`)}
    ${v.var && !c.fx.Visible ? `<div class="grid2">${F('Variable', `<input data-p="visibleWhen.var" data-t="ident" list="varList" value="${esc(v.var)}"><datalist id="varList">${vars.map((x) => `<option value="${esc(x)}">`).join('')}</datalist>`)}${F('Is', pick('visibleWhen.op', [['true', 'true'], ['false', 'false'], ['eq', 'equal to'], ['neq', 'not equal to']], v.op || 'true', true))}</div>${v.op === 'eq' || v.op === 'neq' ? F('Value', inp('visibleWhen.value', v.value)) : ''}
      <p class="note">Use "Toggle variable" or "Set variable" actions on a button to show and hide it. Try it in Play.</p>` : ''}
    ${c.fx.Visible ? F('Condition', ta('fx.Visible', c.fx.Visible, '', 2)) : ''}</div>`;
}
function knownVars() {
  const set = new Set();
  for (const s of P.screens) for (const [c] of walk(s.controls)) {
    if (c.visibleWhen?.var) set.add(c.visibleWhen.var);
    Object.values(c.actions || {}).flat().forEach((a) => { if ((a.type === 'setVar' || a.type === 'toggleVar') && a.name) set.add(a.name); });
  }
  return [...set];
}
function eventsFor(c) {
  if (c.type === 'timer') return ['OnTimerEnd'];
  if (c.type === 'textinput' || c.type === 'dropdown') return ['OnChange'];
  if (c.type === 'toggle' || c.type === 'checkbox') return ['OnCheck', 'OnUncheck'];
  if (c.type === 'person' || c.type === 'datepicker' || c.type === 'raw') return [];
  return ['OnSelect'];
}
function inspActions(f) {
  const c = f.ctl, evs = eventsFor(c);
  if (!evs.length) return '';
  if (!evs.includes(S.actEvent)) S.actEvent = evs[0];
  const ev = S.actEvent, acts = c.actions[ev] || [];
  const screensOpt = [['', '(screen)'], ...P.screens.map((s) => [s.id, s.name])];
  const card = (a, i) => {
    const pth = (k) => `data-o="act" data-ev="${ev}" data-i="${i}" data-p="${k}"`;
    const S2 = (k, opts, v, re) => `<select ${pth(k)} ${re ? 'data-re="1"' : ''}>${sel(opts, v)}</select>`;
    const I2 = (k, v, ph = '') => `<input ${pth(k)} value="${esc(v ?? '')}" placeholder="${esc(ph)}">`;
    const T2 = (k, v, ph = '') => `<textarea ${pth(k)} rows="3" placeholder="${esc(ph)}">${esc(v ?? '')}</textarea>`;
    const lists = [['', '(list)'], ...P.lists.map((l) => l.name)];
    const typeOpts = Object.entries(ACTION_TYPES).filter(([k]) => k !== 'openItem' || f.parent || c.type === 'gallery');
    let body = '';
    switch (a.type) {
      case 'navigate': body = `<div class="grid2">${F('Screen', S2('screen', screensOpt, a.screen))}${F('Transition', S2('transition', ['Fade', 'None', 'Cover', 'UnCover', 'CoverRight', 'UnCoverRight'], a.transition || 'Fade'))}</div>`; break;
      case 'newItem': case 'openItem': body = `<div class="grid2">${F('List', S2('list', lists, a.list))}${F('Then go to', S2('screen', [['', '(stay)'], ...screensOpt.slice(1)], a.screen))}</div>`; break;
      case 'save': body = `<div class="grid2">${F('List', S2('list', lists, a.list, true))}${F('After saving', S2('then', [['back', 'Go back'], ['navigate', 'Go to screen'], ['stay', 'Stay, clear the form']], a.then || 'back', true))}</div>${a.then === 'navigate' ? F('Screen', S2('screen', screensOpt, a.screen)) : ''}${F('Message', I2('message', a.message, 'Saved'))}<p class="note">Writes every input on this screen that is bound to ${esc(a.list || 'the list')}. New item if nothing is selected, otherwise updates it. Disabled until required fields are filled.</p>`; break;
      case 'delete': body = `<div class="grid2">${F('List', S2('list', lists, a.list))}${F('Then', S2('then', [['back', 'Go back'], ['', 'Stay'], ['navigate', 'Go to screen']], a.then ?? 'back', true))}</div>${a.then === 'navigate' ? F('Screen', S2('screen', screensOpt, a.screen)) : ''}`; break;
      case 'setField': {
        const col = colOf(P, a.list, a.field);
        body = `<div class="grid2">${F('List', S2('list', lists, a.list, true))}${F('Column', S2('field', [['', '(column)'], ...fieldOpts(a.list, ['Choice', 'YesNo', 'Text', 'Number', 'Currency', 'DateTime', 'Person'])], a.field, true))}</div>
          ${col ? F('New value', col.type === 'Choice' ? S2('value', col.choices, a.value) : col.type === 'YesNo' ? S2('value', ['true', 'false'], a.value) : col.type === 'DateTime' ? '<input disabled value="Today">' : col.type === 'Person' ? '<input disabled value="The signed-in user">' : I2('value', a.value)) : ''}${F('Message', I2('message', a.message, 'Updated'))}`;
        break;
      }
      case 'setVar': body = `<div class="grid2">${F('Variable', I2('name', a.name, 'varShowPanel'))}${F('Value (Power Fx)', I2('value', a.value, 'true'))}</div>`; break;
      case 'toggleVar': body = F('Variable', I2('name', a.name, 'varShowPanel')); break;
      case 'notify': body = `<div class="grid2">${F('Message', I2('text', a.text))}${F('Kind', S2('kind', ['Information', 'Success', 'Warning', 'Error'], a.kind || 'Information'))}</div>`; break;
      case 'launch': body = F('URL', I2('url', a.url, 'https://')); break;
      case 'refresh': body = F('List', S2('list', lists, a.list)); break;
      case 'runFlow': {
        const bf = (P.flows || []).filter((x) => x.kind === 'button');
        body = bf.length ? `<div class="grid2">${F('Flow', S2('flow', [['', '(flow)'], ...bf.map((x) => [x.id, x.name])], a.flow))}${F('Then', S2('then', [['', 'Stay'], ['back', 'Go back'], ['navigate', 'Go to screen']], a.then || '', true))}</div>${a.then === 'navigate' ? F('Screen', S2('screen', screensOpt, a.screen)) : ''}${F('Message when done', I2('message', a.message, 'Done'))}<p class="note">Sends the selected item's ID and waits for the flow to answer.</p>`
          : '<p class="note">Add a flow of the <b>Run from a button</b> kind on the Flows tab first.</p>';
        break;
      }
      case 'decide': {
        const boxes = [['', '(comment box)'], ...[...walk(f.screen.controls)].map(([x]) => x).filter((x) => x.type === 'textinput').map((x) => x.name)];
        body = `<div class="grid2">${F('List', S2('list', lists, a.list))}${F('Decision', S2('decision', ['Approved', 'Rejected'], a.decision || 'Approved'))}</div>${F('Comment box', S2('commentCtl', boxes, a.commentCtl))}<p class="note">${P.app.chain ? `Records the decision for the item's current stage of the ${esc(P.app.chain.list)} approval chain, moves it to the next stage and adds a history line. Reject needs a comment.` : 'Turn on an approval chain in the wizard first.'}</p>`;
        break;
      }
      case 'resubmit': body = `${F('List', S2('list', lists, a.list))}<p class="note">Sends a rejected item back to stage 1 with every stage reset to Pending.</p>`; break;
      case 'reassign': {
        const pk = [['', '(people picker)'], ...[...walk(f.screen.controls)].map(([x]) => x).filter((x) => x.type === 'person' && !x.field).map((x) => x.name)];
        body = `${F('List', S2('list', lists, a.list))}${F('People picker', S2('picker', pk, a.picker))}<p class="note">${P.app.chain?.reassign ? 'Hands the waiting stage to the picked person (a stand-in; the stage\'s approver column is kept) and adds a history line.' : 'Turn on "Approvers can reassign" for the approval chain first (wizard or Ask for a change).'}</p>`;
        break;
      }
      case 'override': {
        const boxes = [['', '(reason box)'], ...[...walk(f.screen.controls)].map(([x]) => x).filter((x) => x.type === 'textinput').map((x) => x.name)];
        body = `<div class="grid2">${F('List', S2('list', lists, a.list))}${F('Decision', S2('decision', ['Approved', 'Rejected'], a.decision || 'Approved'))}</div>${F('Reason box', S2('commentCtl', boxes, a.commentCtl))}<p class="note">${P.app.chain?.override ? 'Decides the waiting stage for an admin, whoever the approver is. The reason is required and goes into the history.' : 'Turn on "Admins can override" for the approval chain first.'}</p>`;
        break;
      }
      case 'custom': body = F('Power Fx', T2('fx', a.fx, 'Set(varX, true); Notify("Hi")')); break;
    }
    return `<div class="act"><div class="act-h"><span class="act-n">${i + 1}</span><select ${pth('type')} data-re="1">${sel(typeOpts, a.type)}</select>
      <button class="mini" data-acmd="up" data-i="${i}" title="Move up">${ui('up')}</button><button class="mini danger" data-acmd="del" data-i="${i}" title="Remove">${ui('trash')}</button></div>${body}</div>`;
  };
  return `<div class="sec"><h3>Actions <small>${c.type === 'gallery' ? 'when a row is tapped' : ''}</small></h3>
    ${evs.length > 1 ? `<div class="seg" style="margin-bottom:10px">${evs.map((e) => `<button data-actev="${e}" class="${e === ev ? 'on' : ''}">${e}</button>`).join('')}</div>` : ''}
    ${acts.map(card).join('')}
    <button class="btn sm" data-acmd="add">${ui('plus')}Add action</button></div>`;
}
const FX_PROPS = ['Visible', 'DisplayMode', 'Text', 'Fill', 'Color', 'BorderColor', 'Tooltip', 'Items', 'Default', 'HoverFill', 'Size', 'X', 'Y', 'Width', 'Height', 'OnSelect', 'OnChange', 'TabIndex', 'AccessibleLabel'];
function inspFx(f) {
  const c = f.ctl;
  const rows = Object.entries(c.fx || {}).filter(([k]) => !(k === 'Text' && dataSourceKind(c) === 'fx') && k !== 'Visible').map(([k, v]) => `<div class="fxrow"><div class="h"><span>${esc(k)}</span><button class="mini danger" data-fxdel="${esc(k)}" title="Remove">${ui('close')}</button></div><textarea class="inp" rows="2" data-p="fx.${esc(k)}">${esc(v)}</textarea></div>`).join('');
  let yaml = ''; try { yaml = controlYaml(P, f.screen, c); } catch (e) { yaml = String(e); }
  return `<div class="sec"><h3>Formulas <small>advanced</small></h3>
    <p class="note">Override any property with your own Power Fx. It replaces what the designer generates.</p>${rows}
    <div class="grid2" style="align-items:end"><label class="f"><span>Property</span><input id="fxNew" list="fxList" placeholder="DisplayMode"><datalist id="fxList">${FX_PROPS.map((x) => `<option value="${x}">`).join('')}</datalist></label><button class="btn sm" data-cmd="fxAdd" style="margin-bottom:10px">Add formula</button></div>
    <details><summary class="note" style="cursor:pointer">Generated YAML for this control</summary><pre class="code small">${esc(yaml)}</pre></details></div>`;
}
const dataSourceKind = (c) => (c.fx.Text ? 'fx' : !c.bind ? 'text' : c.bind.kind || c.bind.src);

// inspector events
function setPath(obj, path, val) {
  const ks = path.split('.'); let o = obj;
  for (let i = 0; i < ks.length - 1; i++) { if (o[ks[i]] == null || typeof o[ks[i]] !== 'object') o[ks[i]] = {}; o = o[ks[i]]; }
  o[ks[ks.length - 1]] = val;
}
function onInspectorInput(e) {
  const t = e.target;
  if (t.dataset.bindkind) return onBindKind(t, e);
  if (t.dataset.vis) return onVisMode(t, e);
  if (!t.dataset.p) return;
  const o = t.dataset.o || 'ctl', path = t.dataset.p, type = t.dataset.t || 'str';
  let val = type === 'bool' ? t.checked : t.value;
  if (type === 'num') { if (t.value === '' || isNaN(Number(t.value))) return; val = Number(t.value); }
  if (o === 'screen') {
    const scr = screen();
    if (path === 'name') {
      if (e.type !== 'change') return;
      const nv = t.value.trim();
      if (!IDENT.test(nv) || RESERVED.has(nv) || (nv !== scr.name && allNames(P).has(nv))) { t.classList.add('bad'); toast('That name is invalid or already used', true); return; }
      t.classList.remove('bad'); scr.name = nv; renderTabs(); renderPanelIf('screens'); commit(); return;
    }
    setPath(scr, path, val); renderCanvas(); commitSoon(); return;
  }
  if (o === 'act') {
    const f = one(); if (!f) return;
    const a = f.ctl.actions[t.dataset.ev][+t.dataset.i];
    setPath(a, path, val);
    if (path === 'type') { for (const k of Object.keys(a)) if (k !== 'type') delete a[k]; defaultsForAction(a, f); }
    if (path === 'list' && a.type === 'setField') a.field = '';
    renderCanvas(); commitSoon();
    if (t.dataset.re && e.type === 'change') renderInspector();
    return;
  }
  const f = one(); if (!f) return;
  const c = f.ctl;
  if (path === 'name') {
    const nv = t.value.trim();
    const bad = !IDENT.test(nv) || RESERVED.has(nv) || (nv !== c.name && allNames(P, c.id).has(nv)) || companions({ ...c, name: nv }).some((x) => allNames(P, c.id).has(x));
    t.classList.toggle('bad', bad); t.title = bad ? 'Use letters, digits and _; must be unique across the app' : '';
    if (bad || e.type !== 'change') return;
    renameRefs(c.name, nv); c.name = nv; drawOverlay(); renderPanelIf('layers'); commit(); renderInspector(); return;
  }
  if (path.startsWith('fx.') && val === '' && e.type === 'change') { delete c.fx[path.slice(3)]; }
  else setPath(c, path, val);
  if (path === 'field.list') { c.field.column = ''; if (!val) delete c.field; }
  if (path === 'field.column' && c.field?.column) {
    const col = colOf(P, c.field.list, c.field.column);
    const lab = (f.parent ? f.parent.children : f.screen.controls).find((x) => x.type === 'label' && Math.abs(x.x - c.x) < 4 && c.y - (x.y + x.h) >= -2 && c.y - (x.y + x.h) < 12 && !x.bind);
    if (lab && col && /^Text|Label|New/.test(lab.props.text || '')) lab.props.text = pretty(col.name) + (col.required ? ' *' : '');
  }
  if (path === 'bind.list') c.bind.field = '';
  if (path === 'data.list' && val) { const l = P.lists.find((x) => x.name === val); c.data.searchField = l.columns[0].name; c.data.filterField = ''; }
  if (path === 'props.radius') c.props.radius = val === '' || val === 'theme' ? 'theme' : Number(val) || 0;
  if (path === 'visibleWhen.var' && !val) delete c.visibleWhen;
  if (path === 'props.templateSize') c.props.templateSize = Math.max(8, Math.min(c.h, val));
  if (['x', 'y', 'w', 'h'].includes(path) || path.startsWith('props.') || path.startsWith('bind.') || path.startsWith('field.') || path.startsWith('data.') || path.startsWith('fx.') || path.startsWith('visibleWhen.')) renderCanvas();
  commitSoon();
  if ((t.dataset.re && e.type === 'change') || (t.type === 'checkbox' && path === 'props.multiline')) renderInspector();
  if (t.closest('.colorf') && e.type === 'change') renderInspector();
}
function renameRefs(oldN, newN) {
  const re = new RegExp(`\\b${oldN}\\b`, 'g');
  for (const s of P.screens) for (const [c] of walk(s.controls)) {
    if (c.data?.searchCtl === oldN) c.data.searchCtl = newN;
    for (const k of Object.keys(c.fx || {})) c.fx[k] = String(c.fx[k]).replace(re, newN);
    Object.values(c.actions || {}).flat().forEach((a) => { if (a.type === 'custom') a.fx = String(a.fx || '').replace(re, newN); });
  }
}
function defaultsForAction(a, f) {
  const l = P.lists[0]?.name || '';
  if (['save', 'delete', 'newItem', 'openItem', 'refresh', 'setField'].includes(a.type)) a.list = (f.ctl.data?.list) || f.ctl.field?.list || l;
  if (a.type === 'save') a.then = 'back';
  if (a.type === 'delete') a.then = 'back';
  if (a.type === 'setVar') { a.name = 'varShow'; a.value = 'true'; }
  if (a.type === 'toggleVar') a.name = 'varShow';
  if (a.type === 'navigate') a.screen = P.screens.find((s) => s.id !== screen().id)?.id || '';
}
function onBindKind(t) {
  const f = one(); if (!f) return;
  const c = f.ctl, k = t.value, l = P.lists[0]?.name || '';
  delete c.fx.Text;
  if (k === 'text') delete c.bind;
  else if (k === 'fx') { delete c.bind; c.fx.Text = fxString(c.props.text || ''); }
  else if (k === 'user') c.bind = { kind: 'user', field: 'FullName', prefix: '' };
  else if (k === 'count') c.bind = { kind: 'count', list: f.parent?.data?.list || l, field: '', value: '' };
  else { const gl = f.parent?.data?.list || (c.type === 'gallery' ? c.data.list : ''); const list = k === 'item' ? gl || l : c.bind?.list || l; c.bind = { src: k, list, field: P.lists.find((x) => x.name === list)?.columns[0]?.name || '', prefix: '' }; }
  renderCanvas(); renderInspector(); commit();
}
function onVisMode(t) {
  const f = one(); if (!f) return;
  const c = f.ctl;
  delete c.fx.Visible; delete c.visibleWhen;
  if (t.value === 'var') c.visibleWhen = { var: knownVars()[0] || 'varShow', op: 'true', value: '' };
  if (t.value === 'fx') c.fx.Visible = 'true';
  renderCanvas(); renderInspector(); commit();
}
async function onInspectorClick(e) {
  const b = e.target.closest('button'); if (!b) return;
  const f = one();
  if (b.dataset.color) return openColorPop(b, b.dataset.color);
  if (b.dataset.align) return alignSel(b.dataset.align);
  if (b.dataset.icon && f) { f.ctl.props.icon = b.dataset.icon; renderCanvas(); renderInspector(); commit(); return; }
  if (b.dataset.actev) { S.actEvent = b.dataset.actev; renderInspector(); return; }
  if (b.dataset.fxdel && f) { delete f.ctl.fx[b.dataset.fxdel]; renderCanvas(); renderInspector(); commit(); return; }
  if (b.dataset.acmd && f) {
    const ev = S.actEvent; const acts = (f.ctl.actions[ev] = f.ctl.actions[ev] || []); const i = +b.dataset.i;
    if (b.dataset.acmd === 'add') { const a = { type: f.ctl.type === 'gallery' ? 'openItem' : 'navigate' }; defaultsForAction(a, f); acts.push(a); }
    if (b.dataset.acmd === 'del') acts.splice(i, 1);
    if (b.dataset.acmd === 'up' && i > 0) acts.splice(i - 1, 0, acts.splice(i, 1)[0]);
    renderCanvas(); renderInspector(); commit(); return;
  }
  const cmd = b.dataset.cmd;
  if (cmd === 'del') deleteSel();
  if (cmd === 'dup') { copySel(); pasteCtls(S.clip); }
  if (cmd === 'saveBlock') saveSelAsBlock();
  if (cmd === 'z') zMove(b.dataset.z);
  if (cmd === 'fullw' && f) { f.ctl.x = 0; f.ctl.w = ff().w; renderCanvas(); renderInspector(); commit(); }
  if (cmd === 'makeStart') { P.app.startScreen = screen().id; renderTabs(); renderInspector(); renderPanelIf('screens'); commit(); }
  if (cmd === 'copyScreen') copyText(screenYaml(P, screen()));
  if (cmd === 'pickImg') { const fi = $('#fileImg'); fi.value = ''; fi.click(); }
  if (cmd === 'clearImg' && f) { f.ctl.props.src = ''; renderCanvas(); renderInspector(); commit(); }
  if (cmd === 'galRows' && f) { const l = P.lists.find((x) => x.name === f.ctl.data.list); if (l) { galleryRows(P, screen(), f.ctl, l); renderCanvas(); renderInspector(); commit(); } }
  if (cmd === 'fxAdd' && f) {
    const k = $('#fxNew').value.trim(); if (!IDENT.test(k)) { toast('Type a property name', true); return; }
    f.ctl.fx[k] = f.ctl.fx[k] || ''; renderInspector(); commit();
    setTimeout(() => $(`#inspector textarea[data-p="fx.${k}"]`)?.focus(), 0);
  }
}
function openColorPop(btn, path) {
  $$('.colorpop').forEach((x) => x.remove());
  const isScreen = path.startsWith('screen:'), key = path.replace(/^screen:/, '');
  const target = isScreen ? screen() : one()?.ctl; if (!target) return;
  const get = () => key.split('.').reduce((o, k) => o?.[k], target);
  const cur = get();
  const pop = document.createElement('div'); pop.className = 'colorpop';
  const toks = THEME_COLORS.map(([k]) => k);
  pop.innerHTML = `<div class="sub" style="margin-top:0">Theme colors (follow theme changes)</div><div class="toks">${toks.map((k) => `<button title="${k}" data-tok="@${k}" class="${cur === '@' + k ? 'on' : ''}" style="background:${P.theme.colors[k]}"></button>`).join('')}</div>
    <div class="row"><input type="color" value="${isHex(cur) ? cur.slice(0, 7) : '#0891b2'}"><button class="btn sm" data-tok="transparent">Transparent</button><button class="btn sm" data-close="1">Done</button></div>`;
  document.body.appendChild(pop);
  const r = btn.getBoundingClientRect();
  pop.style.left = Math.min(window.innerWidth - 260, r.left - 210) + 'px'; pop.style.top = Math.min(window.innerHeight - 160, r.bottom + 6) + 'px';
  const set = (v) => { setPath(target, key, v); renderCanvas(); renderInspector(); commitSoon(); };
  pop.addEventListener('click', (e) => { const t = e.target.closest('[data-tok]'); if (t) { set(t.dataset.tok); pop.remove(); } if (e.target.closest('[data-close]')) pop.remove(); });
  pop.querySelector('input[type=color]').addEventListener('input', (e) => set(e.target.value.toUpperCase()));
  setTimeout(() => document.addEventListener('pointerdown', function off(ev) { if (!pop.contains(ev.target)) { pop.remove(); document.removeEventListener('pointerdown', off); } }), 0);
}

// ------------------------------------------------------------------ modals
function modal(html, cls = '') {
  const opener = document.activeElement;
  const bg = document.createElement('div'); bg.className = 'modal-bg';
  bg.innerHTML = `<div class="modal ${cls}" role="dialog" aria-modal="true">${html}</div>`;
  $('#modalRoot').appendChild(bg);
  const title = bg.querySelector('.modal-h h2');
  if (title) { title.id = 'mdl' + uid(); bg.firstElementChild.setAttribute('aria-labelledby', title.id); }
  // focus the first field (or button) inside, and give focus back to the opener on close
  setTimeout(() => (bg.querySelector('.modal-b input, .modal-b textarea, .modal-b select, .modal-b button, .modal-f .pri') || bg.querySelector('button'))?.focus(), 0);
  new MutationObserver((_, ob) => { if (!bg.isConnected) { ob.disconnect(); if (opener?.isConnected) opener.focus(); } }).observe($('#modalRoot'), { childList: true });
  bg.addEventListener('keydown', (e) => {
    if (e.key !== 'Tab') return;
    const f = [...bg.querySelectorAll('button, input, select, textarea, [tabindex="0"]')].filter((x) => !x.disabled && x.offsetParent);
    if (!f.length) return;
    if (e.shiftKey && document.activeElement === f[0]) { e.preventDefault(); f[f.length - 1].focus(); }
    else if (!e.shiftKey && document.activeElement === f[f.length - 1]) { e.preventDefault(); f[0].focus(); }
  });
  const close = () => bg.remove();
  bg.addEventListener('pointerdown', (e) => { if (e.target === bg) bg.dataset.down = '1'; });
  bg.addEventListener('click', (e) => { if (e.target === bg && bg.dataset.down) close(); bg.dataset.down = ''; });
  bg.querySelectorAll('[data-close]').forEach((x) => x.addEventListener('click', close));
  return { el: bg.firstElementChild, close };
}
function ask(title, msg, okText = 'OK') {
  return new Promise((res) => {
    const m = modal(`<div class="modal-h"><h2>${esc(title)}</h2></div><div class="modal-b"><p style="margin:0">${esc(msg)}</p></div><div class="modal-f"><button class="btn" data-no>Cancel</button><button class="btn pri" data-ok>${esc(okText)}</button></div>`, 'sm');
    m.el.querySelector('[data-ok]').onclick = () => { m.close(); res(true); };
    m.el.querySelector('[data-no]').onclick = () => { m.close(); res(false); };
  });
}
function askText(title, label, value = '', area = false) {
  return new Promise((res) => {
    const m = modal(`<div class="modal-h"><h2>${esc(title)}</h2></div><div class="modal-b"><label class="f"><span>${esc(label)}</span>${area ? `<textarea class="inp" rows="12" id="askv">${esc(value)}</textarea>` : `<input id="askv" value="${esc(value)}">`}</label></div><div class="modal-f"><button class="btn" data-no>Cancel</button><button class="btn pri" data-ok>OK</button></div>`, area ? 'md' : 'sm');
    const v = m.el.querySelector('#askv'); setTimeout(() => v.focus(), 0);
    m.el.querySelector('[data-ok]').onclick = () => { m.close(); res(v.value); };
    m.el.querySelector('[data-no]').onclick = () => { m.close(); res(null); };
  });
}

// new screen from template
function newScreenDialog() {
  const lists = P.lists;
  const m = modal(`<div class="modal-h"><h2>New screen</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b"><div class="picks" id="tplPicks">${Object.entries(SCREEN_TEMPLATES).map(([k, t], i) => `<label><input type="radio" name="tpl" value="${k}" ${i === 0 ? 'checked' : ''}><span><b>${esc(t.label)}</b><small>${{ blank: 'A header and an empty canvas.', home: 'Welcome, KPI counts and tiles to your other screens.', browse: 'Search box and a gallery of the list; tap a row to open it.', detail: 'Every column of the selected item, with Edit and Delete.', edit: 'Inputs for every column with Save and Cancel, for new and existing items.', approvals: 'Queue of items waiting for a decision with Approve / Reject.', loading: 'Animated monogram and progress ring, then opens your home screen. Becomes the start screen.' }[k]}</small></span></label>`).join('')}</div>
      <div class="grid2">${F('List', `<select id="tplList">${listOptions(lists[0]?.name, !lists.length, '(no lists yet)')}</select>`)}${F('Screen name (optional)', '<input id="tplName" placeholder="scrSomething">')}</div></div>
    <div class="modal-f"><button class="btn" data-close>Cancel</button><button class="btn pri" id="tplGo">Add screen</button></div>`, 'md');
  m.el.querySelector('#tplGo').onclick = () => {
    const k = m.el.querySelector('input[name=tpl]:checked').value, list = P.lists.find((l) => l.name === m.el.querySelector('#tplList').value);
    let name = m.el.querySelector('#tplName').value.trim();
    if (name && (!IDENT.test(name) || allNames(P).has(name))) { toast('That screen name is invalid or taken', true); return; }
    if (SCREEN_TEMPLATES[k].needsList && !list) { toast('This template needs a list. Add one on the Data tab.', true); return; }
    const opts = name ? { name } : {};
    const findScr = (pfx) => P.screens.find((s) => s.name === pfx + core.slug(list?.name || ''));
    let scr;
    if (k === 'blank') scr = genBlank(P, name || 'scrScreen');
    if (k === 'home') { const tiles = P.screens.filter((s) => /^scrBrowse/.test(s.name)).map((s) => ({ label: pretty(s.name.replace(/^scrBrowse/, '')), sub: 'Open ' + pretty(s.name.replace(/^scrBrowse/, '')).toLowerCase(), icon: 'list', actions: [{ type: 'navigate', screen: s.id }] })); scr = genHome(P, { ...opts, tiles }); }
    if (k === 'browse') scr = genBrowse(P, list, { ...opts, editScreen: findScr('scrEdit'), detailScreen: findScr('scrDetail'), back: P.screens.length > 0 });
    if (k === 'detail') scr = genDetail(P, list, { ...opts, editScreen: findScr('scrEdit') });
    if (k === 'edit') scr = genEdit(P, list, opts);
    if (k === 'approvals') scr = genApprovals(P, list, { ...opts, back: true });
    if (k === 'loading') {
      if (P.screens.some((s) => s.role === 'loading')) { toast('This app already has a loading screen', true); return; }
      const next = P.screens.find((s) => /^scrHome/.test(s.name)) || P.screens[0];
      if (!next) { toast('Add the screen it should open first (for example a Home dashboard)', true); return; }
      P.app.features = { ...(P.app.features || {}), loading: true };
      scr = genLoading(P, { next: next.id, lists: P.lists.map((l) => l.name) });
      P.app.startScreen = scr.id;
    }
    S.screenId = scr.id; S.sel = []; m.close(); renderAll(); commit();
  };
}

// scenario gallery: thumbnails of each scenario's home screen, cached per theme
const THUMBS = new Map(), INCLUDES = new Map();
function galleryIncludes(k) { if (!INCLUDES.has(k)) { try { INCLUDES.set(k, scenarioIncludes(k)); } catch { INCLUDES.set(k, ''); } } return INCLUDES.get(k); }
function envFor(p) {
  const sample = {}; for (const l of p.lists) sample[l.name] = sampleRows(l, p.lists, 6);
  const vars = { varIsAdmin: true, locInited: true, locStep: 2 }; for (const l of p.lists) vars[varSel(l.name)] = sample[l.name]?.[0] || {};
  const noop = () => {};
  return { project: p, theme: p.theme, mode: 'design', sample, rt: { db: sample, vars, inputs: {}, galSel: {}, ctlValues: {}, api: { navigate: noop, back: noop, toast: noop, launch: noop } } };
}
function scenarioThumb(k, theme) {
  const key = k + '|' + theme;
  if (THUMBS.has(key)) return THUMBS.get(key).cloneNode(true);
  const p = buildScenario(k, { theme, flows: false });
  const scr = p.screens.find((s) => s.name === 'scrHome') || p.screens.find((s) => !/^scrLoading/.test(s.name)) || p.screens[0];
  const host = document.createElement('div'); host.className = 'screen g-screen';
  renderScreen(host, scr, envFor(p));
  THUMBS.set(key, host);
  return host.cloneNode(true);
}
function fillThumbs(root, wz) {
  const slots = [...root.querySelectorAll('[data-thumb]')].filter((s) => s.dataset.thumb !== 'blank');
  const theme = wz.theme || THEME_PRESETS[0].name;
  const next = () => {
    const slot = slots.shift();
    if (!slot || !root.isConnected) return;
    try {
      const el = scenarioThumb(slot.dataset.thumb, theme);
      el.style.transform = `scale(${(slot.clientWidth || 280) / 1366})`;
      slot.innerHTML = ''; slot.appendChild(el);
    } catch (e) { slot.innerHTML = '<span class="g-err">Preview unavailable</span>'; }
    requestAnimationFrame(next);
  };
  requestAnimationFrame(next);
}
/** Open Play on a scenario's finished app without touching the current project. */
function previewScenario(k, wz) {
  let p;
  try { p = buildScenario(k, { formFactor: wz?.ff, theme: wz?.theme }); } catch (e) { toast('Could not build the preview: ' + e.message, true); return; }
  play({ project: p, preview: true });
}

// requirements wizard
function wizard(first = false) {
  const wz = { step: 1, scen: 'requests', name: SCENARIOS.requests.app, ff: P?.app.formFactor || 'tablet', lists: SCENARIOS.requests.lists(), screens: new Set(['home', 'browse', 'detail', 'edit']), approvals: true, theme: P?.theme?.name || THEME_PRESETS[0].name, chain: { on: false, stages: [{ name: 'Manager', approver: 'column', column: '', groupId: '', notify: '' }, { name: 'Finance', approver: 'group', column: '', groupId: '', notify: '' }] }, ft: { loading: true, confirm: true, stepper: true, scope: 'all', usage: false, adminGroup: '' } };
  const m = modal(`<div class="modal-h"><h2>${first ? 'Build a Power App' : 'New app from requirements'}</h2><div class="steps" id="wzSteps"></div><button class="mini" data-close>${ui('close')}</button></div><div class="modal-b" id="wzBody"></div>
    <div class="modal-f"><span class="grow" id="wzNote"></span>${first ? '<button class="btn" id="wzDemo">Explore a finished demo</button><button class="btn" id="wzBlank">Start with a blank canvas</button>' : ''}<button class="btn" id="wzBack">Back</button><button class="btn pri" id="wzNext">Next</button></div>`);
  const body = m.el.querySelector('#wzBody');
  const draw = () => {
    m.el.querySelector('#wzSteps').innerHTML = ['What', 'Data', 'Screens and look'].map((t, i) => `${i ? '<i></i>' : ''}<span class="${wz.step === i + 1 ? 'on' : wz.step > i + 1 ? 'done' : ''}"><b>${i + 1}</b>${t}</span>`).join('');
    m.el.querySelector('#wzBack').style.visibility = wz.step === 1 ? 'hidden' : 'visible';
    m.el.querySelector('#wzNext').textContent = wz.step === 3 ? 'Generate app' : 'Next';
    m.el.querySelector('#wzNote').textContent = wz.step === 1 ? 'You can change everything afterwards in the designer.' : wz.step === 2 ? 'These become SharePoint lists. Required columns get validation in the form.' : '';
    if (wz.step === 1) {
      body.innerHTML = `<button class="describe-cta" id="wzDescribe"><i>${svgIcon('edit', '#fff', 2)}</i><span><b>Describe it in your own words</b><small>Write the requirements in plain English. Claude plans the lists, screens and flows, and you review the plan before anything is built. Needs your Anthropic API key.</small></span><em>Describe →</em></button>
        <div class="sub" style="margin-top:16px">Or start from a scenario</div>
        <div class="gallery">${Object.entries(SCENARIOS).map(([k, s]) => `<div class="g-card ${k === wz.scen ? 'on' : ''}" data-scen="${k}" tabindex="0" role="button" aria-pressed="${k === wz.scen}" aria-label="${esc(s.label)}">
          <div class="g-thumb" data-thumb="${k}">${k === 'blank' ? `<i class="g-blank">${svgIcon(s.icon, 'currentColor', 1.6)}</i>` : '<span class="g-load"></span>'}</div>
          <div class="g-txt"><b>${esc(s.label)}</b><small>${esc(s.blurb)}</small>${k === 'blank' ? '' : `<em>${esc(galleryIncludes(k))}</em>`}</div>
          <div class="g-acts">${k === 'blank' ? '' : `<button class="btn sm" data-gprev="${k}">${ui('eye')}Preview</button>`}<button class="btn sm pri" data-guse="${k}">Use this</button></div></div>`).join('')}</div>
        <div class="grid2" style="margin-top:18px">${F('App name', `<input id="wzName" value="${esc(wz.name)}">`)}${F('Format', `<select id="wzFF">${Object.entries(FORM_FACTORS).map(([k, v]) => `<option value="${k}" ${k === wz.ff ? 'selected' : ''}>${v.label}</option>`).join('')}</select>`)}</div>`;
      const pickScen = (k) => {
        const sc = SCENARIOS[k];
        wz.scen = k; wz.name = sc.app; wz.lists = sc.lists(); wz.approvals = !!sc.approvals;
        wz.ft = { loading: true, confirm: true, stepper: true, scope: 'all', usage: false, adminGroup: wz.ft.adminGroup || '', ...(sc.features || {}) };
        wz.chain = sc.chain ? { on: true, stages: clone(sc.chain.stages), reassign: !!sc.chain.reassign, override: !!sc.chain.override } : { on: false, stages: wz.chain.on ? [{ name: 'Manager', approver: 'column', column: '', groupId: '', notify: '' }, { name: 'Finance', approver: 'group', column: '', groupId: '', notify: '' }] : wz.chain.stages };
        wz.flows = true;
      };
      body.querySelectorAll('[data-scen]').forEach((b) => {
        b.onclick = (e) => {
          const k = b.dataset.scen;
          if (e.target.closest('[data-gprev]')) { previewScenario(k, wz); return; }
          pickScen(k);
          if (e.target.closest('[data-guse]')) { wz.step = 2; draw(); return; }
          body.querySelectorAll('.g-card').forEach((x) => { x.classList.toggle('on', x === b); x.setAttribute('aria-pressed', x === b); });
          body.querySelector('#wzName').value = wz.name;
        };
        b.onkeydown = (e) => { if ((e.key === 'Enter' || e.key === ' ') && e.target === b) { e.preventDefault(); b.click(); } };
      });
      fillThumbs(body, wz);
      body.querySelector('#wzName').oninput = (e) => (wz.name = e.target.value);
      body.querySelector('#wzDescribe').onclick = () => { m.close(); describeDialog(); };
      body.querySelector('#wzFF').onchange = (e) => (wz.ff = e.target.value);
    }
    if (wz.step === 2) {
      body.innerHTML = `<div class="wide-data" id="wzData">${dataEditorHtml(wz.lists)}</div><button class="btn sm" id="wzAddList">${ui('plus')}Add list</button>`;
      const root = body.querySelector('#wzData');
      bindDataEditor(root, wz.lists, () => {}, () => { root.innerHTML = dataEditorHtml(wz.lists); });
      body.querySelector('#wzAddList').onclick = () => { let n = 'NewList', i = 2; while (wz.lists.some((l) => l.name === n)) n = 'NewList' + i++; wz.lists.push(newList(n)); root.innerHTML = dataEditorHtml(wz.lists); };
    }
    if (wz.step === 3) {
      const appr = wz.lists.filter((l) => l.columns.some((c) => c.type === 'Choice'));
      const pk = (k, t, d) => `<label><input type="checkbox" data-scr="${k}" ${wz.screens.has(k) ? 'checked' : ''}><span><b>${t}</b><small>${d}</small></span></label>`;
      body.innerHTML = `<div class="wiz-grid"><div><div class="sub" style="margin-top:0">Screens for each list</div><div class="picks">
        ${pk('home', 'Home dashboard', 'Welcome, status counts and tiles to every screen.')}
        ${pk('browse', 'Browse', 'Search plus a gallery of items.')}${pk('detail', 'Details', 'Every column of one item, with Edit and Delete.')}${pk('edit', 'New / edit form', 'Validated inputs for every column, with Save.')}
        ${appr.length ? `<label><input type="checkbox" id="wzAppr" ${wz.approvals ? 'checked' : ''}><span><b>Approvals queue</b><small>Approve / reject items in <select id="wzApprList">${appr.map((l) => `<option ${l.name === SCENARIOS[wz.scen]?.approvals?.list ? 'selected' : ''}>${esc(l.name)}</option>`).join('')}</select></small></span></label>` : ''}
      </div><div id="wzChain"></div>${(SCENARIOS[wz.scen]?.flows || []).length ? `<div class="picks"><label><input type="checkbox" id="wzFlows" ${wz.flows !== false ? 'checked' : ''}><span><b>Include the scenario's ${SCENARIOS[wz.scen].flows.length} email flows</b><small>${esc(SCENARIOS[wz.scen].flows.map((f) => f.name).join(' · '))}. Edit or remove them on the Flows tab.</small></span></label></div>` : ''}</div><div><div class="sub" style="margin-top:0">Look</div><div class="presets">${THEME_PRESETS.map((p) => `<button class="preset ${p.name === wz.theme ? 'on' : ''}" data-wtheme="${esc(p.name)}"><div class="bar">${['HeaderBg', 'Primary', 'Accent', 'Bg', 'Surface'].map((k) => `<i style="background:${p.colors[k]}"></i>`).join('')}</div><span>${esc(p.name)}</span></button>`).join('')}</div>
      <p class="note">Fine-tune colors, font and corners on the Theme tab later.</p>
      <div class="sub">Extras</div><div class="picks">
        ${[['loading', 'Animated loading screen', 'Your app\'s monogram flies in while lists load. Opens the home screen when ready.'], ['confirm', 'Confirmation page', 'A "submitted" page after a new item is saved, with Add another and Done.'], ['stepper', 'Status stepper', 'Details show the item\'s progress through its status values.'], ['mine', 'Show "Mine" first', 'Lists open filtered to items the signed-in person created.'], ['usage', 'Track app usage', 'One row per user in an AppUsage list: first and last visit, visit count.']].map(([k, t, d]) => `<label><input type="checkbox" data-ft="${k}" ${k === 'mine' ? (wz.ft.scope === 'mine' ? 'checked' : '') : wz.ft[k] ? 'checked' : ''}><span><b>${t}</b><small>${d}</small></span></label>`).join('')}
      </div>
      <label class="f"><span>Admin group ID <small>(optional: Microsoft 365 group whose members can approve and delete)</small></span><input id="wzAdmin" value="${esc(wz.ft.adminGroup)}" placeholder="00000000-0000-0000-0000-000000000000"></label></div></div>`;
      body.querySelectorAll('[data-scr]').forEach((x) => x.onchange = () => { x.checked ? wz.screens.add(x.dataset.scr) : wz.screens.delete(x.dataset.scr); });
      body.querySelector('#wzAppr')?.addEventListener('change', (e) => { wz.approvals = e.target.checked; drawChain(); });
      body.querySelector('#wzApprList')?.addEventListener('change', drawChain);
      drawChain();
      body.querySelectorAll('[data-wtheme]').forEach((b) => b.onclick = () => { wz.theme = b.dataset.wtheme; draw(); });
      body.querySelectorAll('[data-ft]').forEach((x) => x.onchange = () => { if (x.dataset.ft === 'mine') wz.ft.scope = x.checked ? 'mine' : 'all'; else wz.ft[x.dataset.ft] = x.checked; });
      body.querySelector('#wzAdmin').oninput = (e) => (wz.ft.adminGroup = e.target.value.trim());
      body.querySelector('#wzFlows')?.addEventListener('change', (e) => (wz.flows = e.target.checked));
    }
  };
  // approval chain editor (step 3, under the approvals queue)
  function drawChain() {
    const host = body.querySelector('#wzChain');
    if (!host) return;
    const list = wz.lists.find((l) => l.name === body.querySelector('#wzApprList')?.value);
    if (!wz.approvals || !list) { host.innerHTML = ''; return; }
    const people = list.columns.filter((c) => c.type === 'Person');
    const c = wz.chain;
    // stage 1 reuses the list's own approver/manager column when it has one
    if (c.on && !c.touched && c.stages[0]?.approver === 'column' && !c.stages[0].column) { const pc = people.find((x) => /approv|manag|supervis/i.test(x.name)); if (pc) c.stages[0].column = pc.name; }
    host.innerHTML = `<div class="chain-ed"><label class="chk"><input type="checkbox" id="wzChainOn" ${c.on ? 'checked' : ''}><span><b>Multi-stage approval chain</b><small>Each stage has its own approver, decision, date and comment, with a history of every decision. Off: one Approve / Reject step for admins (as before).</small></span></label>
      ${c.on ? `<div class="chain-stages">${c.stages.map((st, i) => `<div class="chain-stage" data-st="${i}"><b class="num">${i + 1}</b><div class="grow">
        <div class="row2"><input data-k="name" value="${esc(st.name)}" placeholder="Stage name" aria-label="Stage ${i + 1} name" maxlength="40">
        <select data-k="who" aria-label="Stage ${i + 1} approver">${people.map((pc) => `<option value="col:${esc(pc.name)}" ${st.approver === 'column' && st.column === pc.name ? 'selected' : ''}>Person column: ${esc(pc.name)}</option>`).join('')}<option value="new" ${st.approver === 'column' && !people.some((pc) => pc.name === st.column) ? 'selected' : ''}>New column ${esc(stageCols(st).K)}Approver</option><option value="group" ${st.approver === 'group' ? 'selected' : ''}>Microsoft 365 group</option></select>
        ${c.stages.length > 1 ? `<button class="mini" data-del="${i}" title="Remove stage" aria-label="Remove stage ${i + 1}">${ui('close')}</button>` : ''}</div>
        ${st.approver === 'group' ? `<div class="row2"><input data-k="groupId" value="${esc(st.groupId)}" placeholder="Group object ID (can be added later in Finish setup)" aria-label="Stage ${i + 1} group ID"><input data-k="notify" value="${esc(st.notify)}" placeholder="Group mailbox for emails (optional)" aria-label="Stage ${i + 1} group mailbox"></div>` : ''}
        <label class="sla">Deadline <input data-k="slaDays" type="number" min="0" max="90" step="1" value="${st.slaDays > 0 ? st.slaDays : ''}" placeholder="none" aria-label="Stage ${i + 1} deadline in days"> days <small>overdue badge and a daily reminder</small></label>
      </div></div>`).join('')}</div>
      ${c.stages.length < MAX_STAGES ? `<button class="btn sm" id="wzAddStage">${ui('plus')}Add stage</button>` : ''}<p class="note">The requester can never approve their own item. Flows email each approver in turn and the requester at the end.</p>
      <label class="chk"><input type="checkbox" id="wzChainReassign" ${c.reassign ? 'checked' : ''}><span><b>Approvers can reassign</b><small>The approver (or an admin) hands a waiting stage to someone else, who can then decide it. The history records it.</small></span></label>
      <label class="chk"><input type="checkbox" id="wzChainOverride" ${c.override ? 'checked' : ''}><span><b>Admins can override</b><small>Members of the admin group approve or reject a waiting stage with a required reason, written to the history.${wz.ft.adminGroup ? '' : ' Needs the admin group (Screens and look step, or App tab later).'}</small></span></label>` : ''}</div>`;
    host.querySelector('#wzChainOn').onchange = (e) => { c.on = e.target.checked; drawChain(); };
    host.querySelector('#wzChainReassign')?.addEventListener('change', (e) => { c.reassign = e.target.checked; });
    host.querySelector('#wzChainOverride')?.addEventListener('change', (e) => { c.override = e.target.checked; });
    host.querySelector('#wzAddStage')?.addEventListener('click', () => { c.stages.push({ name: ['Manager', 'Finance', 'Director', 'Legal'].find((n) => !c.stages.some((x) => x.name === n)) || 'Stage ' + (c.stages.length + 1), approver: 'column', column: '', groupId: '', notify: '' }); drawChain(); });
    host.querySelectorAll('[data-del]').forEach((b) => b.onclick = () => { c.stages.splice(+b.dataset.del, 1); drawChain(); });
    host.querySelectorAll('[data-st]').forEach((row) => {
      const st = c.stages[+row.dataset.st];
      row.querySelectorAll('input[data-k]').forEach((x) => x.oninput = () => { st[x.dataset.k] = x.value.trim(); });
      row.querySelector('select[data-k=who]').onchange = (e) => { c.touched = true; const v = e.target.value; if (v === 'group') st.approver = 'group'; else { st.approver = 'column'; st.column = v.startsWith('col:') ? v.slice(4) : ''; } drawChain(); };
    });
  }
  m.el.querySelector('#wzBack').onclick = () => { wz.step--; draw(); };
  m.el.querySelector('#wzBlank')?.addEventListener('click', () => { m.close(); });
  m.el.querySelector('#wzDemo')?.addEventListener('click', () => { m.close(); openDemo(); });
  m.el.querySelector('#wzNext').onclick = async () => {
    if (wz.step === 2) {
      const names = wz.lists.map((l) => l.name);
      if (!wz.lists.length) { toast('Add at least one list', true); return; }
      if (new Set(names).size !== names.length) { toast('Two lists share a name', true); return; }
    }
    if (wz.step < 3) { wz.step++; draw(); return; }
    const hasWork = P && P.screens.some((s) => s.controls.length);
    if (hasWork && !(await ask('Replace the current app?', 'Generating replaces your current screens and lists. Save a project file first if you want to keep them.', 'Replace'))) return;
    const p = core.newProject();
    p.app.name = wz.name || 'My App'; p.app.formFactor = wz.ff;
    p.theme = clone(THEME_PRESETS.find((t) => t.name === wz.theme) || THEME_PRESETS[0]);
    p.lists = clone(wz.lists);
    const apprList = body.querySelector('#wzApprList')?.value;
    const sc = SCENARIOS[wz.scen];
    const approvals = wz.approvals && apprList ? (sc.approvals?.list === apprList ? sc.approvals : { list: apprList }) : null;
    if (wz.ft.adminGroup && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(wz.ft.adminGroup)) { toast('The admin group ID should look like 00000000-0000-0000-0000-000000000000 (Entra → Groups → Object ID)', true); return; }
    p.app.features = { ...wz.ft };
    let chain = null;
    if (approvals && wz.chain.on) {
      const st = wz.chain.stages;
      if (!st.length || st.some((x) => !x.name.trim())) { toast('Give every approval stage a name', true); return; }
      if (new Set(st.map((x) => stageCols(x).K)).size !== st.length) { toast('Two approval stages share a name', true); return; }
      const bad = st.find((x) => x.approver === 'group' && x.groupId && !isGuid(x.groupId));
      if (bad) { toast(`The ${bad.name} group ID should look like 00000000-0000-0000-0000-000000000000 (Entra → Groups → Object ID)`, true); return; }
      chain = { list: approvals.list, stages: st.map((x) => ({ ...x })), reassign: !!wz.chain.reassign, override: !!wz.chain.override };
    }
    generateApp(p, { screens: [...wz.screens], approvals, chain });
    if (p.app.chain) addChainFlows(p);
    if (wz.flows !== false && sc.flows) addScenarioFlows(p, sc);
    if (!p.screens.length) p.screens.push(newScreen(p, 'scrHome'));
    m.close();
    setProject(p, true); commit();
    toast(`Generated ${p.screens.length} screens. Press Play to try it.`);
    afterBuild();
  };
  draw();
}

// code / export
function codeDialog(tab) {
  if (tab) S.codeTab = tab;
  let F2 = lint(P);
  const errs = F2.filter((x) => x.level === 'error').length, warns = F2.filter((x) => x.level === 'warn').length;
  const items = [
    ['grp', 'Paste into Power Apps'], ['deploy', 'Paste assistant'], ['steps', 'All steps as text'], ['formulas', 'App.Formulas'], ...(onStartFx(P) ? [['onstart', 'App.OnStart']] : []), ['startscreen', 'App.StartScreen'],
    ...P.screens.map((s) => ['scr:' + s.id, s.name]),
    ...((P.flows || []).filter((f) => f.enabled !== false).length ? [['grp', 'Flows (Import Package, Legacy)'], ...P.flows.filter((f) => f.enabled !== false).map((f) => ['flow:' + f.id, f.name])] : []),
    ['grp', 'Check'], ['lint', 'Lint report'],
    ['grp', 'Data and project'], ['script', 'Provisioning script'], ['project', 'Project file'],
  ];
  const m = modal(`<div class="modal-h"><h2>Get YAML</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b"><div class="code-layout"><nav class="code-nav" id="codeNav">${items.map(([k, l]) => k === 'grp' ? `<div class="grp">${l}</div>` : `<button data-ct="${k}">${k.startsWith('scr:') ? ui('screen') : ''}${esc(l)}${k === 'lint' ? `<span class="cnt ${errs ? '' : warns ? 'warn' : 'ok'}">${errs || warns || '✓'}</span>` : ''}</button>`).join('')}</nav><div class="code-main" id="codeMain"></div></div></div>
    <div class="modal-f"><span class="grow">${errs ? `<b style="color:var(--danger)">${errs} error${errs > 1 ? 's' : ''}</b> to fix before pasting. ` : ''}Paste screens one at a time with Ctrl+V on an empty canvas in Studio.</span><button class="btn" id="guideBtn">${ui('file')}Export guide</button><button class="btn" id="zipAll">${ui('download')}Download all (.zip)</button><button class="btn pri" data-close>Done</button></div>`);
  const main = m.el.querySelector('#codeMain');
  const show = (k) => {
    S.codeTab = k;
    m.el.querySelectorAll('[data-ct]').forEach((b) => b.classList.toggle('on', b.dataset.ct === k));
    let title = '', text = '', fname = '', extra = '';
    if (k === 'deploy') { deployView(main, m.close); return; }
    if (k === 'steps') {
      main.innerHTML = `<div class="bar"><h3>Deploy steps</h3></div><ol class="checklist">${checklist(P).map((s) => `<li>${esc(s)}</li>`).join('')}</ol>
        <p class="note">Re-pasting a screen deletes controls you added by hand in Studio. Copy those first (select, Ctrl+C, paste into Notepad) and paste them back afterwards, or import the screen here so the designer keeps them.</p>`;
      return;
    }
    if (k === 'lint') {
      main.innerHTML = `<div class="bar"><h3>Lint report</h3><button class="btn sm" id="cReport" title="Paste the error Studio showed; see the control, cause and fix">Studio showed an error?</button></div>${F2.length ? F2.map((x, i) => `<div class="finding" data-fi="${i}"><span class="lv ${x.level}">${x.level.toUpperCase()}</span><div><b>${esc(x.where)} · ${esc(x.code)}</b>${esc(x.msg)}</div>${SETUP_CODES.has(x.code) && (x.code !== 'PS041' || /varIsAdmin/.test(x.msg)) ? '<button class="btn sm" data-fix="1">Fix in Finish setup</button>' : ''}${x.list && x.code === 'PS035' && P.lists.some((l) => l.name === x.list && l.size !== 'large') ? `<button class="btn sm" data-size="large" data-list="${esc(x.list)}" title="Counts show 2000+ and lint flags formulas that need reshaping">Mark ${esc(x.list)} large</button>` : ''}${x.list && x.code === 'PS059' && P.lists.some((l) => l.name === x.list && l.size === 'large') ? `<button class="btn sm" data-size="small" data-list="${esc(x.list)}" title="Use this when the list will stay under 2,000 items">${esc(x.list)} stays small</button>` : ''}</div>`).join('') : '<p>No problems found. Every rule the generator knows about passes.</p>'}
        <p class="note">Lint checks the paste rules learned from real tenants: unique names app-wide, no Navigate in OnStart/OnVisible, ThisItem only in galleries, valid bindings and targets, required columns covered by Save, clickable controls not covered. To silence one rule on one control, add a comment such as <code>// lint-ok PS040</code> to any of its formulas.</p>`;
      main.querySelector('#cReport').onclick = () => { m.close(); feedbackDialog(); };
      main.querySelectorAll('[data-fix]').forEach((b) => b.onclick = (e) => { e.stopPropagation(); m.close(); setupDialog(); });
      main.querySelectorAll('[data-size]').forEach((b) => b.onclick = (e) => {
        e.stopPropagation();
        const l = P.lists.find((x) => x.name === b.dataset.list); if (!l) return;
        if (b.dataset.size === 'large') l.size = 'large'; else delete l.size;
        commit(); renderAll(); toast(b.dataset.size === 'large' ? `${l.name} marked large. ${SIZES.large.note}` : `${l.name} is no longer marked large.`);
        F2 = lint(P); show('lint');
      });
      main.querySelectorAll('[data-fi]').forEach((d) => d.onclick = () => { const x = F2[+d.dataset.fi]; const nm = x.where.split(' › ')[1]; for (const s of P.screens) for (const [c] of walk(s.controls)) if (c.name === nm) { S.screenId = s.id; S.sel = [c.id]; m.close(); renderAll(); return; } });
      return;
    }
    if (k === 'formulas') { title = 'App.Formulas'; text = themeFormulas(P); fname = 'App.Formulas.txt'; extra = 'In Studio: Tree view → App → choose Formulas in the property list → paste.'; }
    if (k === 'onstart') { title = 'App.OnStart'; text = onStartFx(P); fname = 'App.OnStart.txt'; extra = 'Expand the formula bar before pasting, then App → Run OnStart.'; }
    if (k === 'startscreen') { title = 'App.StartScreen'; text = '=' + startScreenFx(P); fname = 'App.StartScreen.txt'; extra = 'Paste into App → StartScreen after the screens exist.'; }
    if (k === 'script') { title = 'Provisioning script (PnP PowerShell)'; text = provisioningScript(P); fname = core.slug(P.app.name) + '-Provision-Lists.ps1'; extra = 'Creates the lists and columns. Fill in the EDIT ME block. Safe to run again.'; }
    if (k === 'project') { title = 'Project file'; text = JSON.stringify(P, null, 2); fname = core.slug(P.app.name) + '.pastudio.json'; extra = 'Save this to continue later, or share it. Open it with Project → Open project file.'; }
    if (k.startsWith('flow:')) {
      const f = P.flows.find((x) => x.id === k.slice(5));
      main.innerHTML = `<div class="bar"><h3>${esc(f.name)}</h3><button class="btn sm" id="cEdit">Edit</button><button class="btn sm" id="cCopy">${ui('copy')}Copy definition</button><button class="btn sm pri" id="cDl">${ui('download')}Download package</button></div>
        <ol class="checklist" style="font-size:13px"><li>Power Automate → My flows → Import → <b>Import Package (Legacy)</b> → Upload <code>${esc(flowFileName(f))}</code>.</li><li>Import setup: <b>Create as new</b>. Related resources: select your SharePoint and Office 365 Outlook connections. Import.</li><li>Open the flow, check it, and turn it on.${f.kind === 'button' ? ` In Power Apps Studio add it from the Power Automate pane; it appears as <code>${esc(flowFxName(f.name))}</code>.` : ''}</li></ol>
        <pre class="code">${esc(JSON.stringify(flowDefinition(P, f), null, 2))}</pre>`;
      main.querySelector('#cCopy').onclick = () => copyText(JSON.stringify(flowDefinition(P, f), null, 2));
      main.querySelector('#cDl').onclick = () => downloadFlow(f);
      main.querySelector('#cEdit').onclick = () => { m.close(); flowDialog(f.id); };
      return;
    }
    if (k.startsWith('scr:')) { const s = P.screens.find((x) => x.id === k.slice(4)); title = s.name; text = screenYaml(P, s); fname = s.name + '.pa.yaml'; extra = `Delete the old ${s.name} in Studio first (names must be unique), click the empty canvas, press Ctrl+V.`; }
    main.innerHTML = `<div class="bar"><h3>${esc(title)}</h3><button class="btn sm" id="cDl">${ui('download')}Download</button><button class="btn sm pri" id="cCopy">${ui('copy')}Copy</button></div><p class="note" style="margin:0">${esc(extra)}</p><pre class="code">${esc(text)}</pre>`;
    main.querySelector('#cCopy').onclick = () => copyText(text);
    main.querySelector('#cDl').onclick = () => download(fname, text);
  };
  m.el.querySelector('#codeNav').addEventListener('click', (e) => { const b = e.target.closest('[data-ct]'); if (b) show(b.dataset.ct); });
  m.el.querySelector('#guideBtn').onclick = () => exportGuide();
  m.el.querySelector('#zipAll').onclick = async () => {
    if (!window.JSZip) { toast('Zip library did not load', true); return; }
    const z = new window.JSZip(), root = z.folder(core.slug(P.app.name));
    root.file('README.txt', 'Power App Studio export: ' + P.app.name + '\n\n' + checklist(P).map((s, i) => `${i + 1}. ${s}`).join('\n') + '\n');
    root.file('App.Formulas.txt', themeFormulas(P));
    if (onStartFx(P)) root.file('App.OnStart.txt', onStartFx(P));
    root.file('App.StartScreen.txt', '=' + startScreenFx(P));
    for (const f of (P.flows || []).filter((x) => x.enabled !== false)) root.file('flows/' + flowFileName(f), await flowZipBlob(f));
    P.screens.forEach((s, i) => root.file(`screens/${String(i + 1).padStart(2, '0')}-${s.name}.pa.yaml`, screenYaml(P, s)));
    root.file('Provision-Lists.ps1', provisioningScript(P));
    root.file('lint.txt', lint(P).map((x) => `${x.level.toUpperCase()}\t${x.code}\t${x.where}\t${x.msg}`).join('\n') || 'No findings.');
    root.file(core.slug(P.app.name) + '.pastudio.json', JSON.stringify(P, null, 2));
    download(core.slug(P.app.name) + '-powerapp.zip', await z.generateAsync({ type: 'blob' }));
  };
  const valid = items.some(([k]) => k === S.codeTab) ? S.codeTab : 'steps';
  show(valid);
}

function importDialog() {
  const m = modal(`<div class="modal-h"><h2>Import screen YAML</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b"><p class="note" style="margin-top:0">In Power Apps Studio, right-click a screen in the Tree view → <b>Copy</b> (or select controls and press Ctrl+C), then paste here. A <code>Src/*.pa.yaml</code> file from an exported .msapp works too.</p>
      <textarea class="inp" id="impText" rows="14" placeholder="Screens:&#10;  Screen1:&#10;    Properties: …"></textarea>
      <div class="btns" style="margin-top:8px"><button class="btn sm" id="impFile">${ui('upload')}Load a .yaml file</button></div><p class="note">Classic controls become editable. Anything else is kept exactly as pasted, so exporting again loses nothing. Names that clash with this app get a suffix and references are updated.</p></div>
    <div class="modal-f"><button class="btn" data-close>Cancel</button><button class="btn pri" id="impGo">Import</button></div>`, 'md');
  m.el.querySelector('#impFile').onclick = () => { const fi = $('#fileYaml'); fi.value = ''; fi.onchange = async () => { m.el.querySelector('#impText').value = await fi.files[0].text(); }; fi.click(); };
  m.el.querySelector('#impGo').onclick = () => {
    try {
      const r = importYaml(P, m.el.querySelector('#impText').value);
      if (!r.screens.length) throw new Error('No controls found.');
      P.screens.push(...r.screens); S.screenId = r.screens[0].id; S.sel = [];
      m.close(); renderAll(); zoomFit(); commit();
      const n = Object.keys(r.renamed).length;
      toast(`Imported ${r.screens.map((s) => s.name).join(', ')}${n ? `, ${n} name${n > 1 ? 's' : ''} renamed` : ''}`);
    } catch (e) { toast('Import failed: ' + e.message, true); }
  };
}
function schemaDialog() {
  askText('Import list schema', 'Paste JSON: {"lists":[{"name":"Requests","columns":[{"name":"Status","type":"Choice","choices":["New","Done"]}]}]} or SharePoint REST fields output', '', true).then((t) => {
    if (!t) return;
    try {
      const j = JSON.parse(t);
      const out = [];
      const mapType = (x) => ({ Text: 'Text', Note: 'Note', Number: 'Number', Currency: 'Currency', DateTime: 'DateTime', Choice: 'Choice', MultiChoice: 'Choice', Boolean: 'YesNo', YesNo: 'YesNo', User: 'Person', Person: 'Person', UserMulti: 'Person', Lookup: 'Lookup' }[x] || 'Text');
      const fromCols = (name, cols) => { const l = newList(name); for (const c of cols) { const nm = c.name || c.InternalName || c.Title; const ty = mapType(c.type || c.TypeAsString); if (!nm || nm === 'Title' || /^_|^(ID|ContentType|Modified|Created|Author|Editor|Attachments|Edit|LinkTitle|DocIcon|ItemChildCount|FolderChildCount|AppAuthor|AppEditor|ComplianceAssetId)$/.test(nm) || c.Hidden || c.ReadOnlyField) continue; l.columns.push(newColumn(String(nm).replace(/\s+/g, ''), ty, { required: !!(c.required || c.Required), choices: c.choices || c.Choices?.results || c.Choices || [], lookupList: c.lookupList || '' })); } return l; };
      if (j.lists) j.lists.forEach((l) => out.push(fromCols(l.name, l.columns || [])));
      else { const rows = j.d?.results || j.value || (Array.isArray(j) ? j : null); if (!rows) throw new Error('Unrecognized JSON'); out.push(fromCols('ImportedList', rows)); }
      for (const l of out) { let n = l.name, i = 2; while (P.lists.some((x) => x.name === n)) n = l.name + i++; l.name = n; P.lists.push(l); }
      renderAll(); commit(); toast(`Imported ${out.length} list${out.length > 1 ? 's' : ''}`);
    } catch (e) { toast('Could not read that JSON: ' + e.message, true); }
  });
}
function helpDialog() {
  modal(`<div class="modal-h"><h2>How Power App Studio works</h2><button class="mini" data-close>${ui('close')}</button></div><div class="modal-b">
    <ol class="checklist">
      <li><b>Describe it.</b> New app → pick a scenario, shape the SharePoint lists and columns, choose screens and a look. Every screen is generated and wired: browse → details → edit → save.</li>
      <li><b>Design it.</b> Drag controls on the canvas, drop them into a gallery's first row to repeat them, bind inputs to columns, add actions (navigate, save, approve, show/hide), and restyle everything from the Theme tab.</li>
      <li><b>Test it.</b> Play runs the app against sample data: search, open, edit and save items, approve them, toggle panels.</li>
      <li><b>Ship it.</b> Get YAML gives App.Formulas, OnStart and one YAML block per screen. In Power Apps Studio, paste each screen with Ctrl+V on an empty canvas. Add your SharePoint lists as data sources with the same names. The provisioning script can create the lists for you.</li>
      <li><b>Keep going.</b> Copy a screen out of Studio and Import it here to keep designing. Save the project file to continue on another machine.</li>
    </ol>
    <p class="note">Everything runs in your browser. Nothing is uploaded; your project is saved in this browser's storage and in the project files you download. The YAML follows rules field-tested in commercial and GCC High tenants: classic controls only, no version pins except the gallery, block scalars where YAML needs them, unique names across the app, theme in one named formula.</p>
  </div><div class="modal-f"><button class="btn pri" data-close>Got it</button></div>`, 'md');
}

// ------------------------------------------------------------------ play mode
function play(opts = {}) {
  const savedP = opts.project ? P : null;
  if (opts.project) { S.previewing = true; P = opts.project; } else commit();
  if (!P.screens.length) { if (savedP) { P = savedP; S.previewing = false; } return; }
  const bg = document.createElement('div'); bg.className = 'play-bg' + (opts.preview ? ' play-over' : '');
  const viewer = !!opts.viewer;
  bg.innerHTML = `${viewer ? `<div class="viewer-banner" role="status"><b>Preview of ${esc(P.app.name)}</b><span>sample data · nothing you do here is saved or sent</span><span class="grow"></span><button class="tb primary" id="plCopy">Open in Studio as a copy</button></div>` : ''}<div class="play-bar"><b>▶ ${esc(P.app.name)}</b><select id="plScr" aria-label="Jump to screen"></select><button class="tb" id="plRestart">Restart</button>${P.app.features?.adminGroup ? '<label class="pl-admin"><input type="checkbox" id="plAdmin" checked>Signed in as an admin</label>' : ''}${(P.app.chain?.stages || []).some((st) => st.approver === 'group') ? '<label class="pl-admin"><input type="checkbox" id="plGroups" checked>In the approver groups</label>' : ''}<span class="grow">${viewer ? '' : opts.preview ? 'Scenario preview · sample data · nothing here changes your app · Esc to close' : 'Sample data · changes are not saved · Esc to close'}</span>${viewer ? '' : '<button class="tb primary" id="plClose">Close</button>'}</div><div class="play-stage"><div class="play-frame" id="plFrame"><div class="screen" id="plScreen"></div></div></div>`;
  document.body.appendChild(bg);
  let rt, toastTimer, focusSave = null;
  const hooks = {
    rerender: (o) => draw(o),
    toast: (t, kind) => { void toastTimer; let box = bg.querySelector('.play-toasts'); if (!box) { box = document.createElement('div'); box.className = 'play-toasts'; bg.querySelector('#plFrame').appendChild(box); } const d = document.createElement('div'); d.className = 'play-toast ' + (kind || ''); d.textContent = t; box.appendChild(d); while (box.children.length > 4) box.firstChild.remove(); setTimeout(() => d.remove(), kind === 'flow' ? 4200 : 2400); },
  };
  let soon = 0;
  const env = () => ({ project: P, theme: P.theme, mode: 'play', rt, rerenderSoon: () => { clearTimeout(soon); soon = setTimeout(draw, 220); } });
  const fit = () => {
    const f = ff(), st = bg.querySelector('.play-stage');
    const z = Math.min((st.clientWidth - 40) / f.w, (st.clientHeight - 40) / f.h, 1.25);
    const fr = bg.querySelector('#plFrame'); fr.style.width = f.w * z + 'px'; fr.style.height = f.h * z + 'px';
    bg.querySelector('#plScreen').style.transform = `scale(${z})`;
  };
  // timers: run the current screen's auto-start timers (loading screens, polling), like the real app
  let timerScreen = null, timers = [];
  const stopTimers = () => { timers.forEach(clearInterval); timers = []; };
  const startTimers = (scr) => {
    stopTimers(); timerScreen = scr.id;
    for (const [c] of walk(scr.controls)) {
      if (c.type !== 'timer' || !c.props.autoStart) continue;
      const h = setInterval(() => {
        if (!rt || rt.current()?.id !== scr.id) return;
        if (c.fx?.Start && !tryFx(c.fx.Start, { project: P, rt }, true)) return;
        rt.run(scr, c.actions?.OnTimerEnd, null, { diff: true });
        if (!c.props.repeat) clearInterval(h);
      }, Math.max(100, Number(c.props.duration) || 1000));
      timers.push(h);
    }
  };
  function draw(opts = {}) {
    const a = document.activeElement, host = bg.querySelector('#plScreen');
    if (a && host.contains(a)) { const c = a.closest('.pa-ctl'); focusSave = c ? { id: c.dataset.id, s: a.selectionStart, e: a.selectionEnd } : null; } else focusSave = null;
    const scr = rt.current();
    renderScreen(host, scr, env(), { diff: !!opts.diff });
    if (timerScreen !== scr.id) startTimers(scr);
    bg.querySelector('#plScr').innerHTML = P.screens.map((s) => `<option value="${s.id}" ${s.id === scr.id ? 'selected' : ''}>${esc(s.name)}</option>`).join('');
    if (focusSave) { const el = host.querySelector(`[data-id="${focusSave.id}"] input, [data-id="${focusSave.id}"] textarea`); if (el) { el.focus(); try { el.setSelectionRange(focusSave.s, focusSave.e); } catch { /* date inputs */ } } }
    fit();
  }
  const start = () => { stopTimers(); timerScreen = null; rt = createRuntime(P, hooks); rt.hooks = hooks; const ad = bg.querySelector('#plAdmin'); if (ad) rt.vars.varIsAdmin = ad.checked; setGroups(); draw(); };
  const setGroups = () => { const g = bg.querySelector('#plGroups'); if (g) for (const st of P.app.chain?.stages || []) if (st.approver === 'group') rt.vars[stageCols(st).groupVar] = g.checked; };
  bg.querySelector('#plGroups')?.addEventListener('change', () => { setGroups(); draw(); });
  bg.querySelector('#plAdmin')?.addEventListener('change', (e) => { rt.vars.varIsAdmin = e.target.checked; draw(); });
  start();
  bg.querySelector('#plScr').onchange = (e) => { rt.api.navigate(e.target.value); draw(); };
  bg.querySelector('#plRestart').onclick = start;
  const close = () => { stopTimers(); bg.remove(); window.removeEventListener('resize', fit); document.removeEventListener('keydown', onKey, true); if (savedP) { P = savedP; S.previewing = false; } };
  const onKey = (e) => { if (e.key === 'Escape' && !viewer) { e.stopPropagation(); close(); } };
  bg.querySelector('#plClose')?.addEventListener('click', close);
  bg.querySelector('#plCopy')?.addEventListener('click', async () => {
    const saved = loadSaved();
    if (saved && (saved.screens || []).some((s) => (s.controls || []).length) && !(await ask('Replace your app in this browser?', `Opening a copy of ${P.app.name} replaces the app saved in this browser. Save a project file first if you want to keep it.`, 'Open copy'))) return;
    const copy = JSON.parse(JSON.stringify(P));
    close(); leaveViewer(copy);
  });
  window.addEventListener('resize', fit); document.addEventListener('keydown', onKey, true);
}

// ------------------------------------------------------------------ wiring
function onKey(e) {
  if ($('.play-bg')) return;
  const t = e.target, typing = t.closest && (t.closest('input,textarea,select,[contenteditable]'));
  const mod = e.ctrlKey || e.metaKey;
  if (e.key === 'Escape') { if ($('.modal-bg')) { $('.modal-bg:last-child').remove(); return; } $$('.colorpop').forEach((x) => x.remove()); S.sel = []; drawOverlay(); renderInspector(); renderPanelIf('layers'); return; }
  if (e.key === 'F5') { e.preventDefault(); play(); return; }
  if ($('.modal-bg')) return;
  if (mod && e.key.toLowerCase() === 'z' && !typing) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
  if (mod && e.key.toLowerCase() === 'y' && !typing) { e.preventDefault(); redo(); return; }
  if (typing) return;
  if (mod && e.key.toLowerCase() === 'c') { copySel(); return; }
  if (mod && e.key.toLowerCase() === 'v' && S.clip?.length) { e.preventDefault(); pasteCtls(S.clip); return; }
  if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); copySel(); pasteCtls(S.clip); return; }
  if (mod && e.key.toLowerCase() === 'a') { e.preventDefault(); S.sel = screen().controls.map((c) => c.id); drawOverlay(); renderInspector(); return; }
  // keyboard selection while the canvas has focus: Tab / Shift+Tab walk the controls, Enter picks the first
  if (t.id === 'viewport' && (e.key === 'Tab' || e.key === 'Enter') && !(e.key === 'Tab' && !S.sel.length)) {
    const flat = [...walk(screen().controls)].map(([c]) => c.id);
    if (!flat.length) return;
    e.preventDefault();
    const i = flat.indexOf(S.sel[S.sel.length - 1]);
    S.sel = [flat[e.key === 'Enter' && i < 0 ? 0 : (i + (e.shiftKey ? -1 : 1) + flat.length) % flat.length]];
    drawOverlay(); renderInspector(); renderPanelIf('layers');
    const f = one(); if (f) $('#viewport').setAttribute('aria-label', `Canvas. Selected ${CATALOG[f.ctl.type].label} ${f.ctl.name}. Tab for next control, arrows to move, Delete to remove, Escape to deselect.`);
    return;
  }
  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSel(); return; }
  const step = e.shiftKey ? 10 : 1;
  if (S.sel.length && e.key.startsWith('Arrow')) { e.preventDefault(); nudge(e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0, e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0); }
}
function runCmd(cmd, el) {
  const i = el?.dataset.i != null ? +el.dataset.i : -1;
  switch (cmd) {
    case 'newScreen': return newScreenDialog();
    case 'wizard': return wizard();
    case 'import': return importDialog();
    case 'code': return codeDialog();
    case 'help': return helpDialog();
    case 'addList': { let n = 'NewList', k = 2; while (P.lists.some((l) => l.name === n)) n = 'NewList' + k++; P.lists.push(newList(n)); renderPanel(); commit(); return; }
    case 'importSchema': return schemaDialog();
    case 'importTable': return importTableDialog();
    case 'drift': return driftDialog();
    case 'feedback': return feedbackDialog();
    case 'history': return historyDialog();
    case 'bundleExport': return exportBundle();
    case 'bundleImport': { const fi = $('#fileBundle'); fi.value = ''; fi.click(); return; }
    case 'showScript': return codeDialog('script');
    case 'setup': return setupDialog();
    case 'guide': exportGuide(); return;
    case 'demo': return ask('Open the demo?', 'This replaces the current project with the finished Request Tracker demo. Save a project file first if you want to keep your work.', 'Open demo').then((ok) => { if (ok) openDemo(); });
    case 'tour': return startTour();
    case 'describe': return describeDialog();
    case 'change': return changeDialog();
    case 'addFlow': return addFlowDialog();
    case 'editFlow': return flowDialog(el.dataset.fid);
    case 'dlFlow': { const f = P.flows.find((x) => x.id === el.dataset.fid); if (f) downloadFlow(f); return; }
    case 'dupFlow': return duplicateFlow(el.dataset.fid);
    case 'delFlow': return deleteFlow(el.dataset.fid);
    case 'flowSample': return loadFlowSample();
    case 'flowSampleClear': delete P.app.flowSample; commit(); renderPanel(); return;
    case 'scrUp': if (i > 0) { P.screens.splice(i - 1, 0, P.screens.splice(i, 1)[0]); renderPanel(); renderTabs(); commit(); } return;
    case 'scrDup': {
      const src = P.screens[i], c = clone(src); c.id = uid(); c.name = core.uniqueScreenName(P, src.name + 'Copy');
      const used = allNames(P); used.add(c.name);
      c.controls.forEach((x) => renameCopy(P, x, used));
      P.screens.splice(i + 1, 0, c); S.screenId = c.id; renderAll(); commit(); return;
    }
    case 'scrDel': {
      if (P.screens.length < 2) { toast('An app needs at least one screen', true); return; }
      const sc = P.screens[i];
      ask('Delete screen?', `Delete ${sc.name} and its ${sc.controls.length} controls? Buttons that go to it will show lint errors.`, 'Delete').then((ok) => {
        if (!ok) return; P.screens.splice(i, 1); if (P.app.startScreen === sc.id) P.app.startScreen = P.screens[0].id; if (S.screenId === sc.id) S.screenId = P.screens[0].id; S.sel = []; renderAll(); commit();
      });
      return;
    }
    case 'save': return download(core.slug(P.app.name) + '.pastudio.json', JSON.stringify(P, null, 2), 'application/json');
    case 'copyProject': return copyText(JSON.stringify(P, null, 2));
    case 'open': { const fi = $('#fileOpen'); fi.value = ''; fi.click(); return; }
    case 'share': return shareDialog();
    case 'paste': return askText('Paste project JSON', 'Project JSON', '', true).then((t) => { if (!t) return; try { setProject(JSON.parse(t)); toast('Project opened'); } catch (e) { toast('Not a valid project: ' + e.message, true); } });
    case 'reset': return ask('Start over?', 'This clears the current project from this browser. Save a project file first if you want to keep it.', 'Start over').then((ok) => { if (ok) { setProject(core.newProject()); wizard(true); } });
  }
}

// ------------------------------------------------------------------ app bundle (.zip in the powerapps-canvas-yaml skill's layout)
async function exportBundle() {
  if (!window.JSZip) { toast('Zip library did not load', true); return; }
  try { const { name, blob } = await bundleZip(P, window.JSZip); download(name, blob); }
  catch (e) { toast('Could not build the bundle: ' + e.message, true); }
}
async function importBundle(file) {
  if (!window.JSZip) { toast('Zip library did not load', true); return; }
  let r;
  try { r = await readBundle(new Uint8Array(await file.arrayBuffer()), window.JSZip, core.newProject); }
  catch (e) { toast('Not an app bundle: ' + e.message, true); return; }
  const choice = isBlank(P) ? 'replace' : await new Promise((res) => {
    const m = modal(`<div class="modal-h"><h2>Open this app bundle?</h2></div><div class="modal-b"><p style="margin:0">${esc(file.name)} holds "${esc(r.project.app.name)}" (${r.project.screens.length} screens, ${r.project.lists.length} lists, ${(r.project.flows || []).length} flows).${r.notes.length ? ' ' + esc(r.notes.join(' ')) : ''}</p>
      <p class="note">Replace: the bundle becomes the open app. Add: its lists, screens and flows join the app you have open; clashing names get a number. Either way the current app is saved as a version first, and Undo brings it back.</p></div>
      <div class="modal-f"><button class="btn" data-c="">Cancel</button><span class="grow"></span><button class="btn" data-c="merge">Add to my app</button><button class="btn pri" data-c="replace">Replace my app</button></div>`, 'sm');
    m.el.querySelectorAll('[data-c]').forEach((b) => b.onclick = () => { m.close(); res(b.dataset.c); });
  });
  if (!choice) return;
  if (choice === 'merge') {
    const next = clone(P), res = mergeProjects(next, r.project);
    replaceProject(next, 'Before adding a bundle');
    renderAll();
    toast(`Added ${res.screens.length} screens, ${res.lists.length} lists and ${res.flows.length} flows.${res.renamed.length ? ` Renamed: ${res.renamed.slice(0, 3).join(', ')}${res.renamed.length > 3 ? '…' : ''}.` : ''} ${res.notes.join(' ')}`);
    return;
  }
  replaceProject(r.project, 'Before importing a bundle', { mergeVersions: true });
  S.screenId = P.screens.find((s) => s.id === P.app.startScreen)?.id || P.screens[0]?.id; S.sel = []; renderAll(); zoomFit();
  toast(r.source === 'embedded' ? 'Bundle opened' : 'Bundle opened: rebuilt from the canonical screens and schema');
}

function init() {
  initFlowsUI({ P: () => P, modal, toast, ui, esc, svgIcon, commit, renderPanel, download, copyText, ask });
  initDescribeUI({ P: () => P, modal, toast, ui, esc, ask, changeDialog: () => changeDialog(), setProject: (p) => { replaceProject(p, 'Before Describe → Build'); S.screenId = P.screens[0]?.id; S.sel = []; renderAll(); zoomFit(); afterBuild(); } });
  initChangeUI({ P: () => P, modal, toast, ui, esc, replaceProject: (p) => replaceProject(p, 'Before Ask for a change'), undo });
  initDriftUI({ P: () => P, modal, toast, ui, esc, replaceProject: (p) => replaceProject(p, 'Before importing Studio controls') });
  initFeedbackUI({ P: () => P, modal, toast, ui, esc, commit, persist, copyText, download, slug: core.slug,
    selectControl: (screenId, id) => { S.screenId = screenId; S.sel = [id]; renderAll(); },
    openScreen: (screenId) => { S.screenId = screenId; S.sel = []; renderAll(); },
    openDataTab: (list) => { S.tab = 'data'; renderRail(); renderPanel(); const i = P.lists.findIndex((l) => l.name === list); $(`#panel [data-li="${i}"]`)?.scrollIntoView({ block: 'start' }); } });
  initHistoryUI({ P: () => P, modal, toast, ui, esc, ask, persist, replaceProject });
  initTableUI({ P: () => P, modal, toast, ui, esc, commit, renderAll: () => { if (!P.screens.find((s) => s.id === S.screenId)) S.screenId = P.screens[0]?.id; renderAll(); } });
  initAssistUI({ P: () => P, modal, toast, ui, esc, commit, copyText, download, downloadFlow, renderThumb, slug: core.slug, reportError: () => feedbackDialog(),
    onSetupChange: () => { lintSoon(); updateSetupBadge(); renderPanelIf('flows'); renderPanelIf('app'); } });
  $('#btnSetup').onclick = () => setupDialog();
  $('#condChip').onclick = () => { S.showCond = !S.showCond; renderCanvas(); };
  // icons in static buttons
  $('#btnUndo').innerHTML = ui('undo'); $('#btnRedo').innerHTML = ui('redo');
  if (isShareHash(location.hash)) { P = core.newProject(); S.viewer = true; openShared(location.hash); }
  else {
    const saved = loadSaved();
    setProject(saved || core.newProject());
    attachVersions(!!saved);
    if (!saved) wizard(true);
  }

  $('#appName').addEventListener('input', (e) => { P.app.name = e.target.value; commitSoon(); renderPanelIf('app'); });
  $('#btnUndo').onclick = undo; $('#btnRedo').onclick = redo;
  $('#btnWizard').onclick = () => wizard();
  $('#btnDescribe').onclick = () => describeDialog();
  $('#btnChange').onclick = () => changeDialog();
  $('#btnPlay').onclick = play;
  $('#btnCode').onclick = () => codeDialog();
  $('#btnReport').onclick = () => feedbackDialog();
  $('#fileBundle').onchange = (e) => { const f = e.target.files[0]; if (f) importBundle(f); };
  const fm = $('#fileMenu');
  $('#btnFile').onclick = (e) => { e.stopPropagation(); fm.classList.toggle('open'); };
  document.addEventListener('click', () => fm.classList.remove('open'));
  fm.querySelector('.dd-menu').addEventListener('click', (e) => { const b = e.target.closest('[data-cmd]'); if (b) { fm.classList.remove('open'); runCmd(b.dataset.cmd, b); } });
  $('.rail').addEventListener('click', (e) => { const b = e.target.closest('[data-tab]'); if (!b) return; S.tab = b.dataset.tab; renderRail(); renderPanel(); });
  $('#screenTabs').addEventListener('click', (e) => { const b = e.target.closest('button'); if (!b) return; if (b.dataset.cmd) return runCmd(b.dataset.cmd, b); S.screenId = b.dataset.scr; S.sel = []; renderTabs(); renderCanvas(); renderInspector(); renderPanelIf('screens'); renderPanelIf('layers'); });
  $$('[data-zoom]').forEach((b) => b.onclick = () => { const z = zoomValue(); S.zoom = Math.max(0.2, Math.min(2, Math.round((z + (b.dataset.zoom === '+' ? 0.1 : -0.1)) * 10) / 10)); applyZoom(); drawOverlay(); });
  $('#zoomVal').onclick = zoomFit;
  $('#showHidden').onchange = (e) => { S.showHidden = e.target.checked; renderCanvas(); };
  window.addEventListener('resize', () => { if (S.zoom === 'fit') { applyZoom(); drawOverlay(); } });

  // canvas
  $('#screen').addEventListener('pointerdown', onScreenDown);
  $('#overlay').addEventListener('pointerdown', onScreenDown);
  $('#viewport').addEventListener('pointerdown', (e) => { if (e.target === $('#viewport') || e.target === $('#board')) { S.sel = []; drawOverlay(); renderInspector(); renderPanelIf('layers'); } });
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
  $('#viewport').addEventListener('dragover', onDragOver);
  $('#viewport').addEventListener('drop', onDrop);
  $('#viewport').addEventListener('dragleave', () => drawOverlay());
  $('#screen').addEventListener('dblclick', (e) => { const el = e.target.closest('.pa-ctl'); if (!el) return; const ta = $('#inspector textarea[data-p="props.text"]'); if (ta) { ta.focus(); ta.select(); } });

  // panel
  const panel = $('#panel');
  panel.addEventListener('dragstart', (e) => { const b = e.target.closest('[data-ins],[data-block],[data-myblock]'); if (!b) return; e.dataTransfer.setData('text/pas', b.dataset.ins || (b.dataset.myblock ? 'myblock:' + b.dataset.myblock : 'block:' + b.dataset.block)); e.dataTransfer.effectAllowed = 'copy'; });
  panel.addEventListener('click', (e) => {
    const b = e.target.closest('button, .scrrow, .layer'); if (!b) return;
    if (b.dataset.ins) return insertControl(b.dataset.ins);
    if (b.dataset.block) return doBlock(b.dataset.block);
    if (b.dataset.myblock) return doMyBlock(b.dataset.myblock);
    if (b.dataset.mbcmd) return myBlockCmd(b.dataset.mbcmd, b.dataset.id);
    if (b.dataset.preset != null) { P.theme = clone(THEME_PRESETS[+b.dataset.preset]); renderPanel(); renderCanvas(); commit(); return; }
    if (b.dataset.cfix) { const [f, bg] = b.dataset.cfix.split('|'), fx = fixContrast(P.theme.colors, f, bg); const was = P.theme.colors[fx.key]; P.theme.colors[fx.key] = fx.value; P.theme.name = 'Custom'; renderPanel(); renderCanvas(); commit(); toast(`${fx.key} ${was} → ${fx.value}`); return; }
    if (b.dataset.lcmd) { S.sel = [b.dataset.id]; zMove(b.dataset.lcmd); e.stopPropagation(); return; }
    if (b.dataset.cmd) { e.stopPropagation(); return runCmd(b.dataset.cmd, b); }
    if (b.classList.contains('scrrow')) { S.screenId = b.dataset.scr; S.sel = []; renderAll(); return; }
    if (b.classList.contains('layer')) { const id = b.dataset.sel; S.sel = e.shiftKey ? (S.sel.includes(id) ? S.sel.filter((x) => x !== id) : [...S.sel, id]) : [id]; drawOverlay(); renderInspector(); renderPanel(); }
  });
  panel.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.pcmd === 'blockList') { S.blockList = t.value; return; }
    if (t.dataset.th) { const k = t.dataset.th; P.theme[k] = t.type === 'checkbox' ? t.checked : k === 'radius' ? Number(t.value) : t.value; P.theme.name = 'Custom'; if (k === 'radius') $('#radVal').textContent = t.value + 'px'; renderCanvas(); commitSoon(); return; }
    if (t.dataset.tc) { P.theme.colors[t.dataset.tc] = t.value.toUpperCase(); P.theme.name = 'Custom'; const hx = panel.querySelector(`[data-tch="${t.dataset.tc}"]`); if (hx) hx.value = t.value.toUpperCase(); renderCanvas(); refreshContrast(); commitSoon(); return; }
    if (t.dataset.tch) { if (!isHex(t.value)) return; P.theme.colors[t.dataset.tch] = t.value.toUpperCase(); P.theme.name = 'Custom'; const sw = panel.querySelector(`[data-tc="${t.dataset.tch}"]`); if (sw) sw.value = t.value.slice(0, 7); renderCanvas(); refreshContrast(); commitSoon(); return; }
    if (t.dataset.appft === 'a11y') { if (t.checked) P.app.features.a11y = true; else delete P.app.features.a11y; commit(); return; }
    if (t.dataset.app) {
      const k = t.dataset.app;
      if (k === 'deepLinkScreen' || k === 'deepLinkList') {
        const d = { screen: '', list: '', ...(P.app.deepLink || {}) };
        d[k === 'deepLinkScreen' ? 'screen' : 'list'] = t.value;
        if (k === 'deepLinkScreen' && t.value && !d.list) { const s = P.screens.find((x) => x.id === t.value); const b = s && [...walk(s.controls)].map(([c]) => c.bind?.list).find(Boolean); if (b) { d.list = b; const ls = panel.querySelector('[data-app="deepLinkList"]'); if (ls) ls.value = b; } }
        if (d.screen || d.list) P.app.deepLink = d; else delete P.app.deepLink;
        commitSoon(); renderTabs(); return;
      }
      P.app[k] = t.value;
      if (k === 'name') updateTop();
      if (k === 'startScreen') renderTabs();
      if (k === 'formFactor') { renderCanvas(); zoomFit(); toast('Format changed. Existing controls keep their positions; generate again or rearrange them.'); }
      commitSoon();
    }
  });

  // inspector
  const insp = $('#inspector');
  insp.addEventListener('input', onInspectorInput);
  insp.addEventListener('change', onInspectorInput);
  insp.addEventListener('click', onInspectorClick);

  // files
  $('#fileBlocks').onchange = (e) => { const f = e.target.files[0]; if (f) importBlocksFile(f); };
  $('#fileOpen').onchange = async (e) => { const f = e.target.files[0]; if (!f) return; try { setProject(JSON.parse(await f.text())); toast('Opened ' + f.name); } catch (err) { toast('Not a valid project file: ' + err.message, true); } };
  $('#fileImg').onchange = async (e) => {
    const f = e.target.files[0], c = one(); if (!f || !c) return;
    if (f.size > 400000) toast('Large image: the screen YAML will be big. Consider a smaller file.', true);
    const r = new FileReader(); r.onload = () => { c.ctl.props.src = r.result; renderCanvas(); renderInspector(); commit(); }; r.readAsDataURL(f);
  };
  document.addEventListener('keydown', onKey);
  $('#btnSmallContinue').onclick = () => $('#smallScreen').classList.add('dismissed');
  window.addEventListener('beforeunload', () => { commit(); });
}
init();
