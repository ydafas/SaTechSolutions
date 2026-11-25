// app.js — SPFx Studio: describe a web part → Claude writes it → live preview with SharePoint stand-ins and the
// real property pane → download a deployable .sppkg (or a standard SPFx source project).
import { EXAMPLES } from './examples.js';
import { buildWebPart, names, componentVersion, GROUPS, uuid, hash20 } from './bundle.js';
import { buildSppkg } from './sppkg.js';
import { sourceFiles } from './source.js';
import { previewDoc } from './preview.js';
import { MODELS, buildRequest, changeHistory, fixHistory, callClaude, cleanSpec, checkCode, AiError } from './ai.js';
import { store, KEY, MODEL } from '../../powerapp-studio/js/keystore.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const clone = (o) => JSON.parse(JSON.stringify(o));
const LS_PROJECT = 'sss.project', LS_HISTORY = 'sss.history', LS_DRAFT = 'sss.draft', LS_VIEW = 'sss.view';

const THEMES = {
  Blue: ['#0f6cbd', '#0c3b5e', '#deecf9', '#eff6fc'], Teal: ['#03787c', '#014446', '#c5e9ea', '#f0f9fa'], Purple: ['#5c2e91', '#32174d', '#e1d5ec', '#f7f4fb'],
  Red: ['#a4262c', '#5c1418', '#f4d0d2', '#fcf4f5'], Navy: ['#1f3864', '#0f1c33', '#d6dce8', '#f3f5f9'], Green: ['#107c10', '#054b16', '#cfe8cf', '#f3faf3'],
};
const SAMPLE_REQUESTS = [
  'A staff directory from a list: photo, name, title, office, phone and email on cards, with a search box and an office filter.',
  'A training sign-up form: pick a course from a Courses list, choose a session date, add a note. Save to a Registrations list with the current user, and show "My registrations" below the form.',
  'An events calendar web part: month view with events from a list, colour-coded by category, click an event to see details. Also a compact "next 5 events" layout option.',
  'A policy library: documents grouped by category with search, a "Recently updated" badge, and an Acknowledge button that records the user in an Acknowledgements list.',
];
const CHANGE_CHIPS = ['Make it more compact', 'Add a search box', 'Add an "Add item" form', 'Add paging after 10 items', 'Make the cards larger with images'];

// ------------------------------------------------------------------ state
const S = {
  project: load(LS_PROJECT, null),
  view: Object.assign({ tab: 'build', codeTab: 'code', width: 'full', section: 'neutral', theme: 'Blue', edit: true, log: false }, load(LS_VIEW, {})),
  busy: null, ready: false, pending: null, rest: [], error: null, warnings: [], notes: [], fieldErrors: {}, pane: null, closedGroups: new Set(), codeErr: '', paneDraft: {},
};
function load(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } }
function saveLS(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { toast('Browser storage is full or blocked. Save a project file to keep your work.', true); } }
const save = () => { if (S.project) saveLS(LS_PROJECT, S.project); else localStorage.removeItem(LS_PROJECT); };
const saveView = () => saveLS(LS_VIEW, S.view);
function snapshot(label) {
  if (!S.project) return;
  const h = load(LS_HISTORY, []);
  h.unshift({ at: new Date().toISOString(), label, project: S.project });
  saveLS(LS_HISTORY, h.slice(0, 25));
}

function newProject(spec, request = '') {
  const ids = { solution: uuid(), feature: uuid(), component: uuid(), assets: uuid() };
  const p = { v: 1, spec: clone(spec), ids, build: 1, props: {}, request, downloaded: '' };
  p.props = defaultProps(p.spec);
  return p;
}
function defaultProps(spec, keep = {}) {
  const out = {};
  for (const p of spec.properties || []) {
    if (p.key in keep) { out[p.key] = keep[p.key]; continue; }
    out[p.key] = p.type === 'toggle' ? /^(true|yes|1)$/i.test(p.default) : (p.type === 'slider' || p.type === 'number') ? (p.default === '' ? '' : Number(p.default)) : p.default;
  }
  return out;
}
/** Replace the spec (after Claude, an example or a code edit); keep property values the editor already set. */
function setSpec(spec, label, { request } = {}) {
  S.project.spec = spec;
  S.project.props = defaultProps(spec, S.project.props || {});
  if (request !== undefined) S.project.request = request;
  save(); snapshot(label); renderAll(); reloadPreview(true);
}
const specHash = (p) => hash20(JSON.stringify(p.spec));

// ------------------------------------------------------------------ toasts, modals
function toast(msg, bad) {
  const t = document.createElement('div'); t.className = 'toast' + (bad ? ' bad' : ''); t.textContent = msg;
  $('#toasts').appendChild(t); setTimeout(() => t.remove(), bad ? 6000 : 3200);
}
function modal(title, body, foot = '<button class="btn pri" data-close>Done</button>') {
  const bg = document.createElement('div'); bg.className = 'modal-bg';
  bg.innerHTML = `<div class="modal" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="modal-h"><h2>${esc(title)}</h2><button class="x" data-close aria-label="Close">×</button></div><div class="modal-b">${body}</div><div class="modal-f">${foot}</div></div>`;
  const close = () => { bg.remove(); document.removeEventListener('keydown', onKey); };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  bg.addEventListener('click', (e) => { if (e.target === bg || e.target.closest('[data-close]')) close(); });
  document.addEventListener('keydown', onKey);
  document.body.appendChild(bg);
  return { el: bg, close };
}
function download(name, data, type = 'application/octet-stream') {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// ------------------------------------------------------------------ preview bridge
const frame = $('#preview');
function post(msg) { if (frame.contentWindow) frame.contentWindow.postMessage({ __sss: 1, ...msg }, '*'); }
function themeObj() {
  const [p, d, l, la] = THEMES[S.view.theme] || THEMES.Blue;
  const palette = { themePrimary: p, themeDark: d, themeDarker: d, themeLighter: l, themeLighterAlt: la, neutralPrimary: '#323130', neutralSecondary: '#605e5c', neutralLight: '#edebe9', neutralLighterAlt: '#faf9f8', white: '#ffffff' };
  if (S.view.section === 'soft') return { palette, semanticColors: { bodyBackground: la, bodyText: '#323130', bodySubtext: '#605e5c', bodyDivider: l, bodyStandoutBackground: '#ffffff', link: p }, isInverted: false };
  if (S.view.section === 'strong') return { palette: { ...palette, themePrimary: '#ffffff', themeDark: d, themeLighter: 'rgba(255,255,255,.25)', themeLighterAlt: 'rgba(255,255,255,.12)' }, semanticColors: { bodyBackground: p, bodyText: '#ffffff', bodySubtext: 'rgba(255,255,255,.82)', bodyDivider: 'rgba(255,255,255,.28)', bodyStandoutBackground: 'rgba(255,255,255,.1)', link: '#ffffff' }, isInverted: true };
  return { palette, semanticColors: { bodyBackground: '#ffffff', bodyText: '#323130', bodySubtext: '#605e5c', bodyDivider: '#edebe9', bodyStandoutBackground: '#faf9f8', link: p }, isInverted: false };
}
function reloadPreview(reseed = false) {
  if (!S.project) return;
  let bundle;
  try { bundle = buildWebPart(S.project).text; } catch (e) { showError({ message: 'Could not build the bundle: ' + e.message }); return; }
  const syntax = checkCode(S.project.spec.code);
  if (syntax) { showError({ message: 'The code has a syntax error: ' + syntax, syntax: true }); S.pane = null; renderPane(); return; }
  S.error = null; S.warnings = []; S.fieldErrors = {}; renderErr();
  const msg = { type: 'load', bundle, props: S.project.props, edit: S.view.edit, lists: S.project.spec.lists, theme: themeObj(), reseed };
  if (S.ready) post(msg); else S.pending = msg;
}
window.addEventListener('message', (e) => {
  if (e.source !== frame.contentWindow || !e.data || !e.data.__sss) return;
  const m = e.data;
  if (m.type === 'ready') { S.ready = true; if (S.pending) { post(S.pending); S.pending = null; } }
  else if (m.type === 'pane') { S.pane = m.pages; renderPane(); }
  else if (m.type === 'error') showError(m);
  else if (m.type === 'warn') { if (!S.warnings.includes(m.message)) { S.warnings.push(m.message); renderErr(); } }
  else if (m.type === 'mounted') { if (S.error && !S.error.syntax) { S.error = null; renderErr(); } }
  else if (m.type === 'rest') { S.rest.push(m); if (S.rest.length > 300) S.rest.shift(); renderLog(); }
  else if (m.type === 'fieldError') { S.fieldErrors[m.key] = m.message; const el = document.querySelector(`[data-ferr="${CSS.escape(m.key)}"]`); if (el) el.textContent = m.message; }
  else if (m.type === 'height') frame.style.height = Math.max(160, Math.min(m.height, 4000)) + 'px';
  else if (m.type === 'link') toast(`Link: ${m.href}${m.newTab ? ' (new tab)' : ''}. Links don't open from the preview.`);
});
function showError(m) { S.error = m; renderErr(); }
function renderErr() {
  const bar = $('#errBar');
  if (!S.error && !S.warnings.length) { bar.hidden = true; return; }
  bar.hidden = false;
  const e = S.error;
  bar.innerHTML = `<div class="grow">${e ? `<b>The web part threw an error.</b> ${esc(e.message)}${e.stack ? `<pre>${esc(e.stack)}</pre>` : ''}` : ''}${S.warnings.map((w) => `<div>⚠ ${esc(w)}</div>`).join('')}</div>
    ${e ? `<button class="btn sm pri" id="btnFix">Ask Claude to fix it</button>` : ''}<button class="btn sm" id="btnErrX">Dismiss</button>`;
  $('#btnErrX', bar).onclick = () => { S.error = null; S.warnings = []; renderErr(); };
  if (e) $('#btnFix', bar).onclick = () => runClaude('fix', `${e.message}${e.stack ? '\n' + e.stack : ''}`);
}
function renderLog() {
  $('#logCount').textContent = S.rest.length;
  const box = $('#restLog'); box.hidden = !S.view.log;
  if (!S.view.log) return;
  box.innerHTML = S.rest.slice(-120).map((r) => `<div class="s${String(r.status)[0]}">${esc(r.verb.padEnd(6))} ${r.status} ${esc(decodeURIComponent(r.url))}${r.error ? ` <em>${esc(r.error)}</em>` : ''}</div>`).join('') || '<div>No calls yet.</div>';
  box.scrollTop = box.scrollHeight;
}

// ------------------------------------------------------------------ property pane
function renderPane() {
  const el = $('#pane');
  if (!S.project) { el.innerHTML = '<div class="pp-empty">The property pane appears here once you have a web part.</div>'; return; }
  const s = S.project.spec;
  const pages = S.pane || [];
  let html = `<div class="pp-h"><h2>${esc(s.title)}</h2><p>${esc(pages[0]?.header || s.description || '')}</p></div>`;
  if (!pages.length) html += '<div class="pp-empty">Waiting for the preview…</div>';
  for (const pg of pages) for (const g of pg.groups) {
    const closed = S.closedGroups.has(g.name);
    html += `<div class="pp-g${closed ? ' closed' : ''}"><button type="button" data-grp="${esc(g.name)}" aria-expanded="${!closed}">${esc(g.name)}</button><div class="pp-fields">${g.fields.map(fieldHtml).join('')}</div></div>`;
  }
  el.innerHTML = html;
}
function fieldHtml(f) {
  const v = f.key && f.key in S.project.props ? S.project.props[f.key] : f.value;
  const k = esc(f.key || ''), lbl = esc(f.label || '');
  const desc = f.description ? `<small>${esc(f.description)}</small>` : '';
  const err = `<span class="err" data-ferr="${k}">${esc(S.fieldErrors[f.key] || '')}</span>`;
  switch (f.type) {
    case 3: return `<label class="pp-f"><span>${lbl}</span>${f.multiline ? `<textarea data-pk="${k}">${esc(v ?? '')}</textarea>` : `<input type="text" data-pk="${k}" value="${esc(v ?? '')}">`}${desc}${err}</label>`;
    case 5: return `<div class="pp-f"><span>${lbl}</span><div class="pp-tog"><button type="button" role="switch" data-pt="${k}" aria-checked="${!!v}" aria-label="${lbl}"></button><span>${v ? esc(f.onText || 'On') : esc(f.offText || 'Off')}</span></div></div>`;
    case 6: return `<label class="pp-f"><span>${lbl}</span><select data-pk="${k}">${(f.options || []).map((o) => `<option value="${esc(o.key)}"${String(o.key) === String(v) ? ' selected' : ''}>${esc(o.text)}</option>`).join('')}</select>${desc}</label>`;
    case 8: return `<label class="pp-f"><span>${lbl}</span><div class="pp-sl"><input type="range" data-pk="${k}" data-num="1" min="${f.min ?? 0}" max="${f.max ?? 100}" step="${f.step || 1}" value="${esc(v ?? f.min ?? 0)}"><output>${esc(v ?? '')}</output></div></label>`;
    case 11: return `<div class="pp-f"><button type="button" class="pp-btn" data-pb="${k}">+ ${esc(f.text || f.label)}</button></div>`;
    case 7: return `<p class="pp-lbl">${esc(f.text || f.label)}</p>`;
    default: return '';
  }
}
let propTimer = 0;
function setProp(key, value) {
  S.project.props[key] = value;
  post({ type: 'prop', key, value });
  clearTimeout(propTimer); propTimer = setTimeout(save, 400);
}
$('#pane').addEventListener('input', (e) => {
  const t = e.target; if (!t.dataset.pk) return;
  const val = t.dataset.num ? Number(t.value) : t.value;
  if (t.type === 'range') t.nextElementSibling.textContent = t.value;
  setProp(t.dataset.pk, val);
});
$('#pane').addEventListener('click', (e) => {
  const g = e.target.closest('[data-grp]');
  if (g) { const n = g.dataset.grp; if (S.closedGroups.has(n)) S.closedGroups.delete(n); else S.closedGroups.add(n); renderPane(); return; }
  const tg = e.target.closest('[data-pt]');
  if (tg) { const v = tg.getAttribute('aria-checked') !== 'true'; tg.setAttribute('aria-checked', String(v)); tg.nextElementSibling.textContent = v ? 'On' : 'Off'; setProp(tg.dataset.pt, v); return; }
  const b = e.target.closest('[data-pb]');
  if (b) post({ type: 'button', key: b.dataset.pb });
});

// ------------------------------------------------------------------ left: Build tab
function renderBuild() {
  const el = $('#tab-build');
  const key = store.get(KEY), model = store.get(MODEL) || MODELS[0][0];
  const has = !!S.project;
  const draft = localStorage.getItem(LS_DRAFT) || '';
  if (S.busy) {
    el.innerHTML = `<div class="card"><div class="progress" id="prog">${esc(S.busy.label)}</div><p class="hint" style="margin-top:8px">A full web part usually takes 40–120 seconds. You can keep this tab in the background.</p><button class="btn sm" id="btnCancel" style="margin-top:10px">Cancel</button></div>`;
    $('#btnCancel', el).onclick = () => S.busy?.ctrl.abort();
    return;
  }
  el.innerHTML = `
    ${has ? `<div><h2 class="sec">Ask for a change</h2><p class="hint">Describe what to change. Claude edits the web part you have and keeps the rest.</p></div>
      <textarea class="ta" id="changeText" placeholder="e.g. Show the office as a coloured chip and sort people by last name"></textarea>
      <div class="chips">${CHANGE_CHIPS.map((c) => `<button class="chip" data-chg="${esc(c)}">${esc(c)}</button>`).join('')}</div>
      <button class="btn pri block" id="btnChange">Apply change with Claude</button>
      ${S.project.spec.notes ? `<div class="notes"><b>Notes from Claude</b>${esc(S.project.spec.notes)}</div>` : ''}
      ${S.notes.length ? `<div class="warnbox">${S.notes.map(esc).join('<br>')}</div>` : ''}
      <hr style="border:0;border-top:1px solid var(--line);width:100%">` : ''}
    <div><h2 class="sec">${has ? 'Start a new web part' : 'Describe your web part'}</h2><p class="hint">What it shows, where the data lives, what editors can configure, and how it should look. Claude writes the web part; the preview runs it against sample data.</p></div>
    <textarea class="ta" id="descText" style="min-height:${has ? 110 : 170}px" placeholder="e.g. A staff directory from a list with photos, search and an office filter…">${esc(draft)}</textarea>
    <div class="chips">${SAMPLE_REQUESTS.map((r, i) => `<button class="chip" data-req="${i}">${esc(r.split(':')[0].split(' from ')[0].replace(/^An? /, ''))}</button>`).join('')}</div>
    <button class="btn ${has ? '' : 'pri'} block" id="btnGen">${has ? 'Generate a new web part' : 'Generate web part'}</button>
    <div class="card"><h3>Your Anthropic API key</h3>
      <p class="hint">Calls go straight from this browser to api.anthropic.com. The key is not saved in project files or sent anywhere else. Power App Studio on this site uses the same key. <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noopener">Get a key</a>.</p>
      <label class="f" style="margin-top:8px"><span class="sr">API key</span><input class="in" id="apiKey" type="password" autocomplete="off" placeholder="sk-ant-…" value="${esc(key)}"></label>
      <label class="f row" style="margin-top:6px"><input type="checkbox" id="remember" ${store.remembered(KEY) ? 'checked' : ''}> Remember on this device</label>
      <label class="f" style="margin-top:8px">Model<select class="in" id="model">${MODELS.map(([v, l]) => `<option value="${v}"${v === model ? ' selected' : ''}>${esc(l)}</option>`).join('')}</select></label>
    </div>`;
  const keyIn = $('#apiKey', el), rem = $('#remember', el);
  const keepKey = () => { if (keyIn.value.trim()) store.set(KEY, keyIn.value.trim(), rem.checked); else store.del(KEY); };
  keyIn.onchange = keepKey; rem.onchange = keepKey;
  $('#model', el).onchange = (e) => store.set(MODEL, e.target.value, true);
  $('#descText', el).oninput = (e) => localStorage.setItem(LS_DRAFT, e.target.value);
  el.querySelectorAll('[data-req]').forEach((b) => b.onclick = () => { const t = $('#descText', el); t.value = SAMPLE_REQUESTS[+b.dataset.req]; localStorage.setItem(LS_DRAFT, t.value); t.focus(); });
  $('#btnGen', el).onclick = () => runClaude('new', $('#descText', el).value);
  if (has) {
    el.querySelectorAll('[data-chg]').forEach((b) => b.onclick = () => { const t = $('#changeText', el); t.value = (t.value ? t.value.trim() + '\n' : '') + b.dataset.chg; t.focus(); });
    $('#btnChange', el).onclick = () => runClaude('change', $('#changeText', el).value);
  }
}

async function runClaude(kind, text) {
  text = String(text || '').trim();
  if (!text) { toast(kind === 'change' ? 'Describe the change first.' : 'Describe the web part first.', true); return; }
  const key = store.get(KEY) || $('#apiKey')?.value.trim();
  if (!key) { switchTab('build'); toast('Enter your Anthropic API key first.', true); $('#apiKey')?.focus(); return; }
  const model = store.get(MODEL) || MODELS[0][0];
  const ctrl = new AbortController(), started = Date.now();
  const label = kind === 'new' ? 'Claude is writing your web part' : kind === 'fix' ? 'Claude is fixing the error' : 'Claude is applying your change';
  S.busy = { label, ctrl }; switchTab('build'); renderBuild();
  const tick = setInterval(() => { const p = $('#prog'); if (p) p.textContent = `${label}… ${Math.round((Date.now() - started) / 1000)}s`; }, 1000);
  try {
    const history = kind === 'new' ? null : kind === 'fix' ? fixHistory(S.project.request, S.project.spec, text) : changeHistory(S.project.request, S.project.spec, text);
    let { spec } = await callClaude({ key, signal: ctrl.signal, request: buildRequest({ text, model, history }) });
    let res = cleanSpec(spec);
    let syntax = checkCode(res.spec.code);
    if (syntax) {
      S.busy.label = 'Fixing a syntax error';
      ({ spec } = await callClaude({ key, signal: ctrl.signal, request: buildRequest({ model, history: fixHistory(kind === 'new' ? text : S.project.request, res.spec, 'SyntaxError: ' + syntax) }) }));
      res = cleanSpec(spec); syntax = checkCode(res.spec.code);
    }
    S.notes = res.notes.concat(syntax ? ['The code still has a syntax error: ' + syntax + '. Open the Code tab or ask for a fix.'] : []);
    if (kind === 'new') {
      S.project = newProject(res.spec, text); S.rest = []; S.pane = null;
      save(); snapshot('Generated'); localStorage.removeItem(LS_DRAFT); renderAll(); reloadPreview(true);
    } else setSpec(res.spec, kind === 'fix' ? 'Fixed an error' : 'Change: ' + text.slice(0, 60));
    toast(kind === 'new' ? 'Web part ready. Try it in the preview, then download the .sppkg.' : 'Done.');
  } catch (e) {
    toast(e instanceof AiError ? e.message : 'Something went wrong: ' + e.message, true);
  } finally {
    clearInterval(tick); S.busy = null; renderBuild();
  }
}

// ------------------------------------------------------------------ left: Code tab
function renderCode() {
  const el = $('#tab-code');
  if (!S.project) { el.innerHTML = '<p class="hint">Generate a web part or open an example first.</p>'; return; }
  const s = S.project.spec, t = S.view.codeTab;
  const val = t === 'code' ? s.code : t === 'css' ? s.css : t === 'props' ? JSON.stringify(s.properties, null, 2) : JSON.stringify(s.lists, null, 2);
  el.innerHTML = `<div class="subtabs">${[['code', 'JavaScript'], ['css', 'CSS'], ['props', 'Properties'], ['lists', 'Lists']].map(([k, l]) => `<button data-ct="${k}" class="${k === t ? 'on' : ''}">${l}</button>`).join('')}</div>
    <p class="hint">${t === 'code' ? 'Body of an async function. In scope: <code>ctx, h, icon, esc, fmt, props, sp</code>. Build into <code>ctx.root</code>.' : t === 'css' ? 'Scoped to the web part automatically. Use <code>var(--sss-primary)</code> and friends to follow the site theme.' : t === 'props' ? 'Property pane fields: key, type (text, multiline, number, toggle, dropdown, slider, list, color, url), label, group, default.' : 'Lists the web part uses. Each gets a Create button in the property pane. sample rows feed the preview.'}</p>
    <textarea class="ta code" id="codeText" spellcheck="false">${esc(val)}</textarea>
    <div class="codeerr" id="codeErr">${esc(S.codeErr)}</div>
    <div style="display:flex;gap:8px"><button class="btn pri" id="btnApply">Apply (Ctrl+S)</button><button class="btn" id="btnRevert">Revert</button></div>`;
  el.querySelectorAll('[data-ct]').forEach((b) => b.onclick = () => { S.view.codeTab = b.dataset.ct; S.codeErr = ''; saveView(); renderCode(); });
  const ta = $('#codeText', el);
  ta.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') { e.preventDefault(); const a = ta.selectionStart; ta.setRangeText('  ', a, ta.selectionEnd, 'end'); }
    if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); applyCode(); }
  });
  $('#btnApply', el).onclick = applyCode;
  $('#btnRevert', el).onclick = () => { S.codeErr = ''; renderCode(); };
}
function applyCode() {
  const t = S.view.codeTab, v = $('#codeText').value, spec = clone(S.project.spec);
  try {
    if (t === 'code') { const err = checkCode(v); if (err) throw new Error('Syntax error: ' + err); spec.code = v; }
    else if (t === 'css') spec.css = v;
    else { const parsed = JSON.parse(v); if (!Array.isArray(parsed)) throw new Error('Expected a JSON array.'); spec[t === 'props' ? 'properties' : 'lists'] = parsed; }
  } catch (e) { S.codeErr = e.message; $('#codeErr').textContent = e.message; return; }
  const res = cleanSpec(spec);
  S.codeErr = ''; S.notes = res.notes;
  setSpec(res.spec, 'Edited ' + ({ code: 'code', css: 'CSS', props: 'properties', lists: 'lists' }[t]));
  toast('Applied.');
}

// ------------------------------------------------------------------ left: Package tab
function renderPackage() {
  const el = $('#tab-package');
  if (!S.project) { el.innerHTML = '<p class="hint">Generate a web part or open an example first.</p>'; return; }
  const p = S.project, s = p.spec, n = names(p);
  const changed = p.downloaded && p.downloaded !== specHash(p);
  el.innerHTML = `
    <label class="f">Toolbox title<input class="in" data-sf="title" value="${esc(s.title)}"></label>
    <label class="f">Description<textarea class="ta" data-sf="description" style="min-height:60px">${esc(s.description)}</textarea></label>
    <label class="f">Toolbox group<select class="in" data-sf="group">${Object.keys(GROUPS).map((g) => `<option${g === s.group ? ' selected' : ''}>${esc(g)}</option>`).join('')}</select></label>
    <label class="f">Toolbox icon (Fluent UI name)<input class="in" data-sf="icon" value="${esc(s.icon)}" placeholder="Page"></label>
    <label class="f">Developer name (shown in the app catalog)<input class="in" data-sf="developer" value="${esc(s.developer || '')}" placeholder="SPFx Studio"></label>
    <label class="f row"><input type="checkbox" data-sf="fullBleed" ${s.fullBleed ? 'checked' : ''}> Allow in full-width sections</label>
    <label class="f row"><input type="checkbox" data-sf="tenantWide" ${s.tenantWide !== false ? 'checked' : ''}> Make available to all sites when deployed</label>
    <div class="card"><h3>Version ${esc(componentVersion(p))}</h3><p class="hint">${changed ? 'You changed the web part since the last download, so the next download is version 1.0.' + (p.build + 1) + '. Upload it over the old package and every page picks it up.' : p.downloaded ? 'Matches your last download. The version goes up by one the next time you download a changed package.' : 'Not downloaded yet.'}</p></div>
    <div class="card"><h3>Identity</h3><div class="kv"><span>Package</span><code>${esc(n.pkg)}.sppkg</code><span>Solution</span><code>${esc(p.ids.solution)}</code><span>Web part</span><code>${esc(p.ids.component)}</code><span>Alias</span><code>${esc(n.alias)}</code></div>
      <p class="hint" style="margin-top:8px">Keep these to upgrade the same web part. <button class="btn sm" id="btnNewIds">New identity…</button></p></div>
    <div><h2 class="sec">Deploy</h2></div>${deploySteps()}
    <button class="btn pri block" id="btnDl2">Download ${esc(n.pkg)}.sppkg</button>
    <button class="btn block" id="btnSrc2">Download source project (.zip)</button>`;
  el.querySelectorAll('[data-sf]').forEach((i) => i.onchange = () => {
    const k = i.dataset.sf; s[k] = i.type === 'checkbox' ? i.checked : i.value;
    if (k === 'title') { $('#wpTitle').value = s.title; }
    save(); renderTop(); renderPane(); reloadPreview();
  });
  $('#btnNewIds', el).onclick = () => {
    const m = modal('New identity', '<p class="hint">A new identity makes this a separate web part: it installs next to the old one instead of upgrading it. Use this when you copied a project to start something new.</p>', '<button class="btn" data-close>Cancel</button><button class="btn pri" id="okIds">Make it a new web part</button>');
    $('#okIds', m.el).onclick = () => { p.ids = { solution: uuid(), feature: uuid(), component: uuid(), assets: uuid() }; p.build = 1; p.downloaded = ''; save(); m.close(); renderAll(); reloadPreview(); toast('New identity set.'); };
  };
  $('#btnDl2', el).onclick = downloadPackage;
  $('#btnSrc2', el).onclick = downloadSource;
}
function deploySteps() {
  const s = S.project?.spec;
  return `<ol class="steps">
    <li>Open your tenant app catalog: SharePoint admin center → <b>More features</b> → <b>Apps</b> → <b>Open</b> → <b>Apps for SharePoint</b>. GCC High works the same way.</li>
    <li><b>Upload</b> the .sppkg. In the dialog, keep <b>Make this solution available to all sites</b> checked and click <b>Deploy</b> (or <b>Enable app</b>).</li>
    <li>Edit a page, click <b>+</b>, and search for <b>${esc(s?.title || 'your web part')}</b>${s?.group ? ` (group: ${esc(s.group)})` : ''}.</li>
    ${s?.lists?.length ? `<li>Open the web part's property pane → <b>Lists</b> → <b>Create "${esc(s.lists[0].name)}" list</b>${s.lists.length > 1 ? ' (and the others)' : ''}. That needs Manage Lists permission on the site. Sample rows are added when a list is created.</li>` : ''}
    <li>To update: download again (the version goes up) and upload over the old package. Pages pick up the new version.</li>
  </ol>`;
}

// ------------------------------------------------------------------ downloads
async function downloadPackage() {
  if (!S.project) return;
  if (!window.JSZip) { toast('The zip library did not load. Refresh the page.', true); return; }
  const syntax = checkCode(S.project.spec.code);
  if (syntax) { toast('Fix the syntax error first: ' + syntax, true); return; }
  const h = specHash(S.project);
  if (S.project.downloaded && S.project.downloaded !== h) S.project.build = (S.project.build || 1) + 1;
  S.project.downloaded = h; save();
  const { data, fileName } = await buildSppkg(S.project, window.JSZip);
  download(fileName, data);
  renderTop(); renderPackage();
  if (!sessionStorage.getItem('sss.deployShown')) { sessionStorage.setItem('sss.deployShown', '1'); deployDialog(fileName); }
  else toast(`Downloaded ${fileName} (version ${componentVersion(S.project)}).`);
}
async function downloadSource() {
  if (!S.project || !window.JSZip) return;
  const { files, folder } = sourceFiles(S.project);
  const z = new window.JSZip(), root = z.folder(folder);
  for (const [p, c] of files) root.file(p, c);
  download(folder + '-spfx-source.zip', await z.generateAsync({ type: 'blob', compression: 'DEFLATE' }));
  toast('Source project downloaded. npm install, then gulp bundle --ship and gulp package-solution --ship.');
}
function deployDialog(fileName) {
  modal('Deploy your web part', `${fileName ? `<p class="hint"><b>${esc(fileName)}</b> (version ${esc(componentVersion(S.project))}) is in your downloads.</p>` : ''}${deploySteps()}
    <p class="hint">The package carries everything it needs (includeClientSideAssets). It makes no calls outside your SharePoint site, so it works in GCC High.</p>`);
}

// ------------------------------------------------------------------ start screen, examples, project menu
function renderStart() {
  const el = $('#start');
  el.hidden = !!S.project;
  $('#spPage').hidden = !S.project;
  if (S.project) return;
  el.innerHTML = `<div class="start-in">
    <div><h1>Describe a SharePoint web part. Download the .sppkg.</h1>
    <p class="lede">Claude writes the web part, the preview runs it against sample data with the real property pane, and the download is a standard SPFx package for your app catalog. No Node, no gulp, nothing loaded from outside SharePoint.</p></div>
    <div class="flow"><div><b>1 · Describe</b>What it shows, where data lives, what editors change.</div><div><b>2 · Preview</b>Sample data, widths, section colours, edit mode.</div><div><b>3 · Refine</b>Ask for changes, or edit code and CSS.</div><div><b>4 · Deploy</b>Upload the .sppkg to the app catalog.</div></div>
    <div><h2 class="sec" style="margin-bottom:10px">Or start from an example</h2><div class="exgrid">${EXAMPLES.map((e) => `<button class="excard" data-ex="${e.key}"><b>${esc(e.spec.title)}</b><span>${esc(e.blurb)}</span></button>`).join('')}</div></div>
  </div>`;
  el.querySelectorAll('[data-ex]').forEach((b) => b.onclick = () => openExample(b.dataset.ex));
}
function openExample(key) {
  const ex = EXAMPLES.find((e) => e.key === key);
  S.project = newProject(ex.spec, ''); S.rest = []; S.pane = null; S.notes = [];
  save(); snapshot('Example: ' + ex.spec.title); renderAll(); reloadPreview(true);
  toast(`${ex.spec.title} loaded. Change it with Claude, or edit the code.`);
}
function examplesDialog() {
  const m = modal('Examples', `<p class="hint">Hand-built web parts. Open one to see the standard Claude aims for, or use it as a starting point. Opening an example replaces the current web part (it stays in Version history).</p><div class="exgrid">${EXAMPLES.map((e) => `<button class="excard" data-ex="${e.key}"><b>${esc(e.spec.title)}</b><span>${esc(e.blurb)}</span></button>`).join('')}</div>`, '<button class="btn" data-close>Close</button>');
  m.el.querySelectorAll('[data-ex]').forEach((b) => b.onclick = () => { m.close(); openExample(b.dataset.ex); });
}
function historyDialog() {
  const h = load(LS_HISTORY, []);
  const m = modal('Version history', h.length ? `<div class="hist">${h.map((x, i) => `<button data-h="${i}"><span>${esc(x.project.spec.title)} · ${esc(x.label)}</span><small>${esc(new Date(x.at).toLocaleString())}</small></button>`).join('')}</div>` : '<p class="hint">Nothing yet. Each generate, change and code edit is kept here (last 25).</p>', '<button class="btn" data-close>Close</button>');
  m.el.querySelectorAll('[data-h]').forEach((b) => b.onclick = () => { snapshot('Before restore'); S.project = clone(h[+b.dataset.h].project); S.pane = null; save(); m.close(); renderAll(); reloadPreview(true); toast('Restored.'); });
}
$('#projMenu').addEventListener('click', (e) => {
  const dd = $('#projMenu');
  if (e.target.id === 'btnProject') { dd.classList.toggle('open'); return; }
  const c = e.target.closest('[data-cmd]')?.dataset.cmd; if (!c) return;
  dd.classList.remove('open');
  if (c === 'new') { if (S.project) snapshot('Before new'); S.project = null; S.pane = null; S.notes = []; save(); renderAll(); }
  if (c === 'save' && S.project) download(names(S.project).pkg + '.spfxstudio.json', JSON.stringify(S.project, null, 2), 'application/json');
  if (c === 'open') $('#fileOpen').click();
  if (c === 'history') historyDialog();
  if (c === 'deploy') deployDialog();
});
document.addEventListener('click', (e) => { if (!e.target.closest('#projMenu')) $('#projMenu').classList.remove('open'); });
$('#fileOpen').onchange = async (e) => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const p = JSON.parse(await f.text());
    if (!p.spec || !p.ids?.component) throw new Error('Not an SPFx Studio project file.');
    const res = cleanSpec(p.spec); p.spec = res.spec; p.props = defaultProps(p.spec, p.props || {});
    if (S.project) snapshot('Before open');
    S.project = p; S.pane = null; save(); renderAll(); reloadPreview(true); toast('Project opened.');
  } catch (err) { toast('Could not open it: ' + err.message, true); }
};

// ------------------------------------------------------------------ top bar, toolbar, tabs
function renderTop() {
  $('#wpTitle').value = S.project?.spec.title || '';
  $('#wpTitle').disabled = !S.project;
  $('#verBadge').textContent = S.project ? 'v' + componentVersion(S.project) : '';
  $('#btnPackage').disabled = !S.project; $('#btnSource').disabled = !S.project;
}
$('#wpTitle').onchange = (e) => { if (!S.project) return; S.project.spec.title = e.target.value.trim() || S.project.spec.title; save(); renderPane(); renderPackage(); reloadPreview(); };
$('#btnPackage').onclick = downloadPackage;
$('#btnSource').onclick = downloadSource;
$('#btnExamples').onclick = examplesDialog;
function switchTab(t) { S.view.tab = t; saveView(); document.querySelectorAll('.tabs [data-tab]').forEach((b) => b.classList.toggle('on', b.dataset.tab === t)); for (const k of ['build', 'code', 'package']) $('#tab-' + k).hidden = k !== t; }
document.querySelectorAll('.tabs [data-tab]').forEach((b) => b.onclick = () => switchTab(b.dataset.tab));
function renderToolbar() {
  document.querySelectorAll('#segWidth button').forEach((b) => b.classList.toggle('on', b.dataset.w === S.view.width));
  document.querySelectorAll('#segSection button').forEach((b) => b.classList.toggle('on', b.dataset.s === S.view.section));
  $('#editMode').checked = S.view.edit;
  const sec = $('#spSection'); sec.className = 'sp-section ' + S.view.width;
  const t = themeObj(); sec.style.background = t.semanticColors.bodyBackground;
  $('#spPage').style.setProperty('--sp-suite', (THEMES[S.view.theme] || THEMES.Blue)[0]);
  $('#spPage').classList.toggle('is-edit', S.view.edit);
  const sel = $('#themeSel'); if (!sel.options.length) sel.innerHTML = Object.keys(THEMES).map((k) => `<option>${k}</option>`).join('');
  sel.value = S.view.theme;
}
$('#segWidth').onclick = (e) => { const b = e.target.closest('[data-w]'); if (!b) return; S.view.width = b.dataset.w; saveView(); renderToolbar(); };
$('#segSection').onclick = (e) => { const b = e.target.closest('[data-s]'); if (!b) return; S.view.section = b.dataset.s; saveView(); renderToolbar(); post({ type: 'theme', theme: themeObj() }); };
$('#themeSel').onchange = (e) => { S.view.theme = e.target.value; saveView(); renderToolbar(); post({ type: 'theme', theme: themeObj() }); };
$('#editMode').onchange = (e) => { S.view.edit = e.target.checked; saveView(); renderToolbar(); post({ type: 'mode', edit: S.view.edit }); };
$('#btnReseed').onclick = () => { if (S.project) { post({ type: 'reseed', lists: S.project.spec.lists }); toast('Sample data reset.'); } };
$('#btnLog').onclick = () => { S.view.log = !S.view.log; saveView(); renderLog(); };

function renderAll() { renderTop(); renderStart(); renderToolbar(); renderBuild(); renderCode(); renderPackage(); renderPane(); renderLog(); renderErr(); }

// ------------------------------------------------------------------ boot
frame.srcdoc = previewDoc();
switchTab(S.view.tab);
renderAll();
if (S.project) { S.project.props = defaultProps(S.project.spec, S.project.props || {}); reloadPreview(false); }
