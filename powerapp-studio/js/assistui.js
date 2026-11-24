// assistui.js — Paste Assistant (in Get YAML), Finish setup card/dialog, customer guide export
// and the first-visit tour. ui.js passes in its helpers through initAssistUI().
import { deploySteps, stepStatus, markDeployed, pendingSteps, removedSinceDeploy } from './deploy.js';
import { setupItems, openSetupItems, applySetup, GUID } from './setup.js';
import { guideHtml } from './guide.js';
import { walk } from './core.js';

let H = null; // { P, modal, toast, ui, esc, commit, copyText, download, downloadFlow, renderThumb, renderAll, lintSoon, slug, onSetupChange }
export function initAssistUI(helpers) { H = helpers; }

// ------------------------------------------------------------------ Finish setup
export function setupCardHtml({ compact = false } = {}) {
  const P = H.P(), esc = H.esc, ui = H.ui;
  const items = setupItems(P);
  if (!items.length) return '';
  const open = items.filter((i) => !i.done).length;
  const row = (it) => {
    let field = '';
    if (it.kind === 'url' || it.kind === 'guid') field = `<input data-su="${it.id}" value="${esc(it.value)}" placeholder="${it.kind === 'guid' ? '00000000-0000-0000-0000-000000000000' : 'https://'}" spellcheck="false" autocomplete="off" aria-label="${esc(it.label)}">`;
    if (it.kind === 'deeplink') field = `<div class="su-two"><select data-su="deepLink" data-part="screen" aria-label="Screen">${['', ...P.screens.map((s) => s.id)].map((v) => `<option value="${esc(v)}" ${v === (it.value.screen || '') ? 'selected' : ''}>${v ? esc(P.screens.find((s) => s.id === v).name) : '(screen)'}</option>`).join('')}</select>
      <select data-su="deepLink" data-part="list" aria-label="List">${['', ...P.lists.filter((l) => !l.system).map((l) => l.name)].map((v) => `<option value="${esc(v)}" ${v === (it.value.list || '') ? 'selected' : ''}>${v ? esc(v) : '(list)'}</option>`).join('')}</select></div>`;
    if (it.kind === 'confirm') field = `<label class="chk" style="margin:4px 0 0"><input type="checkbox" data-su="${it.id}" ${it.value ? 'checked' : ''}>Done: the list exists and users can contribute</label>`;
    return `<div class="su-item ${it.done ? 'done' : ''}" data-item="${it.id}"><i class="su-st" aria-hidden="true">${it.done ? ui('check') : ''}</i><div class="su-body"><b>${esc(it.label)}</b>${compact ? '' : `<small>${esc(it.help)}</small>`}${field}<em class="su-hint">${esc(it.hint || '')}</em></div></div>`;
  };
  return `<div class="setup-card ${open ? '' : 'all-done'}" id="setupCard"><div class="su-h"><b>${open ? 'Finish setup' : 'Setup complete'}</b><span class="su-count">${open ? `${open} of ${items.length} left` : `${items.length} of ${items.length} done`}</span></div>${items.map(row).join('')}</div>`;
}
/** Wire a rendered setup card: inputs update the project live and repaint check marks without losing focus. */
export function bindSetupCard(root) {
  const card = root.querySelector('#setupCard'); if (!card) return;
  const P = H.P();
  const repaint = () => {
    const items = setupItems(P), open = items.filter((i) => !i.done).length;
    for (const it of items) {
      const el = card.querySelector(`[data-item="${it.id}"]`); if (!el) continue;
      el.classList.toggle('done', it.done);
      el.querySelector('.su-st').innerHTML = it.done ? H.ui('check') : '';
      el.querySelector('.su-hint').textContent = it.hint || '';
    }
    card.classList.toggle('all-done', !open);
    card.querySelector('.su-h b').textContent = open ? 'Finish setup' : 'Setup complete';
    card.querySelector('.su-count').textContent = open ? `${open} of ${items.length} left` : `${items.length} of ${items.length} done`;
    H.onSetupChange();
  };
  const handler = (e) => {
    const t = e.target; if (!t.dataset.su) return;
    const id = t.dataset.su;
    if (id === 'deepLink') {
      applySetup(P, 'deepLink', { [t.dataset.part]: t.value });
      const ls = card.querySelector('[data-su="deepLink"][data-part="list"]');
      if (ls && P.app.deepLink?.list) ls.value = P.app.deepLink.list;
    } else applySetup(P, id, t.type === 'checkbox' ? t.checked : t.value);
    if (id === 'adminGroup') t.classList.toggle('bad', !!t.value.trim() && !GUID.test(t.value.trim()));
    H.commit(); repaint();
  };
  card.addEventListener('input', handler);
  card.addEventListener('change', (e) => { if (e.target.tagName === 'SELECT' || e.target.type === 'checkbox') handler(e); });
}
export function setupDialog() {
  const P = H.P();
  if (!setupItems(P).length) { H.toast('Nothing to set up: this app has no flows, email links, admin group or usage tracking.'); return; }
  const m = H.modal(`<div class="modal-h"><h2>Finish setup</h2><button class="mini" data-close>${H.ui('close')}</button></div>
    <div class="modal-b">${setupCardHtml()}<p class="note">These values live in the project. Fixing them here clears the matching lint errors right away.</p></div>
    <div class="modal-f"><span class="grow"></span><button class="btn pri" data-close>Done</button></div>`, 'md');
  bindSetupCard(m.el);
  setTimeout(() => m.el.querySelector('.su-item:not(.done) input, .su-item:not(.done) select')?.focus(), 30);
}
export const openSetupCount = () => openSetupItems(H.P()).length;

// ------------------------------------------------------------------ Paste Assistant
const STATUS = { new: 'New', changed: 'Changed', unchanged: 'Unchanged' };
export function deployView(main, closeDialog) {
  const P = H.P(), esc = H.esc, ui = H.ui;
  const st = { mode: Object.keys(P.deployed || {}).length ? 'changed' : 'all', i: 0 };
  const all = () => deploySteps(P);
  const list = () => (st.mode === 'changed' ? pendingSteps(P, all()) : all());
  const draw = () => {
    const steps = list(), every = all();
    if (st.i >= steps.length) st.i = Math.max(0, steps.length);
    const changedCount = pendingSteps(P, every).length;
    const doneHere = steps.filter((s) => stepStatus(P, s) === 'unchanged').length;
    const deployedBefore = Object.keys(P.deployed || {}).length > 0;
    const removed = removedSinceDeploy(P);
    const cur = steps[st.i];
    const openN = openSetupItems(P).length, anySetup = setupItems(P).length;
    const setup = openN ? setupCardHtml({ compact: true }) : anySetup ? `<div class="su-ok">${ui('check')}<span>Setup complete: site, links${P.app.features?.adminGroup ? ', admin group' : ''}${P.app.features?.usage ? ', usage list' : ''}.</span><button class="btn sm" id="dpSetup">Review</button></div>` : '';
    main.innerHTML = `<div class="bar"><h3>Paste assistant</h3>
        <div class="seg" role="group" aria-label="Which steps"><button data-mode="all" class="${st.mode === 'all' ? 'on' : ''}">All steps (${every.length})</button><button data-mode="changed" class="${st.mode === 'changed' ? 'on' : ''}" ${deployedBefore ? '' : 'disabled title="Available after your first deploy"'}>Only what changed (${deployedBefore ? changedCount : '–'})</button></div></div>
      <div class="dp-prog" role="progressbar" aria-valuemin="0" aria-valuemax="${steps.length}" aria-valuenow="${doneHere}"><i style="width:${steps.length ? (100 * doneHere) / steps.length : 100}%"></i></div>
      <p class="note" style="margin:0">${doneHere} of ${steps.length} ${st.mode === 'changed' ? 'changed steps' : 'steps'} done. Each step is marked with what it pastes, so next time only changed screens and flows need re-pasting.</p>
      ${setup ? `<div class="dp-setup">${setup}</div>` : ''}
      <div class="dp-grid"><ol class="dp-list">${steps.map((s, i) => { const ss = stepStatus(P, s); return `<li><button data-go="${i}" class="${i === st.i ? 'on' : ''} ${ss}"><i>${ss === 'unchanged' ? ui('check') : i + 1}</i><span>${esc(s.title)}</span>${deployedBefore && s.hash ? `<em class="chip ${ss}">${STATUS[ss]}</em>` : ''}</button></li>`; }).join('')}</ol>
      <div class="dp-step">${cur ? stepHtml(cur, st.i, steps.length) : finishedHtml(removed, deployedBefore)}</div></div>`;
    main.querySelectorAll('[data-mode]').forEach((b) => b.onclick = () => { st.mode = b.dataset.mode; st.i = 0; draw(); });
    main.querySelectorAll('[data-go]').forEach((b) => b.onclick = () => { st.i = +b.dataset.go; draw(); });
    const q = (s) => main.querySelector(s);
    if (cur) {
      q('#dpCopy') && (q('#dpCopy').onclick = () => H.copyText(cur.text));
      q('#dpDl') && (q('#dpDl').onclick = () => H.downloadFlow(cur.flow));
      q('#dpDone').onclick = () => { markDeployed(P, cur); H.commit(); const steps2 = list(); st.i = st.mode === 'changed' ? steps2.findIndex((s) => stepStatus(P, s) !== 'unchanged') : st.i + 1; if (st.i < 0) st.i = steps2.length; draw(); };
      q('#dpSkip').onclick = () => { st.i += 1; draw(); };
      q('#dpBack') && (q('#dpBack').onclick = () => { st.i = Math.max(0, st.i - 1); draw(); });
    } else {
      q('#dpRestart') && (q('#dpRestart').onclick = () => { st.mode = 'all'; st.i = 0; draw(); });
      q('#dpForget') && (q('#dpForget').onclick = () => { P.deployed = {}; H.commit(); st.mode = 'all'; st.i = 0; draw(); H.toast('Deploy record cleared'); });
      q('#dpClose') && (q('#dpClose').onclick = closeDialog);
    }
    bindSetupCard(main);
    q('#dpSetup') && (q('#dpSetup').onclick = () => setupDialog());
  };
  const stepHtml = (s, i, n) => {
    const ss = stepStatus(P, s);
    const repaste = s.kind === 'screen' && P.deployed?.[s.key];
    return `<div class="dp-head"><span class="dp-n">Step ${i + 1} of ${n}</span>${P.deployed?.[s.key] ? `<em class="chip ${ss}">${STATUS[ss]}</em>` : ''}</div>
      <h4>${esc(s.title)}</h4><p class="dp-how">${esc(s.how)}</p>
      ${repaste ? `<div class="note-box warn"><b>Before you delete ${esc(s.screen.name)} in Studio</b>Re-pasting replaces the whole screen. Controls you added by hand in Studio (an attachments form, for example) are lost: select them, Ctrl+C, paste into Notepad, and paste them back after this step.</div>` : ''}
      <div class="dp-acts">${s.text ? `<button class="btn pri big" id="dpCopy">${ui('copy')}${esc(s.copyLabel || 'Copy')}</button>` : ''}${s.flow ? `<button class="btn pri big" id="dpDl">${ui('download')}Download ${esc(s.file)}</button>` : ''}
        <button class="btn big" id="dpDone">${ui('check')}Done, next</button></div>
      <div class="dp-sub">${i > 0 ? '<button class="btn sm" id="dpBack">Back</button>' : ''}<button class="btn sm" id="dpSkip">Skip</button>${s.text ? `<details><summary class="note">Show what gets copied</summary><pre class="code small">${esc(s.text.length > 6000 ? s.text.slice(0, 6000) + '\n…' : s.text)}</pre></details>` : ''}</div>`;
  };
  const finishedHtml = (removed, before) => `<div class="dp-done">${ui('check')}<h4>${before ? 'Everything is deployed' : 'Nothing left in this list'}</h4>
    <p class="note">Edits you make from now on show up here as Changed, so only those screens and flows need re-pasting.</p>
    ${removed.length ? `<div class="note-box warn"><b>Removed since the last deploy</b>Delete these in Studio or Power Automate too: ${removed.map((r) => esc(r.name)).join(', ')}.</div>` : ''}
    <div class="btns"><button class="btn" id="dpRestart">Show all steps</button><button class="btn danger" id="dpForget">Forget deploy record</button><button class="btn pri" id="dpClose">Close</button></div></div>`;
  draw();
}

// ------------------------------------------------------------------ guide export
/** CSS rules the rendered controls need, copied from the studio's own stylesheet. */
function pageControlCss() {
  const out = [];
  for (const sh of document.styleSheets) {
    let rules; try { rules = sh.cssRules; } catch { continue; }
    for (const r of rules || []) if (r.selectorText && /\.pa-/.test(r.selectorText) && !/\.design|:hover|pa-cond|pa-ghost|pa-tpl|pa-gal-empty|ghost-row/.test(r.selectorText)) out.push(r.cssText);
  }
  return out.join('\n');
}
export function exportGuide() {
  const P = H.P();
  const thumbs = {};
  for (const s of P.screens) { try { thumbs[s.id] = H.renderThumb(s); } catch { /* skip a screen that fails to render */ } }
  const html = guideHtml(P, { thumbs, css: pageControlCss() });
  H.download(H.slug(P.app.name) + '-guide.html', html, 'text/html');
  return html;
}

// ------------------------------------------------------------------ tour
const TOUR = [
  { sel: '.rail [data-tab="insert"]', title: 'Build with blocks', text: 'Insert has every control plus ready-made blocks. Drag them onto the screen; Data, Flows and Theme live in this rail too.' },
  { sel: '#viewport', title: 'The canvas is the real app', text: 'Click any control to edit it on the right. Drop a control into a gallery\'s first row and it repeats on every row.' },
  { sel: '#btnPlay', title: 'Try it with sample data', text: 'Play runs the app: search, open items, save, approve, and see which flows would email.' },
  { sel: '#btnCode', title: 'Paste it into Power Apps', text: 'Get YAML opens the Paste assistant: one copy button per step, and next time only what changed.' },
];
const TOUR_KEY = 'pas.tour.done';
export const tourDone = () => { try { return localStorage.getItem(TOUR_KEY) === '1'; } catch { return false; } };
export function startTour() {
  document.querySelector('.tour')?.remove();
  let i = 0;
  const box = document.createElement('div');
  box.className = 'tour'; box.setAttribute('role', 'dialog'); box.setAttribute('aria-modal', 'false'); box.setAttribute('aria-label', 'Tour');
  document.body.appendChild(box);
  const end = (done) => { box.remove(); window.removeEventListener('resize', place); document.removeEventListener('keydown', key, true); if (done) { try { localStorage.setItem(TOUR_KEY, '1'); } catch { /* ignore */ } } };
  const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); end(true); } };
  function place() {
    const step = TOUR[i], t = document.querySelector(step.sel);
    if (!t) { end(true); return; }
    const r = t.getBoundingClientRect(), pad = 6;
    box.innerHTML = `<div class="tour-hole" style="left:${r.left - pad}px;top:${r.top - pad}px;width:${r.width + pad * 2}px;height:${r.height + pad * 2}px"></div>
      <div class="tour-tip"><span class="tour-n">${i + 1} / ${TOUR.length}</span><b>${H.esc(step.title)}</b><p>${H.esc(step.text)}</p>
      <div class="btns"><button class="btn sm" data-t="skip">Skip tour</button><span style="flex:1"></span>${i ? '<button class="btn sm" data-t="back">Back</button>' : ''}<button class="btn sm pri" data-t="next">${i === TOUR.length - 1 ? 'Finish' : 'Next'}</button></div></div>`;
    const tip = box.querySelector('.tour-tip'), tw = 320, vw = window.innerWidth, vh = window.innerHeight;
    let x = r.right + 14, y = r.top;
    if (x + tw > vw - 12) x = Math.max(12, r.left - tw - 14);
    if (r.width > vw * 0.5) { x = r.left + r.width / 2 - tw / 2; y = r.top + 24; }
    tip.style.left = Math.max(12, Math.min(vw - tw - 12, x)) + 'px'; tip.style.top = Math.max(12, Math.min(vh - 200, y)) + 'px';
    box.querySelector('[data-t="next"]').focus();
  }
  box.addEventListener('click', (e) => {
    const b = e.target.closest('[data-t]'); if (!b) return;
    if (b.dataset.t === 'skip') return end(true);
    if (b.dataset.t === 'back') i = Math.max(0, i - 1);
    if (b.dataset.t === 'next') { if (i === TOUR.length - 1) return end(true); i++; }
    place();
  });
  window.addEventListener('resize', place); document.addEventListener('keydown', key, true);
  place();
}
/** Number of controls on a screen with a Visible rule, for the designer's "N conditional" chip. */
export function conditionalCount(screen) {
  let n = 0; for (const [c] of walk(screen.controls)) if (c.visibleWhen?.var || c.fx?.Visible) n++;
  return n;
}
