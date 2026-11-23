// describeui.js — the "Describe your app" dialog: requirements text → Claude plan → review → build.
import { MODELS, callClaude, refineHistory, specToProject, AiError } from './ai.js';
import { FLOW_KINDS, recipientsText } from './flows.js';
import { COLTYPES } from './core.js';

let H = null; // { P, modal, toast, ui, esc, ask, setProject, commit }
export function initDescribeUI(helpers) { H = helpers; }

const KEY = 'pas.aikey', DRAFT = 'pas.aidraft', MODEL = 'pas.aimodel';
const store = {
  get(k) { try { return localStorage.getItem(k) ?? sessionStorage.getItem(k) ?? ''; } catch { return ''; } },
  set(k, v, remember) { try { (remember ? localStorage : sessionStorage).setItem(k, v); (remember ? sessionStorage : localStorage).removeItem(k); } catch { /* storage blocked */ } },
  del(k) { try { localStorage.removeItem(k); sessionStorage.removeItem(k); } catch { /* storage blocked */ } },
  remembered(k) { try { return !!localStorage.getItem(k); } catch { return false; } },
};
const EXAMPLE = `We need an app for employees to request training.
A request has a course name, the provider, cost, start date, a justification and the employee's supervisor.
Supervisors approve or reject requests.
Email the supervisor when a request is submitted, and email the employee when it is approved or rejected.
Every morning, remind supervisors about requests that have been waiting more than 3 days.
Employees should only see their own requests; supervisors see everything waiting for them.`;

export function describeDialog() {
  const esc = H.esc, ui = H.ui;
  const st = { text: store.get(DRAFT), plan: null, raw: '', notes: [], usage: null, busy: null };
  const m = H.modal(`<div class="modal-h"><h2>Describe your app</h2><span class="kindtag">Uses Claude</span><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b" id="dBody"></div><div class="modal-f" id="dFoot"></div>`);
  const body = m.el.querySelector('#dBody'), foot = m.el.querySelector('#dFoot');
  const opts = (arr, v) => arr.map(([val, lab]) => `<option value="${esc(val)}" ${val === v ? 'selected' : ''}>${esc(lab)}</option>`).join('');

  function drawWrite() {
    body.innerHTML = `<div class="flow-grid">
      <div><label class="f"><span>What should the app do? Who uses it, what do they track, who approves, who gets emailed and when.</span>
        <textarea class="plain" id="dText" rows="15" placeholder="${esc(EXAMPLE)}">${esc(st.text)}</textarea></label>
        <div class="btns"><button class="btn sm" id="dEx">Use the example</button></div></div>
      <div>
        <label class="f"><span>Anthropic API key</span><div class="rrow" style="margin:0"><input type="password" id="dKey" autocomplete="off" spellcheck="false" placeholder="sk-ant-…" value="${esc(store.get(KEY))}"><button class="btn sm" id="dShow" type="button">Show</button></div></label>
        <label class="chk"><input type="checkbox" id="dRem" ${store.remembered(KEY) ? 'checked' : ''}>Remember the key on this browser</label>
        <label class="f"><span>Model</span><select id="dModel">${opts(MODELS, store.get(MODEL) || MODELS[0][0])}</select></label>
        <label class="f"><span>SharePoint site URL <small>(optional, for flows)</small></span><input id="dSite" value="${esc(H.P().app.siteUrl || '')}" placeholder="https://contoso.sharepoint.com/sites/Team"></label>
        <div class="note-box">
          <b>Before you send</b>
          <ul><li>Your description goes from this browser straight to Anthropic's API and is billed to your key. Nothing passes through this website.</li>
          <li>The key stays in this browser tab, or in this browser if you tick Remember. It is never saved in project files or exports.</li>
          <li>Do not include sensitive, personal or controlled information.</li>
          <li>Government and corporate networks sometimes block api.anthropic.com.</li></ul>
          <a href="https://platform.claude.com/" target="_blank" rel="noopener">Get an API key in the Claude Console</a>
        </div>
      </div></div>`;
    foot.innerHTML = `<span class="grow">You review the plan before anything is built.</span><button class="btn" data-x>Cancel</button><button class="btn pri" id="dGo">Plan my app</button>`;
    const t = body.querySelector('#dText');
    t.addEventListener('input', () => { st.text = t.value; store.set(DRAFT, t.value, true); });
    body.querySelector('#dEx').onclick = () => { t.value = EXAMPLE; st.text = EXAMPLE; store.set(DRAFT, EXAMPLE, true); t.focus(); };
    body.querySelector('#dShow').onclick = (e) => { const k = body.querySelector('#dKey'); k.type = k.type === 'password' ? 'text' : 'password'; e.target.textContent = k.type === 'password' ? 'Show' : 'Hide'; };
    foot.querySelector('[data-x]').onclick = m.close;
    foot.querySelector('#dGo').onclick = () => run();
    setTimeout(() => t.focus(), 0);
  }
  function creds() {
    const key = (body.querySelector('#dKey')?.value ?? store.get(KEY)).trim();
    const remember = body.querySelector('#dRem')?.checked ?? store.remembered(KEY);
    const model = body.querySelector('#dModel')?.value || store.get(MODEL) || MODELS[0][0];
    const site = body.querySelector('#dSite')?.value ?? st.site ?? '';
    return { key, remember, model, site };
  }
  async function run(change) {
    const c = creds();
    if (!change && !st.text.trim()) { H.toast('Describe the app first', true); return; }
    if (!c.key) { H.toast('Enter your Anthropic API key', true); return; }
    store.set(KEY, c.key, c.remember);
    store.set(MODEL, c.model, true);
    st.site = c.site; st.key = c.key; st.model = c.model;
    const ctrl = new AbortController(), t0 = Date.now();
    st.busy = ctrl;
    body.innerHTML = `<div class="ai-wait"><div class="ai-spin"></div><b>${change ? 'Revising the plan…' : 'Planning your app…'}</b><span id="dTime">0 s</span><small>Claude is reading the requirements and designing lists, screens and flows. This usually takes 20–90 seconds.</small></div>`;
    foot.innerHTML = `<span class="grow"></span><button class="btn" id="dCancel">Cancel</button>`;
    foot.querySelector('#dCancel').onclick = () => ctrl.abort();
    const timer = setInterval(() => { const el = body.querySelector('#dTime'); if (el) el.textContent = Math.round((Date.now() - t0) / 1000) + ' s'; }, 500);
    try {
      const r = await callClaude({ key: c.key, model: c.model, text: st.text, history: change ? refineHistory(st.text, st.raw, change) : [], signal: ctrl.signal });
      st.plan = r.plan; st.raw = r.raw; st.usage = r.usage;
      const built = specToProject(r.plan, { siteUrl: st.site });
      st.project = built.project; st.notes = built.notes;
      drawReview();
    } catch (e) {
      const msg = e instanceof AiError ? e.message : 'Something went wrong: ' + (e?.message || e);
      if (st.plan) drawReview(msg); else { drawWrite(); H.toast(msg, true); }
    } finally { clearInterval(timer); st.busy = null; }
  }
  function drawReview(error) {
    const p = st.project, pl = st.plan;
    const scr = p.screens.map((s) => s.name).join(', ');
    body.innerHTML = `${error ? `<div class="note-box err">${esc(error)}</div>` : ''}
      <div class="flow-grid"><div>
        <h3 class="rv-h">${esc(p.app.name)} <small>${p.app.formFactor === 'phone' ? 'Phone' : 'Tablet'} · ${esc(p.theme.name)}</small></h3>
        ${p.lists.map((l) => `<div class="listcard"><div class="listcard-h"><b>${esc(l.name)}</b></div>
          ${(pl.lists || []).find((x) => x.name === l.name)?.purpose ? `<p class="note" style="padding:6px 10px 0;margin:0">${esc(pl.lists.find((x) => x.name === l.name).purpose)}</p>` : ''}
          <table class="rv-t">${l.columns.map((c) => `<tr><td>${esc(c.name)}${c.required ? ' <i>*</i>' : ''}</td><td>${esc(COLTYPES[c.type]?.label || c.type)}${c.type === 'Choice' ? `<small>${esc(c.choices.join(' · '))}</small>` : c.type === 'Lookup' ? `<small>→ ${esc(c.lookupList)}</small>` : ''}</td></tr>`).join('')}</table></div>`).join('')}
        <div class="sub">Screens</div><p class="note" style="margin-top:0">${esc(scr)}</p>
        <div class="sub">Flows</div>
        ${p.flows.length ? p.flows.map((f) => `<div class="rv-flow"><b>${esc(f.name)}</b><small>${esc(FLOW_KINDS[f.kind].label)} · ${esc(f.list)} · to ${esc(recipientsText(f))}</small></div>`).join('') : '<p class="note" style="margin-top:0">None.</p>'}
      </div><div>
        ${(pl.questions || []).length ? `<div class="note-box warn"><b>Questions worth answering</b><ul>${pl.questions.map((q) => `<li>${esc(q)}</li>`).join('')}</ul></div>` : ''}
        ${(pl.assumptions || []).length ? `<div class="note-box"><b>Assumptions Claude made</b><ul>${pl.assumptions.map((q) => `<li>${esc(q)}</li>`).join('')}</ul></div>` : ''}
        ${st.notes.length ? `<div class="note-box"><b>Adjusted to fit the generator</b><ul>${st.notes.map((q) => `<li>${esc(q)}</li>`).join('')}</ul></div>` : ''}
        <label class="f"><span>Want changes? Answer the questions or describe what to change.</span><textarea class="plain" id="dChange" rows="5" placeholder="Add a Department choice column (Finance, IT, HR). Don't email on new requests."></textarea></label>
        <div class="btns"><button class="btn" id="dRefine">${ui('refresh')}Update the plan</button></div>
        ${st.usage ? `<p class="note">Used ${Number(st.usage.input_tokens || 0).toLocaleString()} input and ${Number(st.usage.output_tokens || 0).toLocaleString()} output tokens.</p>` : ''}
      </div></div>`;
    foot.innerHTML = `<span class="grow">Building replaces the current project. Everything stays editable afterwards.</span><button class="btn" id="dBack">Edit the description</button><button class="btn pri" id="dBuild">Build this app</button>`;
    body.querySelector('#dRefine').onclick = () => { const ch = body.querySelector('#dChange').value.trim(); if (!ch) { H.toast('Describe the change first', true); return; } run(ch); };
    foot.querySelector('#dBack').onclick = drawWrite;
    foot.querySelector('#dBuild').onclick = async () => {
      const cur = H.P();
      if (cur.screens.some((s) => s.controls.length) && !(await H.ask('Replace the current app?', 'Building replaces your current screens, lists and flows. Save a project file first if you want to keep them.', 'Replace'))) return;
      m.close();
      H.setProject(st.project);
      H.toast(`Built ${st.project.screens.length} screens and ${st.project.flows.length} flow${st.project.flows.length === 1 ? '' : 's'}. Press Play to try it.`);
    };
  }
  drawWrite();
  // closing the dialog cancels a request in flight
  const obs = new MutationObserver(() => { if (!document.body.contains(m.el)) { st.busy?.abort(); obs.disconnect(); } });
  obs.observe(document.getElementById('modalRoot'), { childList: true });
}
