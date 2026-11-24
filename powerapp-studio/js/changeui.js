// changeui.js — the "Ask for a change" dialog: plain-English change → Claude edit operations → review → apply (one undo step).
import { MODELS, callClaude, AiError } from './ai.js';
import { clone } from './core.js';
import { lint } from './emit.js';
import { buildChangeRequest, reviewOps, applyOps } from './change.js';
import { store, KEY, MODEL, CHANGE_DRAFT } from './keystore.js';

let H = null; // { P, modal, toast, ui, esc, replaceProject, undo }
export function initChangeUI(helpers) { H = helpers; }

const EXAMPLES = [
  'Add a Cost Center text column to the form and the details screen.',
  'Add a Department choice column (Finance, IT, HR, Operations) and show it in the list rows.',
  'Email the approver every Monday at 8 about requests still pending.',
  'Only show admin buttons to group 11111111-2222-4333-8444-555555555555.',
];
const findingKey = (f) => `${f.code}|${f.where}|${f.msg}`;

export function changeDialog() {
  const esc = H.esc, ui = H.ui;
  if (!H.P().screens.some((s) => s.controls.length)) { H.toast('Build or describe an app first, then ask for changes', true); return; }
  const st = { text: store.get(CHANGE_DRAFT), result: null, rows: [], busy: null, usage: null };
  const m = H.modal(`<div class="modal-h"><h2>Ask for a change</h2><span class="kindtag">Uses Claude</span><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b" id="cBody"></div><div class="modal-f" id="cFoot"></div>`);
  const body = m.el.querySelector('#cBody'), foot = m.el.querySelector('#cFoot');
  const opts = (arr, v) => arr.map(([val, lab]) => `<option value="${esc(val)}" ${val === v ? 'selected' : ''}>${esc(lab)}</option>`).join('');

  function drawWrite() {
    const p = H.P();
    body.innerHTML = `<div class="flow-grid">
      <div><label class="f"><span>What should change in <b>${esc(p.app.name)}</b>? Columns, choices, fields on the form, screens, flows, options or the theme.</span>
        <textarea class="plain" id="cText" rows="9" placeholder="${esc(EXAMPLES[0])}">${esc(st.text)}</textarea></label>
        <div class="cg-exs">${EXAMPLES.map((x, i) => `<button class="cg-ex" data-ex="${i}">${esc(x)}</button>`).join('')}</div>
        <p class="note">Claude sees a summary of your lists, screens and flows (names and types, no data) and returns small edits. You review each one before anything changes. Screens the change does not touch stay exactly as they are, and one Undo puts everything back.</p></div>
      <div>
        <label class="f"><span>Anthropic API key</span><div class="rrow" style="margin:0"><input type="password" id="cKey" autocomplete="off" spellcheck="false" placeholder="sk-ant-…" value="${esc(store.get(KEY))}"><button class="btn sm" id="cShow" type="button">Show</button></div></label>
        <label class="chk"><input type="checkbox" id="cRem" ${store.remembered(KEY) ? 'checked' : ''}>Remember the key on this browser</label>
        <label class="f"><span>Model</span><select id="cModel">${opts(MODELS, store.get(MODEL) || MODELS[0][0])}</select></label>
        <div class="note-box"><b>Before you send</b><ul>
          <li>The request goes from this browser straight to Anthropic's API and is billed to your key.</li>
          <li>The key is never saved in project files or exports.</li>
          <li>Do not include sensitive, personal or controlled information.</li></ul></div>
      </div></div>`;
    foot.innerHTML = `<span class="grow">You review the edits before they are applied.</span><button class="btn" data-x>Cancel</button><button class="btn pri" id="cGo">Plan the change</button>`;
    const t = body.querySelector('#cText');
    t.addEventListener('input', () => { st.text = t.value; store.set(CHANGE_DRAFT, t.value, true); });
    body.querySelectorAll('[data-ex]').forEach((b) => { b.onclick = () => { t.value = EXAMPLES[+b.dataset.ex]; st.text = t.value; store.set(CHANGE_DRAFT, t.value, true); t.focus(); }; });
    body.querySelector('#cShow').onclick = (e) => { const k = body.querySelector('#cKey'); k.type = k.type === 'password' ? 'text' : 'password'; e.target.textContent = k.type === 'password' ? 'Show' : 'Hide'; };
    foot.querySelector('[data-x]').onclick = m.close;
    foot.querySelector('#cGo').onclick = run;
    setTimeout(() => t.focus(), 0);
  }

  async function run() {
    const key = (body.querySelector('#cKey')?.value ?? store.get(KEY)).trim();
    const remember = body.querySelector('#cRem')?.checked ?? store.remembered(KEY);
    const model = body.querySelector('#cModel')?.value || store.get(MODEL) || MODELS[0][0];
    if (!st.text.trim()) { H.toast('Describe the change first', true); return; }
    if (!key) { H.toast('Enter your Anthropic API key', true); return; }
    store.set(KEY, key, remember); store.set(MODEL, model, true);
    const ctrl = new AbortController(), t0 = Date.now();
    st.busy = ctrl;
    body.innerHTML = `<div class="ai-wait"><div class="ai-spin"></div><b>Planning the change…</b><span id="cTime">0 s</span><small>Claude is reading your project summary and working out the smallest set of edits. This usually takes 10–40 seconds.</small></div>`;
    foot.innerHTML = `<span class="grow"></span><button class="btn" id="cCancel">Cancel</button>`;
    foot.querySelector('#cCancel').onclick = () => ctrl.abort();
    const timer = setInterval(() => { const el = body.querySelector('#cTime'); if (el) el.textContent = Math.round((Date.now() - t0) / 1000) + ' s'; }, 500);
    try {
      const r = await callClaude({ key, model, request: buildChangeRequest(H.P(), st.text, model), signal: ctrl.signal });
      const plan = r.plan || {};
      st.result = { summary: String(plan.summary || ''), ops: Array.isArray(plan.ops) ? plan.ops : [], notes: Array.isArray(plan.notes) ? plan.notes.map(String) : [] };
      st.usage = r.usage;
      st.rows = reviewOps(H.P(), st.result.ops).map((x) => ({ ...x, on: !x.error }));
      drawReview();
    } catch (e) {
      drawWrite();
      H.toast(e instanceof AiError ? e.message : 'Something went wrong: ' + (e?.message || e), true);
    } finally { clearInterval(timer); st.busy = null; }
  }

  /** Dry-run the ticked operations on a copy so the review shows exactly what applying would do. */
  function dryRun() {
    const q = clone(H.P()), sel = st.rows.filter((r) => r.on), ops = sel.map((r) => clone(r.op));
    const res = applyOps(q, ops);
    res.skippedRows = new Map(res.skipped.map((x) => [sel[ops.indexOf(x.op)], x.reason]));
    return { q, res };
  }
  function drawReview() {
    const r = st.result, p = H.P();
    const { q, res } = dryRun();
    const names = [...res.touched].map((id) => q.screens.find((s) => s.id === id)?.name).filter(Boolean);
    const added = q.screens.filter((s) => !p.screens.some((x) => x.id === s.id)).map((s) => s.name);
    const changed = names.filter((n) => !added.includes(n));
    const skippedNow = res.skippedRows;
    body.innerHTML = `<div class="flow-grid"><div>
        ${r.summary ? `<p class="cg-sum">${esc(r.summary)}</p>` : ''}
        ${st.rows.length ? `<div class="cg-ops">${st.rows.map((row, i) => `<label class="cg-op ${row.error ? 'bad' : ''} ${row.destructive ? 'destr' : ''}">
          <input type="checkbox" data-op="${i}" ${row.on ? 'checked' : ''} ${row.error ? 'disabled' : ''}>
          <span><b>${esc(row.desc)}</b>${row.op?.reason ? `<small>${esc(row.op.reason)}</small>` : ''}
          ${row.error ? `<em class="cg-err">Skipped: ${esc(row.error)}</em>` : ''}
          ${row.destructive && !row.error ? `<em class="cg-warn">${row.op.op === 'setChoices' ? 'Items that use a removed choice keep it until edited.' : 'Formulas you wrote by hand that use the old name or column need a look afterwards.'}</em>` : ''}</span></label>`).join('')}</div>`
        : '<p class="note">Claude did not suggest any edits.</p>'}
      </div><div>
        <div class="note-box"><b>What applying does</b><ul>
          <li>${added.length ? `New screens: ${esc(added.join(', '))}` : 'No new screens'}</li>
          <li>${changed.length ? `Screens that change: ${esc(changed.join(', '))}` : 'No existing screen changes'}</li>
          <li>Every other screen stays exactly as it is. One Undo reverts the whole change.</li></ul></div>
        ${skippedNow.size ? `<div class="note-box warn"><b>Will be skipped</b><ul>${[...skippedNow].map(([x, why]) => `<li>${esc(x.desc)}: ${esc(why)}</li>`).join('')}</ul></div>` : ''}
        ${res.notes.length ? `<div class="note-box"><b>Good to know</b><ul>${res.notes.map((n) => `<li>${esc(n)}</li>`).join('')}</ul></div>` : ''}
        ${r.notes.length ? `<div class="note-box warn"><b>Claude's notes</b><ul>${r.notes.map((n) => `<li>${esc(n)}</li>`).join('')}</ul></div>` : ''}
        ${st.usage ? `<p class="note">Used ${Number(st.usage.input_tokens || 0).toLocaleString()} input and ${Number(st.usage.output_tokens || 0).toLocaleString()} output tokens.</p>` : ''}
      </div></div>`;
    const n = st.rows.filter((x) => x.on).length;
    foot.innerHTML = `<span class="grow">${n} of ${st.rows.length} edit${st.rows.length === 1 ? '' : 's'} selected</span><button class="btn" id="cBack">Change the request</button><button class="btn pri" id="cApply" ${n ? '' : 'disabled'}>Apply ${n} edit${n === 1 ? '' : 's'}</button>`;
    body.querySelectorAll('[data-op]').forEach((cb) => { cb.onchange = () => { st.rows[+cb.dataset.op].on = cb.checked; drawReview(); }; });
    foot.querySelector('#cBack').onclick = drawWrite;
    foot.querySelector('#cApply').onclick = apply;
  }
  function apply() {
    const before = new Set(lint(H.P()).map(findingKey));
    const q = clone(H.P());
    const res = applyOps(q, st.rows.filter((r) => r.on).map((r) => clone(r.op)));
    if (!res.applied.length) { H.toast('None of the selected edits could be applied', true); return; }
    H.replaceProject(q);
    const fresh = lint(H.P()).filter((f) => !before.has(findingKey(f)) && f.level !== 'info');
    body.innerHTML = `<div class="cg-done"><b>Applied ${res.applied.length} edit${res.applied.length === 1 ? '' : 's'}.</b>
      ${res.skipped.length ? `<div class="note-box warn"><b>Skipped</b><ul>${res.skipped.map((x) => `<li>${esc(x.op.op)}: ${esc(x.reason)}</li>`).join('')}</ul></div>` : ''}
      ${res.notes.length ? `<div class="note-box"><b>Good to know</b><ul>${res.notes.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
      ${fresh.length ? `<div class="note-box warn"><b>New checks to look at</b><ul>${fresh.slice(0, 12).map((f) => `<li><code>${esc(f.code)}</code> ${esc(f.where)}: ${esc(f.msg)}</li>`).join('')}</ul></div>` : '<p class="note">No new problems found by the checks.</p>'}</div>`;
    foot.innerHTML = `<span class="grow">Press Play to try it, or Get YAML to paste the changed screens.</span><button class="btn" id="cUndo">${ui('undo')}Undo the change</button><button class="btn" id="cMore">Ask for another change</button><button class="btn pri" data-x>Done</button>`;
    foot.querySelector('#cUndo').onclick = () => { H.undo(); m.close(); H.toast('Change undone'); };
    foot.querySelector('#cMore').onclick = () => { st.text = ''; store.set(CHANGE_DRAFT, '', true); drawWrite(); };
    foot.querySelector('[data-x]').onclick = m.close;
    store.set(CHANGE_DRAFT, '', true);
    H.toast(`Applied ${res.applied.length} edit${res.applied.length === 1 ? '' : 's'}`);
  }
  drawWrite();
  const obs = new MutationObserver(() => { if (!document.body.contains(m.el)) { st.busy?.abort(); obs.disconnect(); } });
  obs.observe(document.getElementById('modalRoot'), { childList: true });
}
