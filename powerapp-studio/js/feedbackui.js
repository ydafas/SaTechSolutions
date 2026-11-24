// feedbackui.js — "Report a paste error": paste what Studio or Power Automate said, get the control,
// the likely cause and the fix, switch off properties the tenant rejects, and keep a findings log.
import { analyze, applyAction, logFindings, exportFindings } from './feedback.js';

let H = null; // { P, modal, toast, ui, esc, commit, persist, selectControl, openScreen, openDataTab, copyText, download, slug }
export function initFeedbackUI(helpers) { H = helpers; }

const FAMILY_LABEL = {
  unknownProperty: 'Property', duplicateName: 'Name', controlVersion: 'Control', parse: 'YAML', thisItem: 'Formula', missingFormulas: 'Formulas',
  missingDataSource: 'Data', missingConnector: 'Connector', missingFlow: 'Flow', missingScreen: 'Screen', missingVariable: 'Variable',
  missingControl: 'Control', unknownName: 'Name', choiceAsText: 'Type', argType: 'Type', columnName: 'Column', delegation: 'Delegation',
  flowConnection: 'Flow', flowTemplate: 'Flow', flowImport: 'Flow', unknown: 'Other',
};

export function feedbackDialog(opts = {}) {
  const esc = H.esc, ui = H.ui;
  const st = { tab: opts.tab || 'report', text: opts.text || '', results: null, done: new Set() };
  const m = H.modal(`<div class="modal-h"><h2>Report a paste error</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="fb-tabs" role="tablist"><button role="tab" data-fbt="report">Explain an error</button><button role="tab" data-fbt="log">Tenant log <b id="fbN"></b></button></div>
    <div class="modal-b" id="fbBody"></div><div class="modal-f" id="fbFoot"></div>`, 'md');
  const body = m.el.querySelector('#fbBody'), foot = m.el.querySelector('#fbFoot');
  m.el.querySelector('.fb-tabs').onclick = (e) => { const b = e.target.closest('[data-fbt]'); if (b) { st.tab = b.dataset.fbt; draw(); } };

  function draw() {
    const P = H.P(), n = (P.findings || []).length;
    m.el.querySelector('#fbN').textContent = n ? String(n) : '';
    m.el.querySelectorAll('[data-fbt]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.fbt === st.tab)));
    st.tab === 'log' ? drawLog() : drawReport();
  }

  function card(f, i) {
    const acts = f.actions.map((a, k) => `<button class="btn sm ${a.kind === 'suppress' || a.kind === 'largeList' ? 'pri' : ''}" data-act="${i}|${k}" ${st.done.has(`${i}|${k}`) ? 'disabled' : ''}>${esc(st.done.has(`${i}|${k}`) ? 'Done' : a.label)}</button>`).join('');
    return `<div class="fb-card">
      <div class="fb-top"><span class="fb-fam">${esc(f.code || FAMILY_LABEL[f.family] || 'Other')}</span><b>${esc(f.title)}</b></div>
      ${f.where ? `<div class="fb-where">${ui('screen')}<span>${esc(f.where)}</span></div>` : ''}
      <p><span class="muted">Why:</span> ${esc(f.cause)}</p>
      <p><span class="muted">Fix:</span> ${esc(f.fix)}</p>
      ${acts ? `<div class="btns">${acts}</div>` : ''}
      <details><summary>What you pasted</summary><pre>${esc(f.excerpt)}</pre></details></div>`;
  }

  function drawReport() {
    body.innerHTML = `<p class="note" style="margin-top:0">Paste the error text from Power Apps Studio (the red banner, the formula bar tooltip, or App checker) or from Power Automate's import page. Several messages at once are fine.</p>
      <textarea class="inp fb-in" id="fbText" rows="6" placeholder="Error PA2108: Unknown property 'IconBackground' on control 'dtpNeededBy'.&#10;Name isn't valid. 'Requests' isn't recognized.">${esc(st.text)}</textarea>
      <div id="fbRes">${st.results ? (st.results.length ? st.results.map(card).join('') : '<div class="note-box">Nothing to explain: paste the message text.</div>') : ''}</div>`;
    foot.innerHTML = `<span class="grow">Each explained error is added to the tenant log.</span><button class="btn" data-x>Close</button><button class="btn pri" id="fbGo">Explain</button>`;
    foot.querySelector('[data-x]').onclick = m.close;
    const ta = body.querySelector('#fbText');
    ta.oninput = () => { st.text = ta.value; };
    foot.querySelector('#fbGo').onclick = () => {
      const P = H.P();
      st.results = analyze(P, st.text); st.done = new Set();
      st.ids = [];
      if (st.results.length) { logFindings(P, st.results); st.ids = P.findings.slice(-st.results.length).map((f) => f.id); H.persist(); }
      draw();
      body.querySelector('.fb-card')?.scrollIntoView({ block: 'nearest' });
    };
    body.querySelector('#fbRes').onclick = (e) => {
      const b = e.target.closest('[data-act]'); if (!b) return;
      const [i, k] = b.dataset.act.split('|').map(Number), a = st.results[i].actions[k];
      run(a, `${i}|${k}`, i);
    };
  }

  function run(a, key, i) {
    const P = H.P();
    if (a.kind === 'select') { m.close(); H.selectControl(a.screenId, a.ctlId); return; }
    if (a.kind === 'screen') { m.close(); H.openScreen(a.screenId); return; }
    if (a.kind === 'dataTab') { m.close(); H.openDataTab(a.list); return; }
    const msg = applyAction(P, a);
    if (!msg) return;
    const f = (P.findings || []).find((x) => x.id === st.ids?.[i]);
    if (f) f.status = 'fixed'; // the logged finding this action answers
    st.done.add(key);
    H.commit();
    H.toast(msg + (a.kind === 'suppress' ? ' Copy the screen again and paste it.' : ''));
    draw();
  }

  function drawLog() {
    const P = H.P(), fs = [...(P.findings || [])].reverse(), sp = Object.entries(P.app.suppressProps || {});
    body.innerHTML = `<p class="note" style="margin-top:0">What this tenant rejected, kept with the project. Export it and send it along when a paste keeps failing: it tells us which rule to add.</p>
      <h3 class="fb-h">Properties switched off</h3>
      ${sp.length ? `<div class="fb-sup">${sp.map(([ct, props]) => props.map((pr) => `<div class="fb-row"><code>${esc(ct)}</code><span>${esc(pr)}</span><button class="btn sm" data-uns="${esc(ct)}|${esc(pr)}">Emit again</button></div>`).join('')).join('')}</div>`
        : '<p class="note">None. Generated YAML includes every property.</p>'}
      <h3 class="fb-h">Findings</h3>
      ${fs.length ? fs.map((f) => `<div class="fb-log ${f.status}"><div><b>${esc(f.code || FAMILY_LABEL[f.family] || 'Other')}</b> ${esc(f.title)}<small>${esc(new Date(f.at).toLocaleString())}${f.where ? ' · ' + esc(f.where) : ''}</small></div>
          <label class="chk"><input type="checkbox" data-fx="${esc(f.id)}" ${f.status === 'fixed' ? 'checked' : ''}>Fixed</label></div>`).join('')
        : '<p class="note">Nothing logged yet.</p>'}`;
    foot.innerHTML = `${fs.length ? '<button class="btn danger" id="fbClr">Clear log</button>' : ''}<span class="grow"></span><button class="btn" id="fbCopy">Copy as text</button><button class="btn pri" id="fbDl">Download .txt</button>`;
    foot.querySelector('#fbCopy').onclick = () => H.copyText(exportFindings(H.P()));
    foot.querySelector('#fbDl').onclick = () => H.download(`${H.slug(H.P().app.name) || 'app'}_tenant-findings.txt`, exportFindings(H.P()));
    const clr = foot.querySelector('#fbClr');
    if (clr) clr.onclick = () => { delete H.P().findings; H.persist(); draw(); };
    body.onclick = (e) => {
      const b = e.target.closest('[data-uns]'); if (!b) return;
      const [control, prop] = b.dataset.uns.split('|');
      H.toast(applyAction(H.P(), { kind: 'unsuppress', control, prop }));
      H.commit(); draw();
    };
    body.onchange = (e) => {
      const t = e.target; if (!t.dataset.fx) return;
      const f = (H.P().findings || []).find((x) => x.id === t.dataset.fx);
      if (f) { f.status = t.checked ? 'fixed' : 'open'; H.persist(); t.closest('.fb-log').className = `fb-log ${f.status}`; }
    };
  }
  draw();
  if (st.tab === 'report') setTimeout(() => body.querySelector('#fbText')?.focus(), 0);
}
