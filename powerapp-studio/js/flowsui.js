// flowsui.js — Flows tab and flow editor. Kept apart from ui.js, which passes in its helpers.
import { FLOW_KINDS, TIME_ZONES, newFlow, flowDefinition, flowPackageFiles, flowFileName, flowFxName, recipientsText, sampleFromPackage, internalName } from './flows.js';
import { sampleRows, displayValue, clone, uid } from './core.js';

let H = null; // helpers from ui.js: { P, modal, toast, ui, esc, commit, renderPanel, download, copyText, ask }
export function initFlowsUI(helpers) { H = helpers; }

const kindLabel = (k) => FLOW_KINDS[k]?.label || k;
const listOf = (f) => H.P().lists.find((l) => l.name === f.list);

// ------------------------------------------------------------------ panel
export function panelFlows() {
  const P = H.P(), esc = H.esc, ui = H.ui;
  const flows = P.flows || [];
  const sample = P.app.flowSample;
  return `<div class="ph"><h2>Flows</h2><button class="btn sm" data-cmd="addFlow">${ui('plus')}Flow</button></div><div class="pb">
    <p class="note">Power Automate flows for this app: emails on new items, on status changes, daily reminders, and flows a button runs. Each one exports as a package for <b>Import Package (Legacy)</b>.</p>
    <label class="f"><span>SharePoint site URL</span><input data-app="siteUrl" value="${esc(P.app.siteUrl)}" placeholder="https://contoso.sharepoint.com/sites/Team"></label>
    <label class="f"><span>App web link <small>(Apps → Details → Web link)</small></span><input data-app="appLink" value="${esc(P.app.appLink)}" placeholder="https://apps.powerapps.com/play/e/…/a/…?tenantId=…"></label>
    ${flows.length ? flows.map((f) => `<div class="flowcard ${f.enabled === false ? 'off' : ''}" data-fid="${esc(f.id)}">
      <div class="fc-h"><i>${ui(FLOW_KINDS[f.kind]?.icon || 'send')}</i><b>${esc(f.name)}</b></div>
      <small>${esc(kindLabel(f.kind))} · ${esc(f.list || '(no list)')}${f.enabled === false ? ' · not exported' : ''}</small>
      <small>To: ${esc(recipientsText(f))}</small>
      <div class="btns"><button class="btn sm" data-cmd="editFlow" data-fid="${esc(f.id)}">Edit</button><button class="btn sm" data-cmd="dlFlow" data-fid="${esc(f.id)}">${ui('download')}.zip</button><button class="mini" data-cmd="dupFlow" data-fid="${esc(f.id)}" title="Duplicate">${ui('copy')}</button><button class="mini danger" data-cmd="delFlow" data-fid="${esc(f.id)}" title="Delete">${ui('trash')}</button></div>
    </div>`).join('') : '<p class="note"><b>No flows yet.</b> Add one to email people when items are created or change.</p>'}
    <div class="sub">Tenant sample (optional)</div>
    <p class="note">${sample ? '✓ Using connector details from your exported flow.' : 'Packages use standard connector details. To copy your tenant\'s exactly, load any flow you exported as a package (.zip) that uses SharePoint or Outlook.'}</p>
    <div class="btns"><button class="btn sm" data-cmd="flowSample">${ui('upload')}${sample ? 'Replace sample' : 'Load an exported flow'}</button>${sample ? '<button class="btn sm danger" data-cmd="flowSampleClear">Clear</button>' : ''}</div>
  </div>`;
}

// ------------------------------------------------------------------ add / edit
export function addFlowDialog() {
  const P = H.P(), esc = H.esc;
  if (!P.lists.length) { H.toast('Add a list on the Data tab first', true); return; }
  let kind = 'created';
  const m = H.modal(`<div class="modal-h"><h2>New flow</h2><button class="mini" data-close>${H.ui('close')}</button></div>
    <div class="modal-b"><div class="scen" id="fkinds" style="grid-template-columns:repeat(2,minmax(0,1fr))">${Object.entries(FLOW_KINDS).map(([k, d]) => `<button data-kind="${k}" class="${k === kind ? 'on' : ''}"><i>${H.svgIcon(d.icon, '#fff', 2)}</i><b>${esc(d.label)}</b><small>${esc(d.blurb)}</small></button>`).join('')}</div>
    <label class="f" style="margin-top:14px"><span>List</span><select id="flist">${P.lists.map((l) => `<option>${esc(l.name)}</option>`).join('')}</select></label></div>
    <div class="modal-f"><button class="btn" data-close>Cancel</button><button class="btn pri" id="fgo">Create flow</button></div>`, 'md');
  m.el.querySelector('#fkinds').addEventListener('click', (e) => { const b = e.target.closest('[data-kind]'); if (!b) return; kind = b.dataset.kind; m.el.querySelectorAll('[data-kind]').forEach((x) => x.classList.toggle('on', x === b)); });
  m.el.querySelector('#fgo').onclick = () => {
    const f = newFlow(P, kind, m.el.querySelector('#flist').value);
    if (kind === 'button') f.name = f.name.replace(/(^|\s)(\w)/g, (_, s, c) => c.toUpperCase()).replace(/\s+/g, '');
    let n = f.name, i = 2; while ((P.flows || []).some((x) => flowFxName(x.name) === flowFxName(n))) n = f.name + ' ' + i++;
    f.name = n;
    (P.flows = P.flows || []).push(f);
    m.close(); H.commit(); H.renderPanel(); flowDialog(f.id);
  };
}

export function flowDialog(id) {
  const P = H.P(), esc = H.esc, ui = H.ui;
  const f = (P.flows || []).find((x) => x.id === id); if (!f) return;
  let tab = 'preview', lastField = 'body';
  const m = H.modal(`<div class="modal-h"><h2>Flow</h2><span class="kindtag">${esc(kindLabel(f.kind))}</span><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b"><div class="flow-grid"><div id="fform"></div><div><div class="seg" id="ftabs" style="margin-bottom:10px"><button data-ft="preview" class="on">Email preview</button><button data-ft="json">Flow definition</button></div><div id="fright"></div></div></div></div>
    <div class="modal-f"><span class="grow" id="fnote"></span><button class="btn danger" id="fdel">Delete</button><button class="btn" id="fcopy">${ui('copy')}Copy definition</button><button class="btn" id="fdl">${ui('download')}Download package</button><button class="btn pri" data-close>Done</button></div>`);
  const form = m.el.querySelector('#fform'), right = m.el.querySelector('#fright');
  const F = (label, inner) => `<label class="f"><span>${label}</span>${inner}</label>`;
  const opts = (arr, v) => arr.map((o) => { const [val, lab] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(val)}" ${String(val) === String(v ?? '') ? 'selected' : ''}>${esc(lab)}</option>`; }).join('');
  const tokens = () => {
    const l = listOf(f); if (!l) return [];
    const t = l.columns.map((c) => internalName(c.name));
    t.push('ID', 'Created', 'Author', 'Link');
    l.columns.filter((c) => c.type === 'Person').forEach((c) => t.push(internalName(c.name) + '.Email'));
    if (f.kind === 'button') t.push('Me');
    return t;
  };
  const drawForm = () => {
    const l = listOf(f);
    const choiceCols = l ? l.columns.filter((c) => c.type === 'Choice') : [];
    const personCols = l ? l.columns.filter((c) => c.type === 'Person') : [];
    const st = l?.columns.find((c) => c.name === f.statusField);
    let kindHtml = '';
    if (f.kind === 'statusChanged') {
      kindHtml = F('Column to watch', `<select data-f="statusField" data-re="1">${opts([['', '(pick a Choice column)'], ...choiceCols.map((c) => c.name)], f.statusField)}</select>`) +
        (st?.choices?.length ? `<div class="f"><span>Send when it changes to <small>(none ticked = any change)</small></span><div class="chips">${st.choices.map((c) => `<label class="chk" style="margin:0"><input type="checkbox" data-fsv="${esc(c)}" ${f.statusValues.includes(c) ? 'checked' : ''}>${esc(c)}</label>`).join('')}</div></div>` : '');
    }
    if (f.kind === 'reminder') {
      const rc = l?.columns.find((c) => c.name === f.statusField);
      kindHtml = `<div class="grid2">${F('Items where', `<select data-f="statusField" data-re="1">${opts([['', '(any item)'], ...(l?.columns || []).filter((c) => ['Choice', 'Text', 'YesNo', 'Number'].includes(c.type)).map((c) => c.name)], f.statusField)}</select>`)}
        ${f.statusField ? F('equals', rc?.type === 'Choice' ? `<select data-f="reminderValue">${opts(rc.choices, f.reminderValue)}</select>` : rc?.type === 'YesNo' ? `<select data-f="reminderValue">${opts(['true', 'false'], f.reminderValue)}</select>` : `<input data-f="reminderValue" value="${esc(f.reminderValue)}">`) : '<span></span>'}</div>
        <div class="grid2">${F('and created at least (days ago)', `<input type="number" min="0" max="365" data-f="days" value="${esc(f.days)}">`)}${F('Send at', `<select data-f="hour">${opts(Array.from({ length: 24 }, (_, h) => [h, `${(h % 12) || 12}:00 ${h < 12 ? 'AM' : 'PM'}`]), f.hour)}</select>`)}</div>
        ${F('Time zone', `<select data-f="timeZone">${opts(TIME_ZONES, f.timeZone)}</select>`)}`;
    }
    if (f.kind === 'button') kindHtml = `<p class="note">In the designer, give a button the action <b>Run a flow → ${esc(f.name)}</b>. It sends the selected ${esc(f.list || 'item')}'s ID; the app waits for the flow, then shows your message. In Studio the flow appears as <code>${esc(flowFxName(f.name))}</code>.</p>`;
    const rkinds = [['column', 'Person column'], ['author', 'Created By'], ['editor', 'Modified By'], ...(f.kind === 'button' ? [['me', 'Person who pressed the button']] : []), ['address', 'Email address']];
    form.innerHTML = `${F('Flow name', `<input data-f="name" value="${esc(f.name)}">`)}
      ${F('List', `<select data-f="list" data-re="1">${opts(P.lists.map((x) => x.name), f.list)}</select>`)}
      ${kindHtml}
      <div class="f"><span>Send to</span>${f.recipients.map((r, i) => `<div class="rrow"><select data-ri="${i}" data-rk="1">${opts(rkinds, r.kind)}</select>${r.kind === 'column' ? `<select data-ri="${i}" data-rv="1">${opts([['', '(column)'], ...personCols.map((c) => c.name)], r.value)}</select>` : r.kind === 'address' ? `<input data-ri="${i}" data-rv="1" value="${esc(r.value)}" placeholder="name@agency.mil; team@agency.mil">` : '<span></span>'}<button class="mini danger" data-rdel="${i}" title="Remove">${ui('close')}</button></div>`).join('')}
        <button class="btn sm" data-radd="1">${ui('plus')}Recipient</button></div>
      ${F('Subject', `<input data-f="subject" value="${esc(f.subject)}">`)}
      ${F('Message', `<textarea class="plain" rows="7" data-f="body">${esc(f.body)}</textarea>`)}
      <div class="f"><span>Insert a value <small>(click to add at the cursor)</small></span><div class="chips">${tokens().map((t) => `<button type="button" class="chip" data-tok="${esc(t)}">{${esc(t)}}</button>`).join('')}</div></div>
      <label class="chk"><input type="checkbox" data-f="enabled" ${f.enabled !== false ? 'checked' : ''}>Include in exports</label>`;
  };
  const sampleRow = () => { const l = listOf(f); return l ? sampleRows(l, P.lists, 2)[0] : {}; };
  const fill = (text) => {
    const row = sampleRow(), l = listOf(f);
    return String(text || '').replace(/\{([A-Za-z_][A-Za-z0-9_]*)(\.Email)?\}/g, (m, name, em) => {
      if (name === 'Link') return P.app.appLink ? `${P.app.appLink}${P.app.appLink.includes('?') ? '&' : '?'}ID=${row.ID}` : `(item ${row.ID})`;
      if (name === 'Author' || name === 'Editor' || name === 'Me') return em ? 'alex.rivera@contoso.com' : 'Alex Rivera';
      if (name === 'ID') return String(row.ID);
      if (name === 'Created' || name === 'Modified') return displayValue({ type: 'DateTime' }, row[name]);
      const col = l?.columns.find((c) => c.name === name || internalName(c.name) === name);
      if (!col) return m;
      if (col.type === 'Person') return em ? (row[col.name]?.Mail || '') : (row[col.name]?.DisplayName || '');
      return displayValue(col, row[col.name]);
    });
  };
  const drawRight = () => {
    m.el.querySelectorAll('[data-ft]').forEach((b) => b.classList.toggle('on', b.dataset.ft === tab));
    if (tab === 'json') { right.innerHTML = `<pre class="code" style="max-height:58vh">${esc(JSON.stringify(flowDefinition(P, f), null, 2))}</pre>`; return; }
    const to = f.recipients.map((r) => r.kind === 'address' ? r.value : r.kind === 'column' ? fill(`{${internalName(r.value)}.Email}`) : r.kind === 'me' ? 'you' : 'alex.rivera@contoso.com').filter(Boolean).join('; ');
    const when = { created: `A new ${f.list || 'item'} is created`, statusChanged: `${f.statusField || 'A column'} changes${f.statusValues.length ? ' to ' + f.statusValues.join(' or ') : ''}`, reminder: `Every day at ${(f.hour % 12) || 12}:00 ${f.hour < 12 ? 'AM' : 'PM'} for each matching ${f.list || 'item'}`, button: 'A button in the app runs it' }[f.kind];
    right.innerHTML = `<div class="mail"><div class="mail-when">${ui('clock')}<span>${esc(when)}</span></div>
      <div class="mail-h"><div><b>To</b> ${esc(to || '(nobody yet)')}</div><div><b>Subject</b> ${esc(fill(f.subject))}</div></div>
      <div class="mail-b">${esc(fill(f.body)).replace(/\n/g, '<br>').replace(/(https?:\/\/[^\s<]+)/g, '<a>Open in ' + esc(P.app.name) + '</a>')}</div></div>
      <p class="note">Preview uses the first sample row. Values come from the item that triggered the flow.</p>`;
  };
  const changed = (re) => { H.commit(); H.renderPanel(); if (re) drawForm(); drawRight(); };
  form.addEventListener('focusin', (e) => { const k = e.target.dataset?.f; if (k === 'subject' || k === 'body') lastField = k; });
  form.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.f) {
      const k = t.dataset.f;
      f[k] = t.type === 'checkbox' ? t.checked : (k === 'days' || k === 'hour') ? Number(t.value) || 0 : t.value;
      if (k === 'list') { f.statusField = ''; f.statusValues = []; f.recipients = f.recipients.filter((r) => r.kind !== 'column'); }
      if (k === 'statusField') { f.statusValues = []; f.reminderValue = ''; }
      changed(!!t.dataset.re);
      return;
    }
    if (t.dataset.fsv != null) { const v = t.dataset.fsv; f.statusValues = t.checked ? [...new Set([...f.statusValues, v])] : f.statusValues.filter((x) => x !== v); changed(false); return; }
    if (t.dataset.ri != null) { const r = f.recipients[+t.dataset.ri]; if (t.dataset.rk) { r.kind = t.value; r.value = ''; changed(true); } else { r.value = t.value; changed(false); } }
  });
  form.addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.radd) { f.recipients.push({ kind: 'address', value: '' }); changed(true); }
    if (b.dataset.rdel != null) { f.recipients.splice(+b.dataset.rdel, 1); changed(true); }
    if (b.dataset.tok) {
      const el = form.querySelector(`[data-f="${lastField}"]`), ins = `{${b.dataset.tok}}`;
      const s = el.selectionStart ?? el.value.length, en = el.selectionEnd ?? s;
      el.value = el.value.slice(0, s) + ins + el.value.slice(en); f[lastField] = el.value;
      el.focus(); el.setSelectionRange(s + ins.length, s + ins.length); changed(false);
    }
  });
  m.el.querySelector('#ftabs').addEventListener('click', (e) => { const b = e.target.closest('[data-ft]'); if (b) { tab = b.dataset.ft; drawRight(); } });
  m.el.querySelector('#fcopy').onclick = () => H.copyText(JSON.stringify(flowDefinition(P, f), null, 2));
  m.el.querySelector('#fdl').onclick = () => downloadFlow(f);
  m.el.querySelector('#fdel').onclick = async () => { if (await H.ask('Delete flow?', `Delete "${f.name}"? Buttons that run it will show lint errors.`, 'Delete')) { P.flows.splice(P.flows.indexOf(f), 1); m.close(); H.commit(); H.renderPanel(); } };
  m.el.querySelector('#fnote').textContent = 'Import: My flows → Import → Import Package (Legacy) → Create as new.';
  drawForm(); drawRight();
}

// ------------------------------------------------------------------ files
export async function flowZipBlob(f) {
  const z = new window.JSZip();
  for (const [p, t] of Object.entries(flowPackageFiles(H.P(), f))) z.file(p, t);
  return z.generateAsync({ type: 'blob' });
}
export async function downloadFlow(f) {
  if (!window.JSZip) { H.toast('Zip library did not load', true); return; }
  H.download(flowFileName(f), await flowZipBlob(f));
}
export function duplicateFlow(id) {
  const P = H.P(), f = P.flows.find((x) => x.id === id); if (!f) return;
  const c = clone(f); c.id = uid(); let n = f.name + ' copy', i = 2; while (P.flows.some((x) => flowFxName(x.name) === flowFxName(n))) n = f.name + ' copy ' + i++; c.name = n;
  P.flows.splice(P.flows.indexOf(f) + 1, 0, c); H.commit(); H.renderPanel();
}
export async function deleteFlow(id) {
  const P = H.P(), f = P.flows.find((x) => x.id === id); if (!f) return;
  if (await H.ask('Delete flow?', `Delete "${f.name}"?`, 'Delete')) { P.flows.splice(P.flows.indexOf(f), 1); H.commit(); H.renderPanel(); }
}
export function loadFlowSample() {
  const fi = document.createElement('input'); fi.type = 'file'; fi.accept = '.zip';
  fi.onchange = async () => {
    try {
      const z = await window.JSZip.loadAsync(await fi.files[0].arrayBuffer()), files = {};
      for (const [p, e] of Object.entries(z.files)) if (!e.dir && /\.json$/i.test(p)) files[p.replace(/^\/+/, '')] = await e.async('string');
      if (!files['manifest.json']) throw new Error('no manifest.json at the top of the zip');
      const s = sampleFromPackage(files);
      if (!Object.keys(s).length) throw new Error('it uses neither SharePoint nor Office 365 Outlook');
      H.P().app.flowSample = s; H.commit(); H.renderPanel(); H.toast('Loaded connector details from ' + fi.files[0].name);
    } catch (e) { H.toast('Could not read that package: ' + e.message, true); }
  };
  fi.click();
}
