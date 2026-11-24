// tableui.js — Data tab "Import from CSV/Excel": file → preview with editable names and types → a new list with sample rows.
import { COLTYPES, slug } from './core.js';
import { parseCSV, readXlsx, inferColumns, listFromTable } from './tabular.js';
import { genEdit, genDetail, genBrowse } from './templates.js';

let H = null; // { P, modal, toast, ui, esc, commit, renderAll }
export function initTableUI(helpers) { H = helpers; }

export function importTableDialog() {
  const esc = H.esc, ui = H.ui;
  const st = { file: '', table: null, headerRow: true, cols: [], titleIndex: 0, name: '', screens: true };
  const m = H.modal(`<div class="modal-h"><h2>Import from CSV or Excel</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b" id="tBody"></div><div class="modal-f" id="tFoot"></div>`, 'wide');
  const body = m.el.querySelector('#tBody'), foot = m.el.querySelector('#tFoot');

  function drawPick() {
    body.innerHTML = `<label class="drop" id="tDrop"><input type="file" id="tFile" accept=".csv,.tsv,.txt,.xlsx" hidden>
        <b>Choose a .csv or .xlsx file</b><small>or drop it here. The file is read in this browser and never uploaded.</small></label>
      <label class="f"><span>…or paste CSV</span><textarea class="plain" id="tPaste" rows="6" placeholder="Title,Owner,Status&#10;Laptop refresh,ana@contoso.com,Open"></textarea></label>
      <p class="note">Columns and types are guessed from the values; you can change them on the next step. Up to 200 rows are kept as sample data for Play.</p>`;
    foot.innerHTML = `<span class="grow"></span><button class="btn" data-x>Cancel</button><button class="btn pri" id="tNext">Preview</button>`;
    const fi = body.querySelector('#tFile'), drop = body.querySelector('#tDrop');
    fi.onchange = () => fi.files[0] && load(fi.files[0]);
    drop.ondragover = (e) => { e.preventDefault(); drop.classList.add('over'); };
    drop.ondragleave = () => drop.classList.remove('over');
    drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) load(f); };
    foot.querySelector('[data-x]').onclick = m.close;
    foot.querySelector('#tNext').onclick = () => { const t = body.querySelector('#tPaste').value; if (!t.trim()) { H.toast('Choose a file or paste CSV', true); return; } setTable(parseCSV(t), 'Imported'); };
  }
  async function load(f) {
    try {
      if (/\.xlsx$/i.test(f.name)) {
        if (!window.JSZip) { H.toast('Zip library did not load', true); return; }
        setTable(await readXlsx(await f.arrayBuffer(), window.JSZip), f.name);
      } else if (/\.xls$/i.test(f.name)) H.toast('Save the workbook as .xlsx or .csv first', true);
      else setTable(parseCSV(await f.text()), f.name);
    } catch (e) { H.toast('Could not read the file: ' + (e?.message || e), true); }
  }
  function setTable(rows, fileName) {
    if (!rows.length) { H.toast('The file has no rows', true); return; }
    st.table = rows; st.file = fileName;
    st.name = String(fileName).replace(/\.[^.]+$/, '').replace(/[^A-Za-z0-9]+(.)?/g, (_, c) => (c ? c.toUpperCase() : '')).replace(/^[^A-Za-z]+/, '') || 'ImportedList';
    infer(); drawPreview();
  }
  const data = () => (st.headerRow ? st.table.slice(1) : st.table);
  function infer() {
    const width = Math.max(...st.table.map((r) => r.length));
    const header = st.headerRow ? Array.from({ length: width }, (_, i) => st.table[0][i] ?? '') : Array.from({ length: width }, (_, i) => `Column${i + 1}`);
    const info = inferColumns(header, data());
    st.cols = info.map((c, i) => ({ ...c, index: i, include: c.filled > 0, name: c.header.replace(/[^A-Za-z0-9]+(.)?/g, (_, x) => (x ? x.toUpperCase() : '')).replace(/^[^A-Za-z]+/, '') || `Column${i + 1}` }));
    const t = st.cols.find((c) => /^(title|name|subject|item)$/i.test(c.header) && c.type === 'Text') || st.cols.find((c) => c.type === 'Text');
    st.titleIndex = t ? t.index : -1;
  }
  function drawPreview() {
    const n = data().length;
    body.innerHTML = `<div class="imp-top">
        <label class="f"><span>List name</span><input id="tName" value="${esc(st.name)}"></label>
        <label class="f"><span>Title column <small>(SharePoint's required first column)</small></span><select id="tTitle"><option value="-1">(leave Title empty)</option>${st.cols.map((c) => `<option value="${c.index}" ${c.index === st.titleIndex ? 'selected' : ''}>${esc(c.header)}</option>`).join('')}</select></label>
        <label class="chk"><input type="checkbox" id="tHead" ${st.headerRow ? 'checked' : ''}>First row has the column names</label>
        <label class="chk"><input type="checkbox" id="tScr" ${st.screens ? 'checked' : ''}>Also add screens (list, details and form)</label></div>
      ${n > 200 ? `<div class="note-box warn">This file has ${n} rows. The first 200 are kept as sample data for Play; import the full file straight into SharePoint (see Export rows).</div>` : ''}
      <div class="imp-wrap"><table class="imp-t"><thead><tr><th></th><th>From the file</th><th>Column name</th><th>Type</th><th>Values</th></tr></thead><tbody>
      ${st.cols.map((c, i) => `<tr class="${c.index === st.titleIndex ? 'is-title' : ''}" data-i="${i}">
        <td><input type="checkbox" data-k="include" ${c.include || c.index === st.titleIndex ? 'checked' : ''} ${c.index === st.titleIndex ? 'disabled' : ''} aria-label="Import ${esc(c.header)}"></td>
        <td>${esc(c.header)}</td>
        <td>${c.index === st.titleIndex ? '<span class="lock">Title</span>' : `<input class="inp" data-k="name" value="${esc(c.name)}" aria-label="Name for ${esc(c.header)}">`}</td>
        <td>${c.index === st.titleIndex ? '<span class="lock">Text</span>' : `<select class="inp" data-k="type" aria-label="Type for ${esc(c.header)}">${Object.entries(COLTYPES).filter(([k]) => k !== 'Lookup').map(([k, v]) => `<option value="${k}" ${k === c.type ? 'selected' : ''}>${v.label}</option>`).join('')}</select>`}</td>
        <td class="imp-s">${esc(c.sample.join(' · ').slice(0, 80))}${c.type === 'Choice' ? `<small>${c.choices.length} choices</small>` : ''}</td></tr>`).join('')}
      </tbody></table></div>
      <p class="note">${n} data row${n === 1 ? '' : 's'} from ${esc(st.file)}. Names are cleaned to letters and digits; names SharePoint already uses get the list's name in front.</p>`;
    foot.innerHTML = `<span class="grow"></span><button class="btn" id="tBack">Back</button><button class="btn pri" id="tGo">Create the list</button>`;
    body.querySelector('#tName').oninput = (e) => { st.name = e.target.value; };
    body.querySelector('#tTitle').onchange = (e) => { st.titleIndex = +e.target.value; drawPreview(); };
    body.querySelector('#tHead').onchange = (e) => { st.headerRow = e.target.checked; infer(); drawPreview(); };
    body.querySelector('#tScr').onchange = (e) => { st.screens = e.target.checked; };
    body.querySelector('tbody').onchange = body.querySelector('tbody').oninput = (e) => {
      const tr = e.target.closest('[data-i]'); if (!tr) return;
      const c = st.cols[+tr.dataset.i], k = e.target.dataset.k;
      if (k === 'include') c.include = e.target.checked;
      if (k === 'name') c.name = e.target.value;
      if (k === 'type') { c.type = e.target.value; if (c.type === 'Choice' && !c.choices.length) c.choices = [...new Set(data().map((r) => String(r[c.index] ?? '').trim()).filter(Boolean))].slice(0, 50); }
    };
    foot.querySelector('#tBack').onclick = drawPick;
    foot.querySelector('#tGo').onclick = create;
  }
  function create() {
    const P = H.P();
    const { list, notes } = listFromTable(P, { name: st.name, rows: data(), cols: st.cols, titleIndex: st.titleIndex });
    if (list.columns.length < 2 && !list.rows.some((r) => r.Title)) { H.toast('Pick at least one column to import', true); return; }
    P.lists.push(list);
    let added = 0;
    const starter = P.screens.length === 1 && !P.screens[0].controls.length ? P.screens[0] : null;
    if (st.screens) {
      const ed = genEdit(P, list), det = genDetail(P, list, { editScreen: ed });
      genBrowse(P, list, { editScreen: ed, detailScreen: det, back: P.screens.some((s) => /^scrHome/.test(s.name)) });
      added = 3;
      if (starter) { // replace the empty starter screen and start on the new list
        P.screens.splice(P.screens.indexOf(starter), 1);
        const br = P.screens.find((s) => s.name === 'scrBrowse' + slug(list.name));
        P.screens.splice(P.screens.indexOf(br), 1); P.screens.unshift(br);
        P.app.startScreen = br.id;
      }
    }
    H.commit(); H.renderAll();
    m.close();
    H.toast(`Created ${list.name} with ${list.columns.length} columns and ${list.rows.length} sample rows${added ? ' plus 3 screens' : ''}`);
    if (notes.length) setTimeout(() => H.toast(notes[0]), 2700);
  }
  drawPick();
}
