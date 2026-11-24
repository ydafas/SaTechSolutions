// driftui.js — Project → "Check against a Studio export": list controls added in Studio and bring them into the project.
import { clone } from './core.js';
import { readStudioExport, driftReport, importDrift } from './drift.js';

let H = null; // { P, modal, toast, ui, esc, replaceProject }
export function initDriftUI(helpers) { H = helpers; }

export function driftDialog() {
  const esc = H.esc, ui = H.ui;
  const st = { studio: null, rep: null, pick: new Set(), pickScr: new Set(), files: '' };
  const m = H.modal(`<div class="modal-h"><h2>Check against a Studio export</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b" id="xBody"></div><div class="modal-f" id="xFoot"></div>`, 'md');
  const body = m.el.querySelector('#xBody'), foot = m.el.querySelector('#xFoot');

  function drawPick() {
    body.innerHTML = `<p class="note" style="margin-top:0">Changed something in Power Apps Studio after pasting? Choose the app's export and this lists every control that exists only in Studio, so the next paste does not lose it.</p>
      <label class="drop"><input type="file" id="xFile" accept=".msapp,.zip,.yaml,.yml" multiple hidden><b>Choose an .msapp, a solution .zip, or Src/*.pa.yaml files</b><small>Read in this browser; nothing is uploaded.</small></label>
      <div class="note-box"><b>Where to get it</b><ul><li>Studio → File → Save as → This computer (.msapp), or export the solution that holds the app (.zip).</li><li>Or open the app's <code>Src</code> folder and choose the screen <code>.pa.yaml</code> files.</li></ul></div>`;
    foot.innerHTML = `<span class="grow"></span><button class="btn" data-x>Close</button>`;
    foot.querySelector('[data-x]').onclick = m.close;
    body.querySelector('#xFile').onchange = async (e) => {
      const fs = [...e.target.files];
      try {
        if (fs.some((f) => !/\.ya?ml$/i.test(f.name)) && !window.JSZip) { H.toast('Zip library did not load', true); return; }
        const files = await Promise.all(fs.map(async (f) => ({ name: f.name, data: /\.ya?ml$/i.test(f.name) ? await f.text() : new Uint8Array(await f.arrayBuffer()) })));
        st.studio = await readStudioExport(files, window.JSZip);
        st.files = fs.map((f) => f.name).join(', ');
        st.rep = driftReport(H.P(), st.studio);
        st.pick = new Set(st.rep.screens.flatMap((s) => s.studioOnly.map((c) => `${s.name}|${c.name}`)));
        st.pickScr = new Set();
        drawReport();
      } catch (err) { H.toast('Could not read the export: ' + (err?.message || err), true); }
    };
  }
  function drawReport() {
    const r = st.rep, n = Object.keys(st.studio.screens).length;
    const only = r.screens.filter((s) => s.studioOnly.length), miss = r.screens.filter((s) => s.missing.length);
    body.innerHTML = `<p class="note" style="margin-top:0">${n} screen${n === 1 ? '' : 's'} read from ${esc(st.files)}.</p>
      ${st.studio.notes.map((x) => `<div class="note-box warn">${esc(x)}</div>`).join('')}
      ${!only.length && !r.newScreens.length ? '<div class="note-box"><b>Nothing was added in Studio.</b>Every control in the export is in your project.</div>' : ''}
      ${only.map((s) => `<div class="dr-scr"><b>${esc(s.name)}</b><small>only in Studio</small>${s.studioOnly.map((c) => `<label class="chk"><input type="checkbox" data-dk="${esc(s.name + '|' + c.name)}" ${st.pick.has(s.name + '|' + c.name) ? 'checked' : ''}><code>${esc(c.name)}</code> <span class="muted">${esc(c.control)}${c.nested ? ` · ${c.nested} inside` : ''}</span></label>`).join('')}</div>`).join('')}
      ${r.newScreens.length ? `<div class="dr-scr"><b>Screens only in Studio</b>${r.newScreens.map((s) => `<label class="chk"><input type="checkbox" data-ds="${esc(s.name)}" ${st.pickScr.has(s.name) ? 'checked' : ''}><code>${esc(s.name)}</code> <span class="muted">${s.controls} controls</span></label>`).join('')}</div>` : ''}
      ${miss.length ? `<div class="note-box"><b>In your project, not in the export</b><ul>${miss.map((s) => `<li>${esc(s.name)}: ${s.missing.length} control${s.missing.length === 1 ? '' : 's'} (${esc(s.missing.slice(0, 4).join(', '))}${s.missing.length > 4 ? '…' : ''}). Paste this screen again, or they were deleted in Studio.</li>`).join('')}</ul></div>` : ''}
      <p class="note">Imported controls keep every property exactly as Studio wrote it. They are placed in front of the screen's other controls; move them in Layers if needed.</p>`;
    const k = st.pick.size + st.pickScr.size;
    foot.innerHTML = `<span class="grow"></span><button class="btn" id="xBack">Choose other files</button><button class="btn pri" id="xGo" ${k ? '' : 'disabled'}>Import ${k} item${k === 1 ? '' : 's'}</button>`;
    body.onchange = (e) => {
      const t = e.target;
      if (t.dataset.dk) t.checked ? st.pick.add(t.dataset.dk) : st.pick.delete(t.dataset.dk);
      if (t.dataset.ds) t.checked ? st.pickScr.add(t.dataset.ds) : st.pickScr.delete(t.dataset.ds);
      drawReport();
    };
    foot.querySelector('#xBack').onclick = drawPick;
    foot.querySelector('#xGo').onclick = () => {
      const q = clone(H.P()), by = {};
      for (const key of st.pick) { const [s, n] = key.split('|'); (by[s] = by[s] || []).push(n); }
      const res = importDrift(q, st.studio, Object.entries(by).map(([screen, names]) => ({ screen, names })), [...st.pickScr]);
      H.replaceProject(q);
      m.close();
      const ren = Object.entries(res.renamed);
      H.toast(`Imported ${res.added} control${res.added === 1 ? '' : 's'} from Studio${ren.length ? `; renamed ${ren.map(([a, b]) => `${a} → ${b}`).join(', ')} (names must be unique)` : ''}`);
    };
  }
  drawPick();
}
