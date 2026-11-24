// historyui.js — Project → "Version history": save named versions, compare any of them with now, restore.
import { takeSnapshot, deleteSnapshot, restoreSnapshot, snapshotProject, diffProjects, describeDiff, diffIsEmpty, snapshotsBytes, WARN_BYTES } from './history.js';
import { MAX_SNAPSHOTS } from './core.js';

let H = null; // { P, modal, toast, ui, esc, ask, persist, replaceProject }
export function initHistoryUI(helpers) { H = helpers; }

const kb = (n) => (n < 1024 * 1024 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1024 / 1024).toFixed(1)} MB`);
const when = (at) => { const d = new Date(at); return isNaN(d) ? '' : d.toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' }); };

export function historyDialog() {
  const esc = H.esc, ui = H.ui;
  const st = { view: 'list', id: null };
  const m = H.modal(`<div class="modal-h"><h2>Version history</h2><button class="mini" data-close>${ui('close')}</button></div>
    <div class="modal-b" id="hBody"></div><div class="modal-f" id="hFoot"></div>`, 'md');
  const body = m.el.querySelector('#hBody'), foot = m.el.querySelector('#hFoot');

  function drawList() {
    const P = H.P(), snaps = [...(P.snapshots || [])].reverse(), bytes = snapshotsBytes(P);
    body.innerHTML = `<p class="note" style="margin-top:0">Versions are kept inside the project, so they travel with the project file and the app bundle. A version is also saved automatically before Describe → Build, Ask for a change, a Studio import and a bundle import. Up to ${MAX_SNAPSHOTS} are kept; the oldest automatic ones go first.</p>
      <div class="h-save"><input class="inp" id="hName" maxlength="80" placeholder="Name this version, e.g. Before the approvals rework" aria-label="Version name"><button class="btn pri" id="hSave">${ui('plus')}Save version</button></div>
      ${bytes > WARN_BYTES ? `<div class="note-box warn"><b>Versions take ${kb(bytes)}.</b>This browser keeps about 5 MB per site. Delete old versions, or save a project file, so nothing is lost.</div>` : ''}
      ${snaps.length ? `<div class="h-list">${snaps.map((s) => `<div class="h-row" data-id="${esc(s.id)}"><div class="h-main"><b>${esc(s.name)}</b><small>${esc(when(s.at))}${s.auto ? ' · <span class="h-auto">automatic</span>' : ''}</small></div>
          <button class="btn sm" data-h="diff">Compare with now</button><button class="btn sm" data-h="restore">Restore</button><button class="mini danger" data-h="del" title="Delete version" aria-label="Delete ${esc(s.name)}">${ui('trash')}</button></div>`).join('')}</div>`
        : '<div class="note-box"><b>No versions yet.</b>Save one before a big change; you can compare and go back to it any time.</div>'}`;
    foot.innerHTML = `<span class="grow">${snaps.length ? `${snaps.length} version${snaps.length === 1 ? '' : 's'} · ${kb(bytes)}` : ''}</span><button class="btn" data-x>Close</button>`;
    foot.querySelector('[data-x]').onclick = m.close;
    const name = body.querySelector('#hName');
    const save = () => {
      const snap = takeSnapshot(H.P(), name.value.trim() || `Version ${(H.P().snapshots || []).length + 1}`);
      if (H.persist()) H.toast(`Saved "${snap.name}"`);
      drawList();
    };
    body.querySelector('#hSave').onclick = save;
    name.onkeydown = (e) => { if (e.key === 'Enter') save(); };
    body.querySelector('.h-list')?.addEventListener('click', (e) => {
      const b = e.target.closest('[data-h]'); if (!b) return;
      const id = b.closest('[data-id]').dataset.id, s = (H.P().snapshots || []).find((x) => x.id === id);
      if (!s) return;
      if (b.dataset.h === 'diff') { st.view = 'diff'; st.id = id; draw(); }
      if (b.dataset.h === 'restore') restore(s);
      if (b.dataset.h === 'del') H.ask('Delete version?', `Delete "${s.name}" from ${when(s.at)}? This cannot be undone.`, 'Delete').then((ok) => { if (ok) { deleteSnapshot(H.P(), id); H.persist(); drawList(); } });
    });
  }

  function drawDiff() {
    const P = H.P(), s = (P.snapshots || []).find((x) => x.id === st.id);
    if (!s) { st.view = 'list'; return draw(); }
    let lines = [], err = '';
    try { const d = diffProjects(snapshotProject(s), P); lines = diffIsEmpty(d) ? [] : describeDiff(d); } catch (e) { err = e.message; }
    body.innerHTML = `<p class="note" style="margin-top:0">What changed from <b>${esc(s.name)}</b> (${esc(when(s.at))}) to the app as it is now.</p>
      ${err ? `<div class="note-box err">${esc(err)}</div>` : lines.length ? `<ul class="h-diff">${lines.map((l) => `<li>${esc(l)}</li>`).join('')}</ul>` : '<div class="note-box"><b>No differences.</b>The app is the same as this version.</div>'}`;
    foot.innerHTML = `<button class="btn" id="hBack">Back to versions</button><span class="grow"></span><button class="btn pri" id="hRestore" ${lines.length ? '' : 'disabled'}>Restore this version</button>`;
    foot.querySelector('#hBack').onclick = () => { st.view = 'list'; draw(); };
    foot.querySelector('#hRestore').onclick = () => restore(s);
  }

  function restore(s) {
    H.ask('Restore this version?', `The app goes back to "${s.name}". The current app is saved as a version first, and Undo also brings it back.`, 'Restore').then((ok) => {
      if (!ok) return;
      try { H.replaceProject(restoreSnapshot(H.P(), s.id)); H.toast(`Restored "${s.name}"`); m.close(); }
      catch (e) { H.toast('Could not restore: ' + e.message, true); }
    });
  }
  const draw = () => (st.view === 'diff' ? drawDiff() : drawList());
  draw();
}
