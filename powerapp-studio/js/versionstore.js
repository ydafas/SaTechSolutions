// versionstore.js — where Version history lives in this browser. Snapshots are whole-project copies, so keeping
// them inside the saved project filled localStorage (about 5 MB per site). They now go to IndexedDB (hundreds of
// MB), or to their own localStorage key when IndexedDB is unavailable (some private windows and sandboxes).
// The project itself stays in localStorage without them. In memory they are still project.snapshots, so
// history.js, project files and app bundles work as before.
const DB = 'pas-versions', STORE = 'versions', REC = 'current', LS_KEY = 'pas.versions.v1';

function openDb(idb) {
  return new Promise((resolve, reject) => {
    if (!idb) { reject(new Error('IndexedDB is not available')); return; }
    let req;
    try { req = idb.open(DB, 1); } catch (e) { reject(e); return; }
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('IndexedDB open failed'));
    req.onblocked = () => reject(new Error('IndexedDB blocked'));
  });
}
function tx(db, mode, fn) {
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode), st = t.objectStore(STORE);
    const r = fn(st);
    t.oncomplete = () => resolve(r?.result);
    t.onerror = () => reject(t.error || new Error('IndexedDB write failed'));
    t.onabort = () => reject(t.error || new Error('IndexedDB write aborted'));
  });
}
const env0 = () => ({ idb: globalThis.indexedDB, ls: globalThis.localStorage });
const clean = (list) => (Array.isArray(list) ? list.filter((x) => x && typeof x.id === 'string' && typeof x.data === 'string') : []);

/** Stored versions ([] when none). Reads IndexedDB first, then the localStorage fallback. */
export async function readVersions(env = env0()) {
  let out = [];
  try { const db = await openDb(env.idb); out = clean(await tx(db, 'readonly', (st) => st.get(REC))); db.close(); }
  catch { /* fall back below */ }
  try { const t = env.ls?.getItem(LS_KEY); if (t) out = [...out, ...clean(JSON.parse(t)).filter((x) => !out.some((y) => y.id === x.id))]; } catch { /* unreadable */ }
  return out;
}
/** Store the versions. Returns 'idb', 'ls' or '' (nothing could be stored). Empty list → removes them. */
export async function writeVersions(list, env = env0()) {
  const data = clean(list);
  try {
    const db = await openDb(env.idb);
    await tx(db, 'readwrite', (st) => (data.length ? st.put(data, REC) : st.delete(REC)));
    db.close();
    try { env.ls?.removeItem(LS_KEY); } catch { /* ignore */ }
    return 'idb';
  } catch { /* fall back */ }
  try {
    if (data.length) env.ls.setItem(LS_KEY, JSON.stringify(data)); else env.ls.removeItem(LS_KEY);
    return 'ls';
  } catch { return ''; }
}
/** A cheap signature: snapshots never change once taken, so ids say whether the list changed. */
export const versionsSig = (list) => (list || []).map((x) => x.id).join(',');
