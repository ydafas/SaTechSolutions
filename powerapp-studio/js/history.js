// history.js — named snapshots of the whole project, a readable diff between any two states,
// and restore. Snapshots live in the project (so they travel with the project file and bundle);
// the oldest automatic ones are dropped first when the cap is reached.
import { uid, clone, walk, MAX_SNAPSHOTS } from './core.js';

export const WARN_BYTES = 1500000; // localStorage holds about 5 MB per site

/** Parts of a project that describe the tenant or the history, not the app: never snapshotted, kept across restores. */
export const META_KEYS = ['snapshots', 'findings', 'deployed'];
/** The project as snapshot text: the app itself, without history, findings or deploy records. */
export function snapshotData(project) {
  const p = clone(project);
  for (const k of META_KEYS) delete p[k];
  return JSON.stringify(p);
}
/** Add a snapshot of `source` (default: the project itself) to `target`. Returns the snapshot. */
export function takeSnapshot(target, name, { auto = false, source = target, at = new Date().toISOString() } = {}) {
  const snap = { id: uid(), name: String(name || 'Snapshot').slice(0, 80), at, auto, data: snapshotData(source) };
  const list = (target.snapshots = [...(target.snapshots || []), snap]);
  while (list.length > MAX_SNAPSHOTS) {
    const i = list.findIndex((x) => x.auto && x !== snap);
    list.splice(i >= 0 ? i : 0, 1);
  }
  return snap;
}
/** Snapshots of both lists, oldest first, without duplicates, capped like takeSnapshot (automatic ones go first). */
export function mergeSnapshots(a = [], b = []) {
  const seen = new Set(), list = [...a, ...b].filter((x) => x && !seen.has(x.id) && seen.add(x.id)).sort((x, y) => String(x.at).localeCompare(String(y.at)));
  while (list.length > MAX_SNAPSHOTS) { const i = list.findIndex((x) => x.auto); list.splice(i >= 0 ? i : 0, 1); }
  return list;
}
export const snapshotsBytes = (project) => (project.snapshots || []).reduce((n, x) => n + x.data.length, 0);
export function deleteSnapshot(project, id) {
  project.snapshots = (project.snapshots || []).filter((x) => x.id !== id);
  if (!project.snapshots.length) delete project.snapshots;
}
export const snapshotProject = (snap) => JSON.parse(snap.data);

/** A new project restored from a snapshot; the snapshot list carries over, plus a snapshot of `current`. */
export function restoreSnapshot(current, id) {
  const snap = (current.snapshots || []).find((x) => x.id === id);
  if (!snap) throw new Error('Snapshot not found.');
  const p = snapshotProject(snap);
  for (const k of META_KEYS) if (current[k]) p[k] = clone(current[k]);
  p.snapshots = [...(current.snapshots || [])];
  takeSnapshot(p, `Before restoring "${snap.name}"`, { auto: true, source: current });
  return p;
}

// ------------------------------------------------------------------ diff
const byId = (arr) => new Map((arr || []).map((x) => [x.id, x]));
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
function diffSet(a, b, name, compare) {
  const A = byId(a), B = byId(b), out = { added: [], removed: [], changed: [] };
  for (const [id, x] of B) if (!A.has(id)) out.added.push(name(x));
  for (const [id, x] of A) if (!B.has(id)) out.removed.push(name(x));
  for (const [id, x] of B) if (A.has(id)) { const d = compare(A.get(id), x); if (d) out.changed.push(d); }
  return out;
}
const flat = (scr) => [...walk(scr.controls || [])].map(([c]) => c);
const ctlCore = (c) => { const x = { ...c }; delete x.children; return x; };

/** What changed from project a (older) to b (newer). */
export function diffProjects(a, b) {
  const lists = diffSet(a.lists, b.lists, (l) => l.name, (x, y) => {
    const cols = diffSet(x.columns, y.columns, (c) => c.name, (c1, c2) => {
      const what = [];
      if (c1.name !== c2.name) what.push(`renamed from ${c1.name}`);
      if (c1.type !== c2.type) what.push(`${c1.type} → ${c2.type}`);
      if (!!c1.required !== !!c2.required) what.push(c2.required ? 'now required' : 'no longer required');
      if (!same(c1.choices, c2.choices)) what.push('choices changed');
      if (c1.lookupList !== c2.lookupList) what.push('lookup target changed');
      return what.length ? `${c2.name} (${what.join(', ')})` : null;
    });
    const rows = !same(x.rows || [], y.rows || []);
    if (x.name === y.name && !cols.added.length && !cols.removed.length && !cols.changed.length && !rows && x.size === y.size) return null;
    return { name: y.name, renamedFrom: x.name !== y.name ? x.name : '', columns: cols, rows, size: x.size !== y.size ? (y.size || 'not set') : '' };
  });
  const screens = diffSet(a.screens, b.screens, (s) => s.name, (x, y) => {
    const ctls = diffSet(flat(x), flat(y), (c) => c.name, (c1, c2) => (same(ctlCore(c1), ctlCore(c2)) ? null : c2.name));
    const props = ['fill', 'onVisible', 'role'].filter((k) => !same(x[k], y[k]));
    if (x.name === y.name && !ctls.added.length && !ctls.removed.length && !ctls.changed.length && !props.length) return null;
    return { name: y.name, renamedFrom: x.name !== y.name ? x.name : '', controls: ctls, props };
  });
  const flows = diffSet(a.flows, b.flows, (f) => f.name, (x, y) => (same(x, y) ? null : y.name));
  const appKeys = [...new Set([...Object.keys(a.app || {}), ...Object.keys(b.app || {})])].filter((k) => !same(a.app?.[k], b.app?.[k]));
  return { lists, screens, flows, theme: !same(a.theme, b.theme), app: appKeys };
}
export function diffIsEmpty(d) {
  const e = (s) => !s.added.length && !s.removed.length && !s.changed.length;
  return e(d.lists) && e(d.screens) && e(d.flows) && !d.theme && !d.app.length;
}
/** Plain-language lines for a diff. */
export function describeDiff(d) {
  const out = [];
  const set = (label, s) => {
    if (s.added.length) out.push(`${label} added: ${s.added.join(', ')}`);
    if (s.removed.length) out.push(`${label} removed: ${s.removed.join(', ')}`);
  };
  set('Lists', d.lists);
  for (const l of d.lists.changed) {
    if (l.renamedFrom) out.push(`List ${l.renamedFrom} renamed to ${l.name}`);
    set(`${l.name} columns`, l.columns);
    for (const c of l.columns.changed) out.push(`${l.name}.${c}`);
    if (l.rows) out.push(`${l.name} sample rows changed`);
    if (l.size) out.push(`${l.name} expected size: ${l.size}`);
  }
  set('Screens', d.screens);
  for (const s of d.screens.changed) {
    if (s.renamedFrom) out.push(`Screen ${s.renamedFrom} renamed to ${s.name}`);
    set(`${s.name} controls`, s.controls);
    if (s.controls.changed.length) out.push(`${s.name} controls changed: ${s.controls.changed.slice(0, 12).join(', ')}${s.controls.changed.length > 12 ? ` and ${s.controls.changed.length - 12} more` : ''}`);
    if (s.props.length) out.push(`${s.name} screen settings changed (${s.props.join(', ')})`);
  }
  set('Flows', d.flows);
  if (d.flows.changed.length) out.push(`Flows changed: ${d.flows.changed.join(', ')}`);
  if (d.theme) out.push('Theme changed');
  const appNames = { name: 'app name', formFactor: 'format', startScreen: 'start screen', features: 'app options', chain: 'approval chain', siteUrl: 'site URL', appLink: 'app link', deepLink: 'email-link screen', suppressProps: 'switched-off properties', onStart: 'extra OnStart', formulas: 'extra formulas' };
  const keys = d.app.filter((k) => k !== 'setup' && k !== 'flowSample');
  if (keys.length) out.push(`App settings changed: ${keys.map((k) => appNames[k] || k).join(', ')}`);
  return out;
}
