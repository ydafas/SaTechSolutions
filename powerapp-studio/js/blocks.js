// blocks.js — "My blocks": save a selection of controls as a reusable block, keep blocks in this browser
// (localStorage, best effort), export/import them as .json, and insert a block into any screen with fresh
// names. Formulas inside the block that point at its own controls (txtSearch.Text, galItems.Selected …)
// follow the new names; everything else (lists, variables, theme) is left as written.
import { clone, uid, walk, allNames, renameCopy, normalizeProject, companions, CATALOG } from './core.js';
import { renameFx } from './import.js';

export const STORE_KEY = 'pas.blocks.v1';
export const FORMAT = 'power-app-studio-blocks';
export const MAX_BLOCKS = 60;
export const MAX_BYTES = 1500000; // one import file
const SAFE_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,63}$/;

/** Names a set of controls (and their gallery children) defines, including generated companions. */
function definedNames(ctls) {
  const out = [];
  for (const [c] of walk(ctls)) out.push(c.name, ...companions(c));
  return out;
}

/** Lists the block reads or writes, so inserting it can warn when the app does not have them. */
export function blockLists(controls) {
  const s = new Set();
  for (const [c] of walk(controls)) {
    for (const v of [c.field?.list, c.bind?.list, c.data?.list, c.data?.previewList]) if (v) s.add(v);
    for (const acts of Object.values(c.actions || {})) for (const a of acts || []) if (a?.list) s.add(a.list);
  }
  return [...s];
}

/** A block from top-level controls of one screen. Positions become relative to the block's top-left corner. */
export function makeBlock(ctls, name) {
  const controls = clone(ctls.filter(Boolean));
  if (!controls.length) return null;
  const x0 = Math.min(...controls.map((c) => c.x)), y0 = Math.min(...controls.map((c) => c.y));
  for (const c of controls) { c.x -= x0; c.y -= y0; }
  for (const [c] of walk(controls)) delete c.__parent;
  const w = Math.max(...controls.map((c) => c.x + c.w)), h = Math.max(...controls.map((c) => c.y + c.h));
  const nm = String(name || '').trim().slice(0, 60) || controls[0].name;
  return { id: uid(), name: nm, created: new Date().toISOString().slice(0, 10), w, h, lists: blockLists(controls), controls };
}

/** Clean one block from storage or an imported file. Returns null when it is not usable. */
export function sanitizeBlock(b) {
  if (!b || typeof b !== 'object' || !Array.isArray(b.controls) || !b.controls.length) return null;
  // reuse the project loader's control checks (ids, types, numbers, colors, actions)
  const tmp = normalizeProject({ app: {}, lists: [], screens: [{ id: 's', name: 'scrBlock', controls: clone(b.controls) }] });
  const controls = tmp.screens[0].controls;
  for (const [c] of walk(controls)) {
    if (!SAFE_NAME.test(c.name)) c.name = (String(c.name).replace(/[^A-Za-z0-9_]/g, '').replace(/^[^A-Za-z_]+/, '') || CATALOG[c.type]?.prefix || 'ctl').slice(0, 60);
    if (c.type === 'raw') return null; // unknown control types cannot be emitted safely
  }
  const name = String(b.name || '').replace(/[<>]/g, '').trim().slice(0, 60) || controls[0].name;
  const w = Math.max(...controls.map((c) => c.x + c.w)), h = Math.max(...controls.map((c) => c.y + c.h));
  return { id: /^[A-Za-z0-9_-]{1,40}$/.test(b.id || '') ? b.id : uid(), name, created: String(b.created || '').slice(0, 10), w, h, lists: blockLists(controls), controls };
}

/** Insert a copy of a block at (x, y). Returns {added, map, warnings}; map is oldName → newName. */
export function insertSavedBlock(project, screen, block, at = { x: 40, y: 96 }) {
  const src = clone(block.controls);
  const used = allNames(project);
  const before = definedNames(src);
  const copies = src.map((c) => renameCopy(project, c, used));
  const after = definedNames(copies);
  const map = {};
  before.forEach((n, i) => { if (after[i] && after[i] !== n) map[n] = after[i]; });
  const fix = (v) => (typeof v === 'string' ? renameFx(v, map) : v);
  for (const [c] of walk(copies)) {
    for (const k of Object.keys(c.fx || {})) c.fx[k] = fix(c.fx[k]);
    if (c.data) {
      for (const k of ['searchCtl', 'statusCtl']) if (map[c.data[k]]) c.data[k] = map[c.data[k]];
      if (c.data.extraFilter) c.data.extraFilter = fix(c.data.extraFilter);
    }
    for (const acts of Object.values(c.actions || {})) for (const a of acts || []) {
      if (a.type === 'custom') a.fx = fix(a.fx);
      if (a.type === 'setVar') a.value = fix(a.value);
      if (map[a.commentCtl]) a.commentCtl = map[a.commentCtl];
    }
  }
  for (const c of copies) { c.x += Math.round(at.x); c.y += Math.round(at.y); screen.controls.push(c); }
  const warnings = [];
  const missing = blockLists(copies).filter((l) => !project.lists.some((x) => x.name === l));
  if (missing.length) warnings.push(`This block uses ${missing.join(', ')}, which ${missing.length > 1 ? 'are' : 'is'} not in this app's data model. Add ${missing.length > 1 ? 'them' : 'it'} on the Data tab or rebind the controls.`);
  const scrIds = new Set(project.screens.map((s) => s.id));
  for (const [c] of walk(copies)) for (const acts of Object.values(c.actions || {})) for (const a of acts || []) if (a?.screen && !scrIds.has(a.screen)) { a.screen = ''; if (!warnings.some((w) => /screen/.test(w))) warnings.push('Screen links inside the block pointed at screens of another app; pick their targets again.'); }
  return { added: copies, map, warnings };
}

// ------------------------------------------------------------------ storage (best effort: private windows and full quotas just do not keep blocks)
export function loadBlocks(store = globalThis.localStorage) {
  try {
    const raw = JSON.parse(store?.getItem(STORE_KEY) || '[]');
    return (Array.isArray(raw) ? raw : []).map(sanitizeBlock).filter(Boolean).slice(0, MAX_BLOCKS);
  } catch { return []; }
}
export function saveBlocks(blocks, store = globalThis.localStorage) {
  try { store?.setItem(STORE_KEY, JSON.stringify(blocks.slice(0, MAX_BLOCKS))); return true; } catch { return false; }
}

// ------------------------------------------------------------------ files
export function exportBlocksJson(blocks) {
  return JSON.stringify({ format: FORMAT, version: 1, blocks }, null, 2);
}
/** Parse an exported file. Returns {blocks, skipped} or throws with a plain message. */
export function importBlocksJson(text) {
  if (String(text).length > MAX_BYTES) throw new Error('That file is too large for a blocks file.');
  let d; try { d = JSON.parse(text); } catch { throw new Error('That file is not JSON.'); }
  const arr = Array.isArray(d) ? d : d?.format === FORMAT ? d.blocks : d?.controls ? [d] : null;
  if (!Array.isArray(arr)) throw new Error('That file has no blocks in it.');
  const blocks = arr.slice(0, MAX_BLOCKS).map(sanitizeBlock);
  return { blocks: blocks.filter(Boolean).map((b) => ({ ...b, id: uid() })), skipped: blocks.filter((b) => !b).length };
}
