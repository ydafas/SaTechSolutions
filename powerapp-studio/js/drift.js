// drift.js — compare the project with a Power Apps Studio export (.msapp, solution .zip or Src/*.pa.yaml files)
// and list what was added in Studio, so it can be brought back into the project instead of being lost on the next paste.
import { emittedNames } from './emit.js';
import { rawControl } from './import.js';
import { allNames, newScreen } from './core.js';

const SKIP = /(^|\/)(App|_EditorState|DataSources|Themes|Connections|ComponentReferences)\b|\.pa\.yaml\.json$/i;
/**
 * Collect screen YAML from a Studio export. files: [{name, data}] where data is a string (a .pa.yaml) or an ArrayBuffer/Uint8Array
 * (.msapp or solution .zip; nested .msapp files inside a solution are opened too). Returns {screens: {name: def}, notes}.
 */
export async function readStudioExport(files, JSZip) {
  const yaml = globalThis.jsyaml;
  const screens = {}, notes = [];
  const addYaml = (path, text) => {
    if (SKIP.test(path)) return;
    let doc;
    try { doc = yaml.load(String(text).replace(/\r\n/g, '\n'), { schema: yaml.FAILSAFE_SCHEMA }); } catch (e) { notes.push(`${path}: not valid YAML (${e.message.split('\n')[0]})`); return; }
    if (!doc?.Screens) return;
    for (const [n, d] of Object.entries(doc.Screens)) screens[n] = d || {};
  };
  const openZip = async (name, data, depth = 0) => {
    const z = await JSZip.loadAsync(data);
    let found = 0;
    for (const [p, f] of Object.entries(z.files)) {
      if (f.dir) continue;
      if (/\.pa\.yaml$/i.test(p)) { found++; addYaml(p, await f.async('string')); }
      else if (/\.msapp$/i.test(p) && depth < 2) found += await openZip(p, await f.async('uint8array'), depth + 1);
    }
    if (!found && depth === 0) notes.push(`${name} has no YAML source (no Src/*.pa.yaml). In Studio, save with the YAML source files, or download each screen's .pa.yaml and choose those instead.`);
    return found;
  };
  for (const f of files) {
    if (typeof f.data === 'string') addYaml(f.name, f.data);
    else await openZip(f.name, f.data);
  }
  return { screens, notes };
}

const kids = (def) => (def?.Children || []).map((ch) => Object.entries(ch)[0]);
const allIn = (def) => kids(def).flatMap(([n, d]) => [n, ...allIn(d)]);

/**
 * What differs between the project and the export, per screen:
 * studioOnly: top-level controls Studio has and the project does not ({name, control, nested}),
 * missing: controls the project emits that the export lacks (not pasted yet, or deleted in Studio),
 * newScreens: screens that exist only in Studio.
 */
export function driftReport(project, studio) {
  const out = { screens: [], newScreens: [] };
  for (const [sn, sd] of Object.entries(studio.screens)) {
    const scr = project.screens.find((s) => s.name === sn);
    if (!scr) { out.newScreens.push({ name: sn, controls: allIn(sd).length }); continue; }
    const ours = new Set(emittedNames(project, scr));
    const theirs = new Set(allIn(sd));
    const studioOnly = [];
    // walk top-down: a Studio-only parent carries its children with it
    const walk = (def) => { for (const [n, d] of kids(def)) { if (!ours.has(n)) studioOnly.push({ name: n, control: String(d?.Control || ''), nested: allIn(d).length }); else walk(d); } };
    walk(sd);
    const missing = [...ours].filter((n) => !theirs.has(n));
    if (studioOnly.length || missing.length) out.screens.push({ name: sn, id: scr.id, studioOnly, missing });
  }
  return out;
}

/**
 * Bring Studio-only controls into the project exactly as Studio wrote them (raw controls), at the front of their screen.
 * picks: [{screen, names:[...]}] for existing screens; newScreens: [names] to import whole. Mutates the project.
 * Returns {added, renamed} — renamed maps names that clashed with other screens to their new names.
 */
export function importDrift(project, studio, picks, newScreens = []) {
  let added = 0;
  const renamed = {};
  const taken = allNames(project);
  const unique = (n) => { if (!taken.has(n)) { taken.add(n); return n; } let i = 2; while (taken.has(`${n}_${i}`)) i++; renamed[n] = `${n}_${i}`; taken.add(renamed[n]); return renamed[n]; };
  const take = (n, d) => rawControl(project, unique(n), d);
  for (const pk of picks) {
    const scr = project.screens.find((s) => s.name === pk.screen), sd = studio.screens[pk.screen];
    if (!scr || !sd) continue;
    const want = new Set(pk.names);
    const find = (def) => kids(def).flatMap(([n, d]) => (want.has(n) ? [[n, d]] : find(d)));
    for (const [n, d] of find(sd)) { scr.controls.push(take(n, d)); added++; }
  }
  for (const sn of newScreens) {
    const sd = studio.screens[sn];
    if (!sd) continue;
    const scr = newScreen(project, unique(sn));
    if (sd.Properties?.OnVisible) scr.onVisible = String(sd.Properties.OnVisible).replace(/^=/, '');
    if (sd.Properties?.Fill) scr.fillFx = String(sd.Properties.Fill).replace(/^=/, '');
    scr.controls = kids(sd).map(([n, d]) => take(n, d));
    project.screens.push(scr);
    added += scr.controls.length;
  }
  return { added, renamed };
}
