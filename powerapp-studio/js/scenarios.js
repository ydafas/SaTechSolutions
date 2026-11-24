// scenarios.js — build a complete app (lists, screens, options, flows) from a scenario in SCENARIOS,
// and describe what a scenario includes for the wizard's gallery.
import { newProject, clone, THEME_PRESETS, slug } from './core.js';
import { SCENARIOS, generateApp } from './templates.js';
import { flowFromSpec, addChainFlows } from './ai.js';
import { flowFxName } from './flows.js';

/** Chain config for a scenario (its stages plus the approvals list/field), or null. */
export function scenarioChain(sc) {
  if (!sc?.chain || !sc.approvals) return null;
  return { list: sc.approvals.list, field: sc.approvals.field || 'Status', stages: clone(sc.chain.stages), ...(sc.chain.reassign ? { reassign: true } : {}), ...(sc.chain.override ? { override: true } : {}) };
}

/** Add a scenario's flows to a project, skipping any whose list or columns no longer exist. Returns the flows added. */
export function addScenarioFlows(project, sc) {
  const out = [];
  const usedNames = new Set((project.flows || []).map((f) => flowFxName(f.name)));
  for (const spec of sc?.flows || []) {
    const l = project.lists.find((x) => x.name === spec.list);
    if (!l) continue;
    const f = flowFromSpec(project, spec, { findList: (n) => project.lists.find((x) => x.name === n), findCol: (list, n) => list.columns.find((c) => c.name === n), usedNames });
    if (!f || !f.recipients.length) continue;
    (project.flows = project.flows || []).push(f); out.push(f);
  }
  if (out.some((f) => /\{Link\}/.test(f.subject + f.body)) && !project.app.deepLink?.screen) {
    const f = out.find((x) => /\{Link\}/.test(x.subject + x.body));
    const det = project.screens.find((s) => s.name === 'scrDetail' + slug(f.list));
    if (det) project.app.deepLink = { screen: det.id, list: f.list };
  }
  return out;
}

/** A finished project for a scenario. opts: formFactor, theme (preset name), features (override), flows (default true). */
export function buildScenario(key, opts = {}) {
  const sc = SCENARIOS[key];
  if (!sc) throw new Error('Unknown scenario ' + key);
  const p = newProject();
  p.app.name = sc.app;
  p.app.formFactor = opts.formFactor === 'phone' ? 'phone' : 'tablet';
  if (opts.theme) p.theme = clone(THEME_PRESETS.find((t) => t.name === opts.theme) || THEME_PRESETS[0]);
  p.lists = sc.lists();
  p.app.features = { loading: true, confirm: true, stepper: true, scope: 'all', usage: false, adminGroup: '', ...(sc.features || {}), ...(opts.features || {}) };
  generateApp(p, { screens: opts.screens || ['home', 'browse', 'detail', 'edit'], approvals: sc.approvals || null, chain: scenarioChain(sc) });
  if (p.app.chain) addChainFlows(p);
  if (opts.flows !== false) addScenarioFlows(p, sc);
  return p;
}

/** One-line summary of what a scenario builds, for gallery cards. */
export function scenarioIncludes(key) {
  const sc = SCENARIOS[key];
  if (!sc || key === 'blank') return '';
  const p = buildScenario(key, { flows: true });
  const screens = p.screens.length, lists = p.lists.filter((l) => !l.system).length, flows = (p.flows || []).length;
  const bits = [`${lists} list${lists === 1 ? '' : 's'}`, `${screens} screens`];
  if (flows) bits.push(`${flows} flow${flows === 1 ? '' : 's'}`);
  if (p.app.chain) bits.push(`${p.app.chain.stages.length}-step approvals`);
  else if (p.app.approvalsCfg) bits.push('approvals');
  if (p.lists.some((l) => l.attachments)) bits.push('attachments');
  return bits.join(' · ');
}
