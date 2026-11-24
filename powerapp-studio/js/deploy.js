// deploy.js — Paste Assistant model: the ordered deploy steps for a project, a stable hash of
// every artifact a step pastes or imports, and change tracking against what was last deployed.
// Pure functions (no DOM) so the tests can drive them. The UI lives in ui.js (deployView).
import { screenYaml, themeFormulas, onStartFx, startScreenFx, usedLists, usesPeople, provisioningScript, scriptKind } from './emit.js';
import { flowDefinition, flowFxName, flowFileName } from './flows.js';
import { FORM_FACTORS } from './core.js';
import { attachSlots, attachHow, attachCopyText } from './attach.js';

/** FNV-1a 32-bit, hex. Deterministic and fast; collisions only matter as "looks unchanged". */
export function hash(text) {
  let h = 0x811c9dc5;
  const s = String(text ?? '');
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
  return h.toString(16).padStart(8, '0');
}

/**
 * Ordered deploy steps. Each: {key, kind, title, how (instructions), text? (what Copy copies),
 * file? (flow package to download), flow?, hash ('' for steps with nothing to track)}.
 */
export function deploySteps(project) {
  const steps = [];
  const ff = FORM_FACTORS[project.app.formFactor || 'tablet'];
  const lists = usedLists(project);
  const groups = !!project.app.features?.adminGroup;
  if (project.lists.length) {
    const t = provisioningScript(project);
    steps.push({ key: 'lists', kind: 'lists', title: 'SharePoint lists', text: t, copyLabel: 'Copy provisioning script',
      how: `Create ${project.lists.map((l) => `"${l.name}"`).join(', ')} on your site, by hand or with this ${scriptKind(project) === 'gcc' ? 'GCC High script (MSAL.PS sign-in + SharePoint REST; paste into PowerShell)' : 'PnP PowerShell script'} (fill in its EDIT ME block; switch flavour in Get YAML → Provisioning script). Skip this step if the lists already exist with these columns.`,
      hash: hash(JSON.stringify(project.lists.map((l) => [l.name, l.columns.map((c) => [c.name, c.type, c.required, c.choices, c.lookupList])])) + (project.app.features?.usage ? '+usage' : '') + project.lists.filter((l) => l.attachments).map((l) => '+att:' + l.name).join('') + scriptKind(project)) });
  }
  steps.push({ key: 'app', kind: 'static', title: 'Blank canvas app', hash: hash(project.app.formFactor),
    how: `In Power Apps, create a blank canvas app in ${project.app.formFactor === 'phone' ? 'Phone' : 'Tablet'} format (${ff.w} × ${ff.h}). Skip this when you are updating an app you already deployed.` });
  const sources = [...lists.map((l) => `SharePoint list "${l}"`), ...(usesPeople(project) ? ['Office 365 Users'] : []), ...(groups ? ['Office 365 Groups'] : []), ...(project.app.features?.usage ? ['SharePoint list "AppUsage"'] : [])];
  if (sources.length) steps.push({ key: 'datasources', kind: 'static', title: 'Data sources', text: sources.join('\n'), copyLabel: 'Copy the list',
    how: `Data → Add data, and add exactly these (names must match):\n${sources.map((x) => '• ' + x).join('\n')}`, hash: hash(sources.join('|')) });
  const fm = themeFormulas(project);
  steps.push({ key: 'formulas', kind: 'formulas', title: 'App.Formulas (theme)', text: fm, copyLabel: 'Copy App.Formulas', hash: hash(fm),
    how: 'Tree view → App → pick Formulas in the property list → select all → paste.' });
  const os = onStartFx(project);
  if (os) steps.push({ key: 'onstart', kind: 'onstart', title: 'App.OnStart', text: os, copyLabel: 'Copy App.OnStart', hash: hash(os),
    how: 'App → OnStart. Expand the formula bar first (multi-line text pasted into a collapsed bar keeps only one line), replace all, then App → Run OnStart.' });
  for (const f of (project.flows || []).filter((x) => x.enabled !== false)) {
    const def = JSON.stringify(flowDefinition(project, f));
    steps.push({ key: 'flow:' + f.id, kind: 'flow', flow: f, title: 'Flow: ' + f.name, file: flowFileName(f), hash: hash(def),
      how: `Power Automate → My flows → Import → Import Package (Legacy) → upload ${flowFileName(f)} → "Create as new" (or "Update" if you imported it before) → pick your SharePoint and Office 365 Outlook connections → Import → open it and turn it on.` +
        (f.kind === 'button' ? `\nThen in Power Apps Studio: Power Automate pane → Add flow → "${f.name}" (it appears as ${flowFxName(f.name)}). Do this before pasting the screens that run it.` : '') });
  }
  for (const scr of project.screens) {
    const y = screenYaml(project, scr);
    const att = attachSlots(project, scr);
    steps.push({ key: 'scr:' + scr.id, kind: 'screen', screen: scr, title: 'Screen ' + scr.name, text: y, copyLabel: 'Copy screen YAML', hash: hash(y),
      how: `Delete the old ${scr.name} (and any copy using the same control names) if it exists. Click an empty spot on the canvas, then press Ctrl+V.`
        + (att.length ? `\nIf you added ${att.map((x) => x.form).join(' / ')} by hand before: select it, Ctrl+C and keep it in Notepad before you delete the screen, then paste it back.` : '') });
    if (att.length) {
      const how = attachHow(scr.name, att);
      steps.push({ key: 'attach:' + scr.id, kind: 'attach', screen: scr, title: 'Attachments on ' + scr.name, text: attachCopyText(att), copyLabel: 'Copy the property values', hash: hash(how + y), how });
    }
  }
  const ss = '=' + startScreenFx(project);
  steps.push({ key: 'startscreen', kind: 'startscreen', title: 'App.StartScreen', text: ss, copyLabel: 'Copy StartScreen', hash: hash(ss),
    how: 'App → StartScreen → paste. Do it after the screens exist, or the names show red.' });
  steps.push({ key: 'test', kind: 'static', title: 'Test with Play (F5)', hash: '',
    how: 'Run OnStart once more, then press Play (F5) from the start screen. Timers and OnVisible do not run on the editing canvas. Try one item end to end' + ((project.flows || []).length ? ' and wait a minute for flows to run.' : '.') });
  return steps;
}

/** 'new' (never deployed), 'changed', 'unchanged'. Steps without a hash are 'new' until marked done once. */
export function stepStatus(project, step) {
  const d = project.deployed?.[step.key];
  if (!d) return 'new';
  if (!step.hash) return 'unchanged';
  return d.hash === step.hash ? 'unchanged' : 'changed';
}
export function markDeployed(project, step, at = new Date().toISOString()) {
  project.deployed = project.deployed && typeof project.deployed === 'object' ? project.deployed : {};
  project.deployed[step.key] = { hash: step.hash, at, name: step.screen?.name || step.flow?.name || step.title };
}
/** Steps to do this time: everything on a first deploy, otherwise only new or changed (plus the final test). */
export function pendingSteps(project, steps = deploySteps(project)) {
  const first = !project.deployed || !Object.keys(project.deployed).length;
  if (first) return steps;
  return steps.filter((s) => s.key === 'test' || stepStatus(project, s) !== 'unchanged');
}
/** Screens and flows deployed before that no longer exist in the project (delete them in Studio). */
export function removedSinceDeploy(project) {
  const keys = Object.keys(project.deployed || {});
  const scr = new Set(project.screens.map((s) => 'scr:' + s.id)), fl = new Set((project.flows || []).map((f) => 'flow:' + f.id));
  return keys.filter((k) => (k.startsWith('scr:') && !scr.has(k)) || (k.startsWith('flow:') && !fl.has(k))).map((k) => ({ key: k, name: project.deployed[k].name || k }));
}
export function deploySummary(project) {
  const steps = deploySteps(project);
  const st = steps.map((s) => stepStatus(project, s));
  return { total: steps.length, changed: st.filter((x) => x === 'changed').length, fresh: st.filter((x) => x === 'new').length, deployedBefore: !!Object.keys(project.deployed || {}).length };
}
