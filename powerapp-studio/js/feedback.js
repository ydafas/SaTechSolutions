// feedback.js — "Report a paste error": read error text copied from Power Apps Studio or Power Automate,
// work out which control/property/list it is about, explain the likely cause with the skill's lessons,
// and offer a fix (select the control, stop emitting a rejected property, add a data source, mark a list
// large). Every report goes into a findings log Tom can export and send back. Pure functions; no DOM.
import { walk, uid, MAX_FINDINGS } from './core.js';
import { flatNodes } from './emit.js';
import { flowFxName } from './flows.js';
import { FORM_RE, allAttachSlots } from './attach.js';

const CONNECTORS = { Office365Users: 'Office 365 Users', Office365Groups: 'Office 365 Groups', Office365Outlook: 'Office 365 Outlook' };

/** Everything the parser can recognise by name: emitted controls, screens, lists, columns, flows, variables. */
export function nameIndex(project) {
  const controls = new Map(), screens = new Map(), lists = new Map(), columns = new Map(), flows = new Map();
  for (const scr of project.screens) {
    screens.set(scr.name, scr);
    const designer = [...walk(scr.controls)].map(([c]) => c);
    for (const n of flatNodes(project, scr)) {
      const owner = designer.filter((c) => n.name === c.name || n.name.startsWith(c.name)).sort((a, b) => b.name.length - a.name.length)[0] || null;
      controls.set(n.name, { name: n.name, control: n.control, base: String(n.control || '').split('@')[0], props: n.props.map(([k]) => k), screen: scr, owner });
    }
  }
  for (const l of project.lists) { lists.set(l.name, l); for (const c of l.columns) columns.set(c.name, { list: l, col: c }); }
  for (const f of project.flows || []) { flows.set(f.name, f); flows.set(flowFxName(f.name), f); }
  return { controls, screens, lists, columns, flows };
}
const tokens = (t) => [...new Set(String(t).match(/[A-Za-z_][A-Za-z0-9_]*/g) || [])];
const quoted = (t) => [...String(t).replace(/\b(\w+)['’](t|s|re|ve|ll|d)\b/gi, '$1$2').matchAll(/['"‘’“”]([^'"‘’“”\n]{1,80})['"‘’“”]/g)].map((m) => m[1].trim());

/** Split pasted text into separate error messages (Studio's error pane copies one per line or per block). */
export function splitReports(text) {
  const blocks = String(text || '').replace(/\r\n/g, '\n').split(/\n\s*\n|\n(?=\s*(?:PA\d{4}|Error|Warning|Delegation warning|\d+\.\s))/i).map((x) => x.trim()).filter(Boolean);
  return blocks.length ? blocks : [];
}

/** Recognise one message. Returns a finding: {family, code, title, cause, fix, where, actions[], excerpt}. */
export function analyzeOne(project, text, idx = nameIndex(project)) {
  const t = String(text), low = t.toLowerCase();
  const code = (t.match(/\bPA\d{4}\b/) || [''])[0];
  const toks = tokens(t), qs = quoted(t);
  const ctl = toks.map((x) => idx.controls.get(x)).find(Boolean) || null;
  const scr = ctl?.screen || toks.map((x) => idx.screens.get(x)).find(Boolean) || null;
  const where = ctl ? `${ctl.screen.name} › ${ctl.name}` : scr ? scr.name : '';
  const select = ctl?.owner ? [{ kind: 'select', ctlId: ctl.owner.id, screenId: ctl.screen.id, label: `Select ${ctl.owner.name}` }] : scr ? [{ kind: 'screen', screenId: scr.id, label: `Open ${scr.name}` }] : [];
  const out = (family, title, cause, fix, actions = []) => ({ family, code, title, cause, fix, where, actions: [...actions, ...select], excerpt: t.slice(0, 2000) });

  // the hand-added attachments form (attach.js) is missing or named differently
  const fm = (t.match(FORM_RE) || [])[0];
  if (fm && /recogni[sz]ed|isn'?t valid|is not valid|invalid name|unknown|not found/i.test(t)) {
    const slot = allAttachSlots(project).find((x) => x.form === fm);
    const acts = slot ? [{ kind: 'select', ctlId: slot.ctl.id, screenId: slot.screen.id, label: `Show where ${fm} goes` }] : [];
    const r = out('attachForm', `${fm} is not on the screen yet`,
      `Attachments cannot be pasted. ${slot ? `${slot.screen.name} expects an Edit form named ${fm} that you insert by hand` : `The app expects an Edit form named ${fm} that you insert by hand`}; the loader and Save refer to it by that name. This error shows right after a paste (or a re-paste, which deletes hand-added controls).`,
      `Follow the Paste assistant's "Attachments on ${slot ? slot.screen.name : 'the screen'}" step: Insert → Edit form, rename it to ${fm}, data source ${slot ? slot.list : 'the list'}, only the Attachments field, Item and DefaultMode as listed. If you do not want attachments, turn Attachments off for the list on the Data tab and copy the screen again.`, acts);
    if (slot) r.where = `${slot.screen.name} › ${slot.ctl.name}`;
    return r;
  }
  // PA2108 — a property this control type doesn't have in this tenant
  if (code === 'PA2108' || /unknown property/i.test(t)) {
    const prop = (t.match(/unknown property[:\s]+['"‘’]?([A-Za-z][A-Za-z0-9]*)/i) || [])[1]
      || qs.find((x) => /^[A-Z][A-Za-z0-9]+$/.test(x) && (!ctl || ctl.props.includes(x)))
      || (ctl && toks.find((x) => ctl.props.includes(x) && !['X', 'Y', 'Width', 'Height'].includes(x)));
    const base = ctl?.base || '';
    const acts = prop && base ? [{ kind: 'suppress', control: base, prop, label: `Stop emitting ${prop} on ${base}` }] : [];
    return out('unknownProperty', prop ? `${base || 'A control'} does not accept ${prop}` : 'Unknown property',
      `Your tenant's version of ${base || 'this control'} has no ${prop || 'such'} property, so the whole paste fails (PA2108). The studio only emits properties field-tested elsewhere; tenants differ.`,
      prop && base ? `Switch ${prop} off for every ${base} in this app, then copy the screen again. The control keeps its default look for that property.` : 'Paste the full error text, including the control and property name, so the studio can pinpoint it.', acts);
  }
  // PA2110 — duplicate name
  if (code === 'PA2110' || /already exists/i.test(t)) {
    const name = qs.find((x) => /^[A-Za-z_]\w*$/.test(x)) || ctl?.name || '';
    return out('duplicateName', `"${name || 'A name'}" already exists in the app`,
      'Control names must be unique across the whole app, not just the screen. This usually means the old copy of the screen (or a renamed copy) is still in Studio.',
      'In Studio, delete the old screen first (Tree view → screen → Delete), then paste again. If both copies must stay, rename the control here and copy the screen again.');
  }
  // PA2109 — control type/version
  if (code === 'PA2109' || /unknown control|control type|version .*(?:not|isn't) (?:supported|available)/i.test(t)) {
    return out('controlVersion', 'Control type or version not available',
      'Studio did not recognise the control type or its version pin. The studio emits classic controls without pins (except Gallery@2.15.0); imported controls keep the pins from their export.',
      ctl?.owner?.type === 'raw' ? `Open Formulas on ${ctl.owner.name} and remove the version (the part after @) from its control type, or delete it and add the control in Studio.` : 'Send this to us (Export log): it means a control type the studio emits is different in your tenant.');
  }
  // PA1001 / YAML parse
  if (code === 'PA1001' || /could not be parsed|failed to parse|mapping values|did not find expected|while parsing|PaModule|yaml/i.test(t)) {
    return out('parse', 'The YAML did not parse',
      'Part of the paste was cut off or changed on the way (a formula with ": " or " #" in a plain line, a lost indent, or a partial copy).',
      'Use the Copy button again (not a manual selection), click an empty spot of the canvas, and press Ctrl+V. If it repeats, export the log: it is a generator bug we fix with a new lint rule.');
  }
  // "X isn't recognized" / "Name isn't valid"
  if (/isn'?t recognized|is not recognized|name isn'?t valid|name is not valid|unknown or unsupported function|invalid name/i.test(t)) {
    const names = [...qs.filter((x) => /^[A-Za-z_]\w*$/.test(x)), ...toks.filter((x) => idx.lists.has(x) || idx.flows.has(x) || CONNECTORS[x] || /^(var|loc|col)/.test(x) || x === 'ThisItem' || x === 'appTheme' || idx.screens.has(x) || idx.controls.has(x))];
    const n = names.find((x) => x !== ctl?.name) || names[0] || '';
    if (n === 'ThisItem') return out('thisItem', 'ThisItem used outside a gallery', 'A control that reads ThisItem landed outside its gallery, which happens when gallery rows are pasted at the wrong indent.', 'Copy the whole screen again with the Copy button and paste it; do not paste single controls out of a gallery. If it repeats, export the log for us.');
    if (n === 'appTheme') return out('missingFormulas', 'appTheme is not defined yet', 'Every control reads its colors from the appTheme named formula.', 'Paste App.Formulas first (App → Formulas), then the screens. Red names clear by themselves once it exists.');
    if (idx.lists.has(n)) return out('missingDataSource', `Data source "${n}" is not added`, 'Formulas use the list by name, and Studio only knows lists added as data sources.', `Data → Add data → SharePoint → your site → tick "${n}" (the name must match exactly).`, [{ kind: 'dataTab', list: n, label: 'Open the Data tab' }]);
    if (CONNECTORS[n]) return out('missingConnector', `${CONNECTORS[n]} is not added`, `${CONNECTORS[n]} is a connector the app calls (people pickers, group checks or emails).`, `Data → Add data → search "${CONNECTORS[n]}" → add it, then touch the red formula to recheck.`);
    if (idx.flows.has(n)) return out('missingFlow', `Flow ${n} is not added to the app`, 'Run-a-flow buttons call the flow by its Power Apps name.', `Import the flow package first, then in Studio: Power Automate (left bar) → Add flow → pick "${idx.flows.get(n).name}".`);
    if (idx.screens.has(n)) return out('missingScreen', `Screen ${n} is not pasted yet`, 'A Navigate() or StartScreen points at a screen that does not exist in Studio yet.', `Paste ${n} too; red references clear once it exists (retype them if they stay red).`);
    if (/^(var|loc|col)/.test(n)) return out('missingVariable', `${n} has no value yet`, 'Variables are set in App.OnStart, and Studio does not rerun OnStart after a paste.', 'Paste App.OnStart if you have not, then App → … → Run OnStart.');
    if (idx.controls.has(n)) return out('missingControl', `Control ${n} is not on this screen in Studio`, 'A formula refers to a control on another screen or one that was not pasted.', `Paste ${idx.controls.get(n).screen.name} as a whole screen (it holds ${n}).`);
    return out('unknownName', n ? `"${n}" is not recognised` : 'A name is not recognised', 'Usually a data source, connector or screen that is not added yet, or a column whose SharePoint name differs.', 'Check Data → the list exists with that exact name and its columns match the Data tab here. Export the log if it stays.');
  }
  // Patch / argument type
  if (/expected record, found text|found text|expecting a record|expected record/i.test(t)) {
    const c = toks.map((x) => idx.columns.get(x)).find(Boolean);
    return out('choiceAsText', c ? `${c.list.name}.${c.col.name} is a different type in SharePoint` : 'A column is a Choice (record) in SharePoint',
      'SharePoint expects a record ({Value: …}) here: the column is Choice, Person or Lookup in SharePoint, but plain text in this project (or the other way round).',
      c ? `Set ${c.list.name}.${c.col.name} to the same type as in SharePoint on the Data tab, then copy the screens that save it again.` : 'Compare the column types on the Data tab with the SharePoint list settings.', c ? [{ kind: 'dataTab', list: c.list.name, label: 'Open the Data tab' }] : []);
  }
  if (/invalid argument type|wrong type|incompatible type|expects? (?:a )?(?:text|number|boolean|record|table)/i.test(t)) {
    return out('argType', 'Invalid argument type', 'A formula passes the wrong kind of value: often IfError with a fallback of a different type, or a column type mismatch with SharePoint.',
      'If it is your own formula override, make both branches the same type (use Blank() as the fallback). Otherwise compare column types on the Data tab with SharePoint.');
  }
  // column name mismatch
  if (/column .* (?:does not exist|doesn't exist)|most similar/i.test(t)) {
    const c = qs.map((x) => idx.columns.get(x.replace(/\s+/g, ''))).find(Boolean);
    return out('columnName', 'Column name does not match SharePoint', 'SharePoint knows the column by a different (internal) name, typically because it was created by hand with spaces (Needed By → Needed_x0020_By).',
      'Create the list with the provisioning script (names without spaces), or rename the column here to the exact SharePoint name.', c ? [{ kind: 'dataTab', list: c.list.name, label: 'Open the Data tab' }] : []);
  }
  // delegation
  if (/delegat|large data sets|may not work correctly/i.test(t)) {
    const l = toks.map((x) => idx.lists.get(x)).find(Boolean);
    return out('delegation', l ? `Delegation warning on ${l.name}` : 'Delegation warning',
      'Studio warns when part of a query cannot run on SharePoint; it then only looks at the first 500 (or 2000) items. Harmless for small lists.',
      l ? `If ${l.name} can grow past 2,000 items, mark it Large: counts show "2000+" and lint flags the formulas that need reshaping.` : 'Mark the list Large on the Data tab if it can grow past 2,000 items.', l ? [{ kind: 'largeList', list: l.name, label: `Mark ${l.name} as large` }] : []);
  }
  // Power Automate import / run
  if (/connection/i.test(t) && /(missing|not found|invalid|select|unauthori[sz]ed|authorization|failed)/i.test(t)) {
    return out('flowConnection', 'Flow connection problem', 'Flow packages carry no connections; Power Automate needs one picked for SharePoint and for Office 365 Outlook during import.',
      'Import again: in Related resources, click "Select during import" for each connection and pick (or create) yours, then Import and turn the flow on.');
  }
  if (/invalidtemplate|template validation|template language|unable to process|invalid expression|InvalidOpenApiFlow/i.test(t)) {
    const action = (t.match(/action ['"]([^'"]+)['"]/i) || [])[1];
    return out('flowTemplate', action ? `Flow action "${action}" failed validation` : 'Flow definition failed validation',
      'An expression or parameter in the flow did not validate in your tenant (a column internal name, a list name, or an action parameter that differs).',
      'Open the flow in Power Automate, fix the red action (usually re-pick the list or column), and export the log so the generator gets fixed.');
  }
  if (/import package|package import|import failed|couldn.t import|could not import/i.test(t)) {
    return out('flowImport', 'Flow package import failed', 'Legacy import rejects packages when Dataverse-solution flows are enforced, or when a resource is missing.',
      'Use My flows → Import → Import Package (Legacy) → Create as new. If your admin enforces solutions, export the log and ask for a solution package.');
  }
  return out('unknown', code ? `${code}: not recognised yet` : 'Error not recognised yet', 'The studio has no rule for this message yet.', 'Keep it in the log and export it: each new message becomes a fix or a lint rule.');
}

/** Analyze pasted text (one or several messages). */
export function analyze(project, text) {
  const idx = nameIndex(project);
  return splitReports(text).map((t) => analyzeOne(project, t, idx));
}

/** Apply a fix action that changes the project. Returns a short description, or '' when nothing changed. */
export function applyAction(project, a) {
  if (a.kind === 'suppress' && a.control && /^[A-Za-z]\w*$/.test(a.prop || '')) {
    const sp = (project.app.suppressProps = project.app.suppressProps || {});
    const arr = (sp[a.control] = sp[a.control] || []);
    if (!arr.includes(a.prop)) arr.push(a.prop);
    return `${a.prop} is no longer emitted on ${a.control}.`;
  }
  if (a.kind === 'unsuppress') {
    const sp = project.app.suppressProps || {};
    sp[a.control] = (sp[a.control] || []).filter((x) => x !== a.prop);
    if (!sp[a.control].length) delete sp[a.control];
    if (!Object.keys(sp).length) delete project.app.suppressProps;
    return `${a.prop} is emitted on ${a.control} again.`;
  }
  if (a.kind === 'largeList') { const l = project.lists.find((x) => x.name === a.list); if (l) { l.size = 'large'; return `${l.name} is marked large.`; } }
  return '';
}

/** Add findings to the project log (newest last, capped). */
export function logFindings(project, findings, at = new Date().toISOString()) {
  project.findings = project.findings || [];
  for (const f of findings) project.findings.push({ id: uid(), at, code: f.code, family: f.family, title: f.title, where: f.where, excerpt: f.excerpt, status: 'open' });
  project.findings = project.findings.slice(-MAX_FINDINGS);
}
/** Plain-text log to send back. */
export function exportFindings(project) {
  const sp = project.app.suppressProps || {};
  const lines = [`Power App Studio – tenant findings for "${project.app.name}"`, `Exported ${new Date().toISOString()}`, ''];
  if (Object.keys(sp).length) { lines.push('Properties switched off:'); for (const [k, v] of Object.entries(sp)) lines.push(`  ${k}: ${v.join(', ')}`); lines.push(''); }
  const fs = project.findings || [];
  if (!fs.length) lines.push('No findings logged.');
  fs.forEach((f, i) => lines.push(`#${i + 1} [${f.status}] ${f.code || f.family} – ${f.title}`, `   when: ${f.at}${f.where ? `   where: ${f.where}` : ''}`, ...f.excerpt.split('\n').map((x) => '   > ' + x), ''));
  return lines.join('\n');
}
/** Control types the app emits, with their emitted properties (for the manual "switch off" picker). */
export function emittedTypes(project) {
  const m = new Map();
  for (const scr of project.screens) for (const n of flatNodes(project, scr)) {
    const b = String(n.control || '').split('@')[0];
    if (!m.has(b)) m.set(b, new Set());
    n.props.forEach(([k]) => { if (!['X', 'Y', 'Width', 'Height'].includes(k)) m.get(b).add(k); });
  }
  return [...m].map(([k, v]) => [k, [...v].sort()]).sort((a, b) => a[0].localeCompare(b[0]));
}
