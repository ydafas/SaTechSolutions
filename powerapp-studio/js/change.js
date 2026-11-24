// change.js — "Ask for a change": edit an existing project with Claude instead of rebuilding it.
// Claude sees a compact summary of the project and returns a list of small, typed edit operations
// (structured outputs). Each operation is validated, shown to the user as a plain sentence, and applied
// through the same helpers the designer uses. Screens an operation does not touch are left alone.
import { COLTYPES, THEME_PRESETS, newList, newColumn, clone, walk } from './core.js';
import { FLOW_KINDS, flowFxName } from './flows.js';
import { MODELS, flowFromSpec, applyFlowFields, addRunFlowButton, STAGE_SPEC, stagesFromSpec, addChainFlows } from './ai.js';
import { ensureChainColumns, normalizeChain, guessChainValues, MAX_STAGES } from './chain.js';
import { genBlank, genHome, genBrowse, genDetail, genEdit, genApprovals, genChainApprovals, genLoading, statusCol } from './templates.js';
import { cleanName, makeColumn, renameListIn, renameColumnIn, removeColumnIn, addFieldToForm, addFieldToDetail, addFieldToGallery,
  refreshStatusVisuals, setAdminGating, screensFor, screenKind } from './edit.js';

export const OPS = {
  renameApp: 'Rename the app', setTheme: 'Change the theme', setOption: 'Change an app option',
  addList: 'Add a list', renameList: 'Rename a list',
  addColumn: 'Add a column', renameColumn: 'Rename a column', removeColumn: 'Remove a column', setChoices: 'Change the choices of a column',
  addFieldToForm: 'Show a column on the form', addFieldToDetail: 'Show a column on the details screen', addFieldToGallery: 'Show a column in the list rows',
  addScreenFromTemplate: 'Add a screen', addFlow: 'Add a flow', editFlow: 'Change a flow', removeFlow: 'Remove a flow',
  setApprovalStages: 'Set up a multi-stage approval chain',
};
const DESTRUCTIVE = new Set(['removeColumn', 'removeFlow', 'renameColumn', 'renameList', 'setChoices', 'setApprovalStages']);
const OPTIONS = ['loadingScreen', 'confirmationPage', 'statusStepper', 'showMineFirst', 'trackUsage', 'adminGroupId'];
const TEMPLATES = ['browse', 'detail', 'edit', 'approvals', 'home', 'blank', 'loading'];

// ------------------------------------------------------------------ schema (structured outputs: all objects closed, all fields required)
const obj = (properties) => ({ type: 'object', additionalProperties: false, required: Object.keys(properties), properties });
const str = (description) => ({ type: 'string', ...(description ? { description } : {}) });
const arr = (items, description) => ({ type: 'array', items, ...(description ? { description } : {}) });
const en = (values, description) => ({ type: 'string', enum: values, ...(description ? { description } : {}) });
export const CHANGE_SCHEMA = obj({
  summary: str('One or two sentences: what the change does'),
  ops: arr(obj({
    op: en(Object.keys(OPS)),
    reason: str('Why this operation is needed, in plain words'),
    list: str('List the operation works on; empty if none'),
    column: str('Existing column name (rename/remove/setChoices/addField*), or the new column name for addColumn'),
    newName: str('New name for renameApp, renameList, renameColumn and editFlow (rename); else empty'),
    type: en(['', ...Object.keys(COLTYPES)], 'addColumn: the column type'),
    required: { type: 'boolean' },
    choices: arr(str(), 'addColumn (Choice) and setChoices: the full list of choices'),
    lookupList: str('addColumn (Lookup): target list'),
    columns: arr(obj({ name: str(), type: en(Object.keys(COLTYPES)), required: { type: 'boolean' }, choices: arr(str()), lookupList: str() }), 'addList: its columns (not Title)'),
    screen: str('Target screen name when there is more than one candidate; else empty'),
    template: en(['', ...TEMPLATES], 'addScreenFromTemplate: which screen'),
    theme: en(['', ...THEME_PRESETS.map((t) => t.name)], 'setTheme: preset name'),
    option: en(['', ...OPTIONS], 'setOption: which option'),
    value: str('setOption: "true"/"false", or the group ID for adminGroupId; renameApp: the new name goes in newName'),
    flowName: str('addFlow: new flow name; editFlow/removeFlow: the existing flow name'),
    flowKind: en(['', ...Object.keys(FLOW_KINDS)], 'addFlow: the kind'),
    recipients: arr(obj({ kind: en(['column', 'author', 'editor', 'me', 'address']), value: str() }), 'addFlow/editFlow: replaces recipients when not empty'),
    subject: str('addFlow/editFlow: email subject (empty keeps it)'),
    body: str('addFlow/editFlow: email body (empty keeps it)'),
    statusField: str('addFlow/editFlow statusChanged or reminder: column'),
    statusValues: arr(str()),
    reminderValue: str(),
    days: { type: 'integer' },
    hour: { type: 'integer' },
    stages: arr(STAGE_SPEC, 'setApprovalStages: the 1-4 stages in order; else empty'),
  })),
  notes: arr(str(), 'Anything you could not do with these operations, or questions for the person'),
});

export const CHANGE_PROMPT = `You edit an existing Microsoft Power Apps canvas app (SharePoint lists + Power Automate email flows) by returning a short list of edit operations. A generator applies them to the person's project; everything you do not mention stays exactly as it is, including screens they designed by hand. You are given a JSON summary of the current project and the person's change request.

Operations (use the fewest that do the job, in the order they must run):
- renameApp (newName). setTheme (theme). setOption (option, value): loadingScreen, confirmationPage, statusStepper, showMineFirst, trackUsage take "true"/"false"; adminGroupId takes a Microsoft 365 group ID (GUID) or "" to remove admin-only buttons.
- addList (list, columns): a new SharePoint list; every list already has Title. Add its screens with addScreenFromTemplate (edit first, then detail, then browse).
- renameList (list, newName). addColumn (list, column, type, required, choices, lookupList). renameColumn (list, column, newName). removeColumn (list, column). setChoices (list, column, choices) for Choice columns: send the complete new list.
- addFieldToForm / addFieldToDetail / addFieldToGallery (list, column, screen optional): show an existing or just-added column on the list's form, details screen or gallery rows. When the person adds a column they almost always want it on the form and details screen too; add it to the gallery only if they ask or it is a key value.
- addScreenFromTemplate (template, list): browse, detail, edit, approvals (list needs a Choice status column), home, blank, loading.
- setApprovalStages (list, column = the Choice status column, stages): turn the list's approvals into a chain of 1 to 4 sequential stages, or change its stages. Each stage: name, approverKind personColumn (approverColumn = an existing Person column, or empty to add one) or group (groupId, notifyEmail if given). It adds the per-stage columns, rebuilds that list's details and approvals screens and adds the hand-off and outcome emails itself: do not add those columns or flows separately. The requester can never approve their own item; the creator's manager cannot be looked up, so use a Person column.
- addFlow (flowKind, flowName, list, recipients, subject, body, statusField, statusValues, reminderValue, days, hour). editFlow (flowName, then only the fields to change; newName to rename). removeFlow (flowName).

Rules:
- Use exact list, column, screen and flow names from the summary. New names: PascalCase, letters and digits, no spaces; never Title, ID, Created, Modified, Author, Editor or names SharePoint already uses (StartDate, EndDate, Location, Order, Type, Name, Path, ContentType, Attachments).
- Column types: Text, Note, Number, Currency, DateTime, Choice (with choices), YesNo, Person, Lookup (to another list).
- Flow recipients: column (a Person column), author, editor, me (button flows only), address (only addresses the person gives). Email tokens: {ColumnName}, {Title}, {ID}, {Created}, {Author}, {PersonColumn.Email}, {Link}.
- Fill every field of every operation; use "", [], 0 and false for fields that do not apply. Put a short plain-words reason on each operation.
- If part of the request cannot be done with these operations (for example moving controls or changing formulas), say so in notes instead of guessing.`;

// ------------------------------------------------------------------ project summary for Claude
export function projectSummary(project) {
  const ft = project.app.features || {};
  return {
    app: { name: project.app.name, formFactor: project.app.formFactor, theme: project.theme.name,
      options: { loadingScreen: !!ft.loading, confirmationPage: ft.confirm !== false, statusStepper: ft.stepper !== false, showMineFirst: ft.scope === 'mine', trackUsage: !!ft.usage, adminGroupId: ft.adminGroup || '' },
      approvals: project.app.approvalsCfg || null, ...(project.app.chain ? { approvalChain: { list: project.app.chain.list, column: project.app.chain.field, stages: project.app.chain.stages.map((s) => ({ name: s.name, approver: s.approver === 'group' ? 'group' : s.column })) } } : {}) },
    lists: project.lists.filter((l) => !l.system).map((l) => ({ name: l.name, columns: l.columns.map((c) => ({ name: c.name, type: c.type, required: !!c.required, ...(c.type === 'Choice' ? { choices: c.choices } : {}), ...(c.type === 'Lookup' ? { lookupList: c.lookupList } : {}) })), ...(l.rows?.length ? { sampleRows: l.rows.length } : {}) })),
    screens: project.screens.map((s) => {
      const cs = [...walk(s.controls)].map(([c]) => c);
      const lists = [...new Set(cs.flatMap((c) => [c.field?.list, c.bind?.list, c.data?.list]).filter(Boolean))];
      const out = { name: s.name, kind: screenKind(s), lists };
      const form = cs.filter((c) => c.field?.column).map((c) => c.field.column);
      const shown = cs.filter((c) => c.bind?.src === 'sel' && c.bind.field).map((c) => c.bind.field);
      const gals = cs.filter((c) => c.type === 'gallery' && c.data?.list).map((c) => ({ name: c.name, list: c.data.list, extraColumns: c.data.metaCols || [] }));
      if (form.length) out.formColumns = [...new Set(form)];
      if (shown.length) out.shownColumns = [...new Set(shown)];
      if (gals.length) out.galleries = gals;
      out.controls = cs.length;
      return out;
    }),
    flows: (project.flows || []).map((f) => ({ name: f.name, kind: f.kind, list: f.list, recipients: f.recipients, subject: f.subject, statusField: f.statusField, statusValues: f.statusValues })),
  };
}
export function buildChangeRequest(project, text, model = MODELS[0][0]) {
  return {
    model, max_tokens: 12000, system: CHANGE_PROMPT,
    messages: [{ role: 'user', content: `Current project:\n${JSON.stringify(projectSummary(project))}\n\nChange request:\n${String(text).trim()}` }],
    output_config: { format: { type: 'json_schema', schema: CHANGE_SCHEMA } },
  };
}

// ------------------------------------------------------------------ describe + validate
const listOf = (p, n) => p.lists.find((l) => l.name === n);
const colOf = (l, n) => l?.columns.find((c) => c.name === n);
/** Plain-language description of one operation. */
export function describeOp(o) {
  const L = o.list ? `${o.list}` : '';
  switch (o.op) {
    case 'renameApp': return `Rename the app to "${o.newName}"`;
    case 'setTheme': return `Switch the theme to ${o.theme}`;
    case 'setOption': return `Set ${o.option} to ${o.option === 'adminGroupId' ? (o.value || '(none)') : o.value}`;
    case 'addList': return `Add a list ${o.list}${o.columns?.length ? ` with ${o.columns.map((c) => `${c.name} (${c.type})`).join(', ')}` : ''}`;
    case 'renameList': return `Rename list ${L} to ${o.newName}`;
    case 'addColumn': return `Add ${o.type || 'Text'} column ${o.column} to ${L}${o.required ? ' (required)' : ''}${o.choices?.length ? `: ${o.choices.join(', ')}` : ''}`;
    case 'renameColumn': return `Rename ${L}.${o.column} to ${o.newName}`;
    case 'removeColumn': return `Remove ${L}.${o.column} and the controls that show it`;
    case 'setChoices': return `Set the choices of ${L}.${o.column} to ${(o.choices || []).join(', ')}`;
    case 'addFieldToForm': return `Show ${o.column} on the ${L} form${o.screen ? ` (${o.screen})` : ''}`;
    case 'addFieldToDetail': return `Show ${o.column} on the ${L} details screen${o.screen ? ` (${o.screen})` : ''}`;
    case 'addFieldToGallery': return `Show ${o.column} in the ${L} list rows`;
    case 'addScreenFromTemplate': return `Add a ${o.template} screen${L ? ` for ${L}` : ''}`;
    case 'addFlow': return `Add flow "${o.flowName}" (${FLOW_KINDS[o.flowKind]?.label || o.flowKind}) on ${L}`;
    case 'editFlow': return `Change flow "${o.flowName}"${o.newName ? ` (rename to ${o.newName})` : ''}`;
    case 'removeFlow': return `Remove flow "${o.flowName}"`;
    case 'setApprovalStages': return `Approve ${L} in ${(o.stages || []).length} stage${(o.stages || []).length === 1 ? '' : 's'}: ${(o.stages || []).map((s) => s.name).join(' → ')} (rebuilds its details and approvals screens)`;
  }
  return o.op;
}
/**
 * Check operations against the project as they would run in order (later ops may use lists/columns earlier ops add).
 * Returns [{op, desc, destructive, error}] — error is a reason when the op cannot run.
 */
export function reviewOps(project, ops) {
  const sim = { lists: new Map(project.lists.map((l) => [l.name, new Set(l.columns.map((c) => c.name))])), flows: new Set((project.flows || []).map((f) => f.name)) };
  const out = [];
  for (const o of ops || []) {
    let error = '';
    const has = (l) => sim.lists.has(l), hasCol = (l, c) => has(l) && (sim.lists.get(l).has(c) || sim.lists.get(l).has(cleanName(c, '')));
    switch (o?.op) {
      case 'renameApp': if (!String(o.newName || '').trim()) error = 'No new name.'; break;
      case 'setTheme': if (!THEME_PRESETS.some((t) => t.name === o.theme)) error = `Unknown theme "${o.theme}".`; break;
      case 'setOption':
        if (!OPTIONS.includes(o.option)) error = `Unknown option "${o.option}".`;
        else if (o.option === 'adminGroupId' && o.value && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(o.value).trim())) error = `"${o.value}" is not a group ID.`;
        break;
      case 'addList': { const n = cleanName(o.list, ''); if (!n) error = 'No list name.'; else if (has(n)) error = `${n} already exists.`; else sim.lists.set(n, new Set(['Title', ...(o.columns || []).map((c) => cleanName(c.name, 'Field'))])); break; }
      case 'renameList': if (!has(o.list)) error = `No list ${o.list}.`; else if (!cleanName(o.newName, '')) error = 'No new name.'; else { sim.lists.set(cleanName(o.newName, ''), sim.lists.get(o.list)); sim.lists.delete(o.list); } break;
      case 'addColumn': if (!has(o.list)) error = `No list ${o.list}.`; else if (!cleanName(o.column, '')) error = 'No column name.'; else if (o.type === 'Choice' && !(o.choices || []).length) error = 'A Choice column needs choices.'; else sim.lists.get(o.list).add(cleanName(o.column, '')).add(String(o.column)); break;
      case 'renameColumn': if (!hasCol(o.list, o.column)) error = `No column ${o.list}.${o.column}.`; else if (o.column === 'Title') error = 'Title cannot be renamed.'; else if (!cleanName(o.newName, '')) error = 'No new name.'; else { sim.lists.get(o.list).delete(o.column); sim.lists.get(o.list).add(cleanName(o.newName, '')); } break;
      case 'removeColumn': if (!hasCol(o.list, o.column)) error = `No column ${o.list}.${o.column}.`; else if (o.column === 'Title') error = 'Title cannot be removed.'; else { sim.lists.get(o.list).delete(o.column); sim.lists.get(o.list).delete(cleanName(o.column, '')); } break;
      case 'setChoices': { const c = colOf(listOf(project, o.list), o.column); if (!hasCol(o.list, o.column)) error = `No column ${o.list}.${o.column}.`; else if (c && c.type !== 'Choice') error = `${o.column} is not a Choice column.`; else if (!(o.choices || []).filter(Boolean).length) error = 'No choices given.'; break; }
      case 'addFieldToForm': case 'addFieldToDetail': case 'addFieldToGallery':
        if (!hasCol(o.list, o.column)) error = `No column ${o.list}.${o.column}.`;
        else if (o.screen && !project.screens.some((s) => s.name === o.screen)) error = `No screen ${o.screen}.`;
        break;
      case 'addScreenFromTemplate':
        if (!TEMPLATES.includes(o.template)) error = `Unknown template "${o.template}".`;
        else if (['browse', 'detail', 'edit', 'approvals'].includes(o.template) && !has(o.list)) error = `The ${o.template} screen needs a list.`;
        else if (o.template === 'loading' && project.screens.some((s) => s.role === 'loading')) error = 'The app already has a loading screen.';
        break;
      case 'addFlow': if (!FLOW_KINDS[o.flowKind]) error = `Unknown flow kind "${o.flowKind}".`; else if (!has(o.list)) error = `No list ${o.list}.`; else sim.flows.add(o.flowName); break;
      case 'editFlow': case 'removeFlow': if (!sim.flows.has(o.flowName)) error = `No flow "${o.flowName}".`; else if (o.op === 'removeFlow') sim.flows.delete(o.flowName); else if (o.newName) { sim.flows.delete(o.flowName); sim.flows.add(o.newName); } break;
      case 'setApprovalStages': if (!has(o.list)) error = `No list ${o.list}.`; else if (!(o.stages || []).length || o.stages.length > MAX_STAGES) error = `An approval chain needs 1 to ${MAX_STAGES} stages.`; else for (const st of o.stages) if (st.approverKind === 'personColumn') sim.lists.get(o.list).add(st.approverColumn || ''); break;
      default: error = `Unknown operation "${o?.op}".`;
    }
    out.push({ op: o, desc: o?.op ? describeOp(o) : 'Unknown operation', destructive: DESTRUCTIVE.has(o?.op), error });
  }
  return out;
}

// ------------------------------------------------------------------ apply
/**
 * Apply operations to a project (mutates it; pass a clone). Returns {applied, skipped, notes, touched}.
 * touched: ids of screens whose YAML may change; every other screen is left exactly as it was.
 */
export function applyOps(project, ops) {
  const applied = [], skipped = [], notes = [], touched = new Set();
  const touch = (s) => { if (s) (Array.isArray(s) ? s : [s]).forEach((x) => x && touched.add(x.id)); };
  const findList = (n) => listOf(project, n) || listOf(project, cleanName(n, ''));
  const findCol = (l, n) => l && (colOf(l, n) || colOf(l, cleanName(n, '')));
  const ft = () => (project.app.features = project.app.features || {});
  for (const o of ops) {
    const rv = reviewOps(project, [o])[0];
    if (rv.error) { skipped.push({ op: o, reason: rv.error }); continue; }
    let err = '';
    try {
      switch (o.op) {
        case 'renameApp': project.app.name = String(o.newName).trim().slice(0, 60); notes.push('Screens that show the app name (home title, loading screen) keep the old text until you edit those labels.'); break;
        case 'setTheme': project.theme = clone(THEME_PRESETS.find((t) => t.name === o.theme)); break;
        case 'setOption': {
          const on = /^(true|yes|1)$/i.test(String(o.value));
          if (o.option === 'loadingScreen') {
            const cur = project.screens.find((s) => s.role === 'loading');
            if (on && !cur) {
              const next = project.screens.find((s) => s.id === project.app.startScreen) || project.screens[0];
              ft().loading = true; const scr = genLoading(project, { next, lists: project.lists }); project.app.startScreen = scr.id; touch(scr);
            } else if (!on && cur) {
              ft().loading = false; project.screens.splice(project.screens.indexOf(cur), 1);
              if (project.app.startScreen === cur.id) project.app.startScreen = project.screens[0]?.id;
            } else ft().loading = on;
          } else if (o.option === 'trackUsage') {
            ft().usage = on;
            if (on && !project.lists.some((l) => l.name === 'AppUsage')) {
              const u = newList('AppUsage', [newColumn('DisplayName', 'Text'), newColumn('AppName', 'Text'), newColumn('FirstSeen', 'DateTime', { withTime: true }), newColumn('LastSeen', 'DateTime', { withTime: true }), newColumn('VisitCount', 'Number')]);
              u.system = true; u.columns[0].required = false; project.lists.push(u);
            }
            if (!on) project.lists = project.lists.filter((l) => !(l.system && l.name === 'AppUsage'));
            const ld = project.screens.find((s) => s.role === 'loading');
            if (ld) notes.push('The loading screen was built before this change; add it again (remove it, then setOption loadingScreen true) if usage tracking should run there instead of in OnStart.');
          } else if (o.option === 'adminGroupId') {
            const gid = String(o.value || '').trim();
            ft().adminGroup = gid; touch(setAdminGating(project, !!gid));
          } else if (o.option === 'showMineFirst') {
            ft().scope = on ? 'mine' : 'all';
            for (const s of project.screens) if (/^scrBrowse/.test(s.name) && /Set\(varScope\w+, "(mine|all)"\)/.test(s.onVisible)) { s.onVisible = s.onVisible.replace(/(Set\(varScope\w+, ")(mine|all)("\))/, `$1${on ? 'mine' : 'all'}$3`); touch(s); }
          } else {
            ft()[o.option === 'confirmationPage' ? 'confirm' : 'stepper'] = on;
            notes.push(`${o.option} applies to screens added from now on; existing screens keep their current layout.`);
          }
          break;
        }
        case 'addList': {
          const name = cleanName(o.list, 'Items'); const l = newList(name);
          for (const c of o.columns || []) { const r = makeColumn(l, c); if (r.col) { l.columns.push(r.col); if (r.note) notes.push(r.note); } else notes.push(r.error); }
          project.lists.push(l);
          for (const c of l.columns) if (c.type === 'Lookup' && !project.lists.some((x) => x.name === c.lookupList && x !== l)) { c.type = 'Text'; c.lookupList = ''; notes.push(`${name}.${c.name} pointed at a list that does not exist, so it is a text column.`); }
          break;
        }
        case 'renameList': {
          const before = new Map(project.screens.map((s) => [s.id, JSON.stringify(s)]));
          err = renameListIn(project, o.list, cleanName(o.newName, '')) || '';
          project.screens.forEach((s) => { if (before.get(s.id) !== JSON.stringify(s)) touched.add(s.id); });
          break;
        }
        case 'addColumn': {
          const l = findList(o.list); const r = makeColumn(l, { name: o.column, type: o.type || 'Text', required: o.required, choices: o.choices, lookupList: o.lookupList });
          if (!r.col) { err = r.error; break; }
          if (r.col.type === 'Lookup' && !project.lists.some((x) => x.name === r.col.lookupList && x !== l)) { r.col.type = 'Text'; r.col.lookupList = ''; notes.push(`${r.col.name} pointed at a list that does not exist, so it is a text column.`); }
          l.columns.push(r.col); if (r.note) notes.push(r.note);
          o.__name = r.col.name; // later ops in this batch may use the original name
          break;
        }
        case 'renameColumn': {
          const before = new Map(project.screens.map((s) => [s.id, JSON.stringify(s)]));
          err = renameColumnIn(project, o.list, findCol(findList(o.list), o.column).name, cleanName(o.newName, '')) || '';
          project.screens.forEach((s) => { if (before.get(s.id) !== JSON.stringify(s)) touched.add(s.id); });
          break;
        }
        case 'removeColumn': {
          const before = new Map(project.screens.map((s) => [s.id, JSON.stringify(s)]));
          const r = removeColumnIn(project, o.list, findCol(findList(o.list), o.column).name);
          if (r.error) { err = r.error; break; }
          notes.push(...r.warnings);
          project.screens.forEach((s) => { if (before.get(s.id) !== JSON.stringify(s)) touched.add(s.id); });
          break;
        }
        case 'setChoices': {
          const c = findCol(findList(o.list), o.column);
          c.choices = [...new Set((o.choices || []).map((x) => String(x).trim()).filter(Boolean))];
          touch(refreshStatusVisuals(project, o.list));
          if (project.screens.some((s) => /^scrHome/.test(s.name)) && statusCol(findList(o.list)) === c) notes.push('The home screen counters were generated for the old values; add the home screen again to update them.');
          for (const f of project.flows || []) if (f.list === o.list && f.statusField === c.name) f.statusValues = f.statusValues.filter((v) => c.choices.includes(v));
          break;
        }
        case 'addFieldToForm': case 'addFieldToDetail': case 'addFieldToGallery': {
          const l = findList(o.list), col = colOf(l, ops.find((x) => x.op === 'addColumn' && x.list === o.list && x.column === o.column)?.__name) || findCol(l, o.column);
          if (!col) { err = `No column ${o.list}.${o.column}.`; break; }
          const fn = { addFieldToForm, addFieldToDetail, addFieldToGallery }[o.op];
          const r = fn(project, l.name, col.name, o.screen || '');
          if (r.error) { err = r.error; break; }
          if (r.warning) notes.push(r.warning);
          touch(r.screen || r.screens);
          break;
        }
        case 'addScreenFromTemplate': {
          const l = findList(o.list);
          const first = (kind) => (l ? screensFor(project, l.name, kind)[0] : null);
          let scr;
          if (o.template === 'blank') scr = genBlank(project, 'scrScreen');
          if (o.template === 'loading') { ft().loading = true; const next = project.screens.find((s) => s.id === project.app.startScreen) || project.screens[0]; scr = genLoading(project, { next, lists: project.lists }); project.app.startScreen = scr.id; }
          if (o.template === 'home') {
            const tiles = project.screens.filter((s) => /^scrBrowse/.test(s.name)).map((s) => ({ label: s.name.replace(/^scrBrowse/, '').replace(/([a-z])([A-Z])/g, '$1 $2'), sub: 'Search and open', icon: 'list', actions: [{ type: 'navigate', screen: s.id }] }));
            scr = genHome(project, { tiles });
          }
          if (o.template === 'edit') scr = genEdit(project, l);
          if (o.template === 'detail') scr = genDetail(project, l, { editScreen: first('edit') });
          if (o.template === 'browse') scr = genBrowse(project, l, { editScreen: first('edit'), detailScreen: first('detail'), back: project.screens.some((s) => /^scrHome/.test(s.name)) });
          if (o.template === 'approvals') {
            const st = statusCol(l);
            if (!st) { err = `${l.name} has no Choice column to approve on.`; break; }
            scr = genApprovals(project, l, { field: st.name, back: true, detailScreen: first('detail') });
          }
          touch(scr);
          break;
        }
        case 'addFlow': {
          const f = flowFromSpec(project, { kind: o.flowKind, name: o.flowName, list: o.list, recipients: o.recipients, subject: o.subject, body: o.body, statusField: o.statusField, statusValues: o.statusValues, reminderValue: o.reminderValue, days: o.days, hour: o.hour }, { findList, findCol, notes });
          if (!f) { err = 'Could not build the flow.'; break; }
          (project.flows = project.flows || []).push(f);
          if (f.kind === 'button') { const b = addRunFlowButton(project, f); if (b) touch(project.screens.find((s) => s.controls.includes(b))); }
          break;
        }
        case 'editFlow': {
          const f = project.flows.find((x) => x.name === o.flowName);
          const l = findList(f.list);
          if (o.newName && String(o.newName).trim()) {
            const nn = String(o.newName).trim().slice(0, 80);
            if (project.flows.some((x) => x !== f && flowFxName(x.name) === flowFxName(nn))) { err = `A flow named ${nn} exists.`; break; }
            f.name = nn;
          }
          applyFlowFields(project, f, { recipients: o.recipients, subject: o.subject, body: o.body, statusField: o.statusField, statusValues: o.statusValues, reminderValue: o.reminderValue, days: o.days || undefined, hour: o.hour || undefined }, l, { findCol, notes });
          break;
        }
        case 'setApprovalStages': {
          const l = findList(o.list);
          const col = findCol(l, o.column);
          const field = col?.type === 'Choice' ? col.name : (statusCol(l)?.name || 'Status');
          const stages = stagesFromSpec(l, o.stages, findCol, notes);
          const r = setApprovalChain(project, { list: l.name, field, values: guessChainValues(l, field), stages });
          notes.push(...r.notes); touch(r.screens);
          break;
        }
        case 'removeFlow': {
          const f = project.flows.find((x) => x.name === o.flowName);
          project.flows.splice(project.flows.indexOf(f), 1);
          for (const s of project.screens) for (const [c] of walk(s.controls)) if ((c.actions?.OnSelect || []).some((a) => a.type === 'runFlow' && a.flow === f.id)) notes.push(`${s.name}.${c.name} ran "${f.name}"; delete that button or pick another flow.`);
          break;
        }
      }
    } catch (e) { err = 'Failed: ' + (e?.message || e); }
    if (err) skipped.push({ op: o, reason: err }); else applied.push(o);
  }
  for (const o of ops) delete o.__name;
  return { applied, skipped, notes, touched };
}

/**
 * Turn a list's approvals into a chain (or change its stages): add the chain columns, put any new approver column on the form,
 * rebuild that list's details and approvals screens in place (same ids and names, so links keep working) and replace the chain's flows.
 * Returns {screens, notes}. Screens of other lists are left alone.
 */
export function setApprovalChain(project, spec) {
  const notes = [], screens = [];
  const ch = normalizeChain(spec);
  const l = ch && project.lists.find((x) => x.name === ch.list);
  if (!l) return { screens, notes: ['No list for the approval chain.'] };
  if (project.app.chain && project.app.chain.list !== ch.list) notes.push(`The app's approval chain moved from ${project.app.chain.list} to ${ch.list}; its old screens keep their controls until you rebuild them.`);
  const before = new Set(l.columns.map((c) => c.name));
  ensureChainColumns(project, ch);
  project.app.chain = ch;
  if (project.app.approvalsCfg?.list === ch.list) project.app.approvalsCfg = null;
  // new approver columns go on the form so new items name their approvers
  for (const st of ch.stages) if (st.approver === 'column' && !before.has(st.column)) {
    const r = addFieldToForm(project, l.name, st.column);
    if (r.screen) screens.push(r.screen); else if (r.error) notes.push(r.error);
  }
  const replace = (old, make) => {
    const at = old ? project.screens.indexOf(old) : -1;
    if (old) project.screens.splice(at, 1);
    const scr = make();
    if (old) { project.screens.splice(project.screens.indexOf(scr), 1); scr.id = old.id; scr.name = old.name; project.screens.splice(at, 0, scr); }
    screens.push(scr);
    return scr;
  };
  const edit = screensFor(project, l.name, 'edit')[0];
  const det = replace(screensFor(project, l.name, 'detail')[0], () => genDetail(project, l, { editScreen: edit }));
  const apOld = screensFor(project, l.name, 'approvals')[0];
  replace(apOld, () => genChainApprovals(project, l, ch, { back: project.screens.some((s) => /^scrHome/.test(s.name)), detailScreen: det }));
  if (!apOld) notes.push('A new approvals screen was added; add a tile or button that goes to it.');
  if (project.screens.some((s) => /^scrHome/.test(s.name))) notes.push('The home screen keeps its current tiles and counters; add the home screen again for a "Waiting on me" list.');
  const flows = addChainFlows(project);
  if (flows.length) notes.push(`Added ${flows.length} approval emails: ${flows.map((f) => f.name).join(', ')}.`);
  return { screens, notes };
}
