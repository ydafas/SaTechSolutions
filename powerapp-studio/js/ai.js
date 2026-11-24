// ai.js — "Describe your app": plain-English requirements → Claude → an app plan (lists, screens,
// flows) → a full project. The browser calls the Claude API directly with the visitor's own key
// (anthropic-dangerous-direct-browser-access), asks for JSON that matches PLAN_SCHEMA through
// structured outputs, then cleans the plan in specToProject() before anything is generated.
// No server, no key storage outside the visitor's own browser.
import { newProject, newList, newColumn, clone, COLTYPES, THEME_PRESETS, SP_RESERVED_COLS, slug, newControl } from './core.js';
import { generateApp } from './templates.js';
import { newFlow, FLOW_KINDS, flowFxName, internalName } from './flows.js';
import { chainFlowSpecs, guessChainValues } from './chain.js';

export const API_URL = 'https://api.anthropic.com/v1/messages';
export const API_VERSION = '2026-10-21';
export const MODELS = [
  ['claude-sonnet-5-5', 'Claude Sonnet 5.5 (recommended: fast, strong)'],
  ['claude-opus-5-5', 'Claude Opus 5.5 (most careful, slower, ~2× cost)'],
  ['claude-haiku-5-5', 'Claude Haiku 5.5 (fastest, cheapest)'],
];

// ------------------------------------------------------------------ schema
const obj = (properties) => ({ type: 'object', additionalProperties: false, required: Object.keys(properties), properties });
const str = (description) => ({ type: 'string', ...(description ? { description } : {}) });
const arr = (items, description) => ({ type: 'array', items, ...(description ? { description } : {}) });
/** One approval-chain stage as Claude describes it (plan and change schemas share it). */
export const STAGE_SPEC = obj({
  name: str('Stage name, 1-3 words, e.g. Manager, Finance, Director'),
  approverKind: { type: 'string', enum: ['personColumn', 'group'] },
  approverColumn: str('personColumn: a Person column of the approvals list that names this stage\'s approver. Empty for group.'),
  groupId: str('group: the Microsoft 365 group object ID if the requirements give one; else empty'),
  notifyEmail: str('group: the group\'s shared mailbox if the requirements give one; else empty'),
  slaDays: { type: 'integer', description: 'Days this stage may wait before it is overdue (a daily reminder goes to its approver). 0 when the requirements give no deadline.' },
});
export const PLAN_SCHEMA = obj({
  appName: str('Short app name, 2-4 words'),
  formFactor: { type: 'string', enum: ['tablet', 'phone'] },
  theme: { type: 'string', enum: THEME_PRESETS.map((t) => t.name) },
  lists: arr(obj({
    name: str('SharePoint list name, plural, PascalCase, no spaces, e.g. TravelRequests'),
    purpose: str('One sentence: what one item in this list is'),
    columns: arr(obj({
      name: str('PascalCase, no spaces. Never Title, ID, Created, Modified, Author or Editor.'),
      type: { type: 'string', enum: Object.keys(COLTYPES) },
      required: { type: 'boolean' },
      choices: arr(str(), 'Choice columns only; first choice is the starting value. Empty for other types.'),
      lookupList: str('Lookup columns only: name of another list in this plan. Empty otherwise.'),
    })),
    attachments: { type: 'boolean', description: 'True only when people attach files (documents, photos, receipts) to items of this list.' },
  })),
  screens: obj({
    home: { type: 'boolean' }, browse: { type: 'boolean' }, detail: { type: 'boolean' }, edit: { type: 'boolean' },
    approvals: obj({ enabled: { type: 'boolean' }, list: str(), field: str('A Choice column'), pending: str(), approve: str(), reject: str(),
      stages: arr(STAGE_SPEC, 'Only for 2-4 sequential approval steps (e.g. manager then finance). Empty for a single approve/reject step.'),
      reassign: { type: 'boolean', description: 'Stages: true when an approver may hand a waiting stage to someone else (delegate).' },
      adminOverride: { type: 'boolean', description: 'Stages: true when admins may approve or reject a waiting stage with a reason. Needs an admin group.' },
    }),
  }),
  flows: arr(obj({
    kind: { type: 'string', enum: Object.keys(FLOW_KINDS) },
    name: str('Short flow name'),
    list: str(),
    recipients: arr(obj({ kind: { type: 'string', enum: ['column', 'author', 'editor', 'me', 'address'] }, value: str('Person column name for column; email address for address; empty otherwise') })),
    subject: str('May use {ColumnName} tokens'),
    body: str('Plain text with line breaks; may use tokens'),
    statusField: str('statusChanged: the Choice column to watch. reminder: column to filter on. Else empty.'),
    statusValues: arr(str(), 'statusChanged: values that trigger the email'),
    reminderValue: str('reminder: value of statusField that means "still waiting"'),
    days: { type: 'integer', description: 'reminder: minimum age in days' },
    hour: { type: 'integer', description: 'reminder: hour of day 0-23 to send' },
  })),
  options: obj({
    loadingScreen: { type: 'boolean', description: 'Animated loading screen with the app monogram' },
    confirmationPage: { type: 'boolean', description: 'Show a "submitted" page after a new item is saved' },
    statusStepper: { type: 'boolean', description: 'Show progress through the status values on the details screen' },
    showMineFirst: { type: 'boolean', description: 'Lists open filtered to the signed-in person\'s own items' },
    trackUsage: { type: 'boolean', description: 'Record one row per user visit in an AppUsage list' },
    adminGroupId: str('Microsoft 365 group object ID (GUID) for admins, only if the requirements give one; else empty'),
  }),
  assumptions: arr(str(), 'Each guess you made where the requirements were silent'),
  questions: arr(str(), 'Open questions whose answers would change the app'),
});

export const SYSTEM_PROMPT = `You design Microsoft Power Apps canvas apps backed by SharePoint lists, plus Power Automate email flows, from a person's plain-English requirements. Your output is a plan that a generator turns into screens and flows, so stay inside what the generator can build.

What the generator can build:
- 1 to 4 SharePoint lists. Every list already has a Title column (single line of text, required): use it for each item's main name and never add Title, ID, Created, Modified, Author (Created By) or Editor (Modified By) yourself; flows can still read {Title}, {ID}, {Created} and {Author}.
- Column types: Text (single line), Note (multi-line), Number, Currency, DateTime (date only), Choice (single choice, list the choices; the first is the starting value), YesNo, Person (one person), Lookup (to another list in this plan; it shows that list's Title).
- Column names: PascalCase, letters and digits only, no spaces, and not names SharePoint already uses such as Status0, StartDate, EndDate, Location, Order, Type, Name, Path, ContentType, Attachments. A workflow state column called Status is fine.
- Screens per list: browse (search + gallery), detail (all columns, Edit, Delete) and edit (form for new and existing items, required columns validated); a home dashboard with counts of the first list's Status-like Choice column and tiles to every screen. Turn on approvals only when the requirements describe someone approving or rejecting items: name the list, the Choice column, the waiting value, the approve value and the reject value, all exactly as you defined them. When the requirements describe 2 to 4 approvals in sequence (for example the manager, then finance, then a director), also list them in order as approvals.stages: a stage approved by a named person needs its own Person column in that list (e.g. ManagerApprover) given as approverColumn; a stage approved by anyone in a team uses approverKind group. The generator then adds the per-stage decision, date, approver and comment columns, a CurrentStage number and an ApprovalHistory note itself, so do not define those, and do not add flows for the stage hand-offs or the final outcome email: the generator adds them too. The person who created an item can never approve it. The creator's manager cannot be looked up automatically: use a Person column.
- Flows (kinds): created = email when a new item is added; statusChanged = email when a Choice column changes to one of statusValues; reminder = every day at "hour", email for each item whose statusField equals reminderValue and that is at least "days" old; button = a button on the item's detail screen runs the flow for that item.
- Flow recipients: column (a Person column of that list), author (who created the item), editor (who last changed it), me (button flows only: whoever pressed the button), address (a fixed email address the requirements give; never invent addresses).
- Email tokens in subject and body: {ColumnName} for any column of the flow's list, {Title}, {ID}, {Created}, {Author}, {PersonColumn.Email}, and {Link}, a link that opens the item in the app. Keep emails short and plain.

How to plan:
- Model exactly what the requirements ask for and what they clearly need to work; do not invent extra features. Prefer fewer, well-named columns.
- When the requirements are silent on something you must decide, choose the conventional option and record it in assumptions.
- Put anything genuinely ambiguous that would change the app into questions (at most 5). Still return a complete plan.
- Pick formFactor phone only when the requirements mention phones, mobile or field use. Pick the theme that fits the audience: Federal Navy for government or defense, Fluent Blue for general Microsoft 365, Midnight for dark, otherwise SaTech Cyan.
- Options: loadingScreen true unless asked for a plain app; confirmationPage true when people submit things (requests, registrations); statusStepper true when a Status choice column moves through stages; showMineFirst true when people mostly track their own items (requests, tickets) and false for shared catalogues (assets, contacts); trackUsage only when asked; adminGroupId only when a group ID is given. When an admin group is set, only admins see approve, reject and delete.
- Use empty strings, empty arrays, 0 and false for fields that do not apply.`;

/** Request body for the Messages API. history: prior [{role, content}] turns for refinements. */
export function buildRequest(text, model = MODELS[0][0], history = []) {
  const messages = history.length ? [...history] : [{ role: 'user', content: `Requirements:\n\n${String(text).trim()}` }];
  return { model, max_tokens: 16000, system: SYSTEM_PROMPT, messages, output_config: { format: { type: 'json_schema', schema: PLAN_SCHEMA } } };
}
export function refineHistory(text, planJson, change) {
  return [
    { role: 'user', content: `Requirements:\n\n${String(text).trim()}` },
    { role: 'assistant', content: planJson },
    { role: 'user', content: `Revise the plan with these changes, and return the complete updated plan:\n\n${String(change).trim()}` },
  ];
}

export class AiError extends Error {}
/** Call Claude from the browser. Returns {plan, raw, usage}. */
export async function callClaude({ key, model, text, history, request, signal, fetchFn = globalThis.fetch }) {
  if (!/^sk-ant-/.test(String(key || '').trim())) throw new AiError('Enter an Anthropic API key (it starts with sk-ant-).');
  let res;
  try {
    res = await fetchFn(API_URL, {
      method: 'POST', signal,
      headers: { 'content-type': 'application/json', 'x-api-key': key.trim(), 'anthropic-version': API_VERSION, 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify(request || buildRequest(text, model, history)),
    });
  } catch (e) {
    if (e?.name === 'AbortError') throw new AiError('Cancelled.');
    throw new AiError('Could not reach api.anthropic.com. Your network may block it (common on government networks), or you are offline.');
  }
  let body = null;
  try { body = await res.json(); } catch { /* not JSON */ }
  if (!res.ok) {
    const msg = body?.error?.message || res.statusText || 'request failed';
    if (res.status === 401) throw new AiError('The API key was rejected. Check it in the Claude Console.');
    if (res.status === 429) throw new AiError('Rate limited by the API. Wait a minute and try again.');
    if (res.status === 529 || res.status >= 500) throw new AiError('The Claude API is busy. Try again shortly.');
    throw new AiError(`Claude API error ${res.status}: ${msg}`);
  }
  return { ...parseResponse(body), usage: body?.usage || null };
}
export function parseResponse(body) {
  if (body?.stop_reason === 'refusal') throw new AiError('Claude declined to plan this request.');
  if (body?.stop_reason === 'max_tokens') throw new AiError('The plan was too long and got cut off. Shorten the requirements or split the app.');
  const text = (body?.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
  let plan;
  try { plan = JSON.parse(text); } catch { throw new AiError('The response was not a valid plan. Try again.'); }
  return { plan, raw: text };
}

// ------------------------------------------------------------------ plan → project
const BUILTINS = new Set(['Title', 'ID', 'Id', 'Created', 'Modified', 'Author', 'Editor', 'CreatedBy', 'ModifiedBy']);
const cleanName = (n, fallback) => String(n || '').replace(/[^A-Za-z0-9]+(.)?/g, (_, c) => (c ? c.toUpperCase() : '')).replace(/^[^A-Za-z]+/, '') || fallback;
const isEmail = (s) => /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/.test(String(s || '').trim());

/** Clean an AI plan and build a project from it. Returns {project, notes} — notes say what was fixed. */
export function specToProject(plan, { siteUrl = '' } = {}) {
  const notes = [];
  const p = newProject();
  p.app.name = String(plan?.appName || '').trim().slice(0, 60) || 'My App';
  p.app.formFactor = plan?.formFactor === 'phone' ? 'phone' : 'tablet';
  p.theme = clone(THEME_PRESETS.find((t) => t.name === plan?.theme) || THEME_PRESETS[0]);
  p.app.siteUrl = String(siteUrl || '').trim();

  // lists and columns
  const listMap = {}, colMap = {};
  for (const rl of (plan?.lists || []).slice(0, 6)) {
    let name = cleanName(rl.name, 'Items'), i = 2;
    while (p.lists.some((l) => l.name === name)) name = cleanName(rl.name, 'Items') + i++;
    if (name !== rl.name) notes.push(`List "${rl.name}" renamed to ${name}.`);
    listMap[rl.name] = name; listMap[name] = name;
    const l = newList(name);
    colMap[name] = {};
    const single = name.replace(/ies$/, 'y').replace(/s$/, '');
    for (const rc of rl.columns || []) {
      let cn = cleanName(rc.name, 'Field');
      if (BUILTINS.has(cn)) { if (cn === 'Title' && rc.required === false) l.columns[0].required = false; colMap[name][rc.name] = cn === 'Title' ? 'Title' : cn; continue; }
      if (SP_RESERVED_COLS.has(cn)) { const nn = single + cn; notes.push(`${name}.${cn} renamed to ${nn} (SharePoint already uses ${cn}).`); cn = nn; }
      let k = 2; const base = cn; while (l.columns.some((c) => c.name === cn)) cn = base + k++;
      let type = COLTYPES[rc.type] ? rc.type : 'Text';
      let choices = (rc.choices || []).map((x) => String(x).trim()).filter(Boolean);
      choices = [...new Set(choices)];
      if (type === 'Choice' && !choices.length) { type = 'Text'; notes.push(`${name}.${cn} had no choices, so it is a text column.`); }
      if (type !== 'Choice') choices = [];
      colMap[name][rc.name] = cn; colMap[name][cn] = cn;
      l.columns.push(newColumn(cn, type, { required: type === 'YesNo' ? false : !!rc.required, choices, lookupList: type === 'Lookup' ? String(rc.lookupList || '') : '' }));
    }
    if (rl.attachments === true) l.attachments = true;
    p.lists.push(l);
  }
  if (!p.lists.length) { p.lists.push(newList('Items')); notes.push('The plan had no lists, so an empty Items list was added.'); }
  for (const l of p.lists) for (const c of l.columns) if (c.type === 'Lookup') {
    const target = listMap[c.lookupList] || cleanName(c.lookupList, '');
    if (p.lists.some((x) => x.name === target && x !== l)) c.lookupList = target;
    else { c.type = 'Text'; c.lookupList = ''; notes.push(`${l.name}.${c.name} pointed at a list that is not in the plan, so it is a text column.`); }
  }
  const findList = (n) => p.lists.find((l) => l.name === (listMap[n] || n));
  const findCol = (l, n) => l && l.columns.find((c) => c.name === (colMap[l.name]?.[n] || cleanName(n, '')));

  // screens and approvals
  const sc = plan?.screens || {};
  const screens = ['home', 'browse', 'detail', 'edit'].filter((k) => sc[k] !== false);
  if (!screens.length) screens.push('browse', 'edit');
  let approvals = null;
  const ap = sc.approvals;
  if (ap?.enabled) {
    const l = findList(ap.list), col = findCol(l, ap.field);
    if (l && col?.type === 'Choice') {
      const pick = (v, re, i) => (col.choices.includes(v) ? v : col.choices.find((c) => re.test(c)) || col.choices[i] || '');
      approvals = { list: l.name, field: col.name, pending: pick(ap.pending, /pend|submit|wait|review|new/i, 0), yes: pick(ap.approve, /approv|accept/i, 1), no: pick(ap.reject, /reject|deny|declin/i, 2) };
    } else notes.push('Approvals were requested but no matching Choice column was found, so the approvals screen was left out.');
  }
  const op = plan?.options || {};
  const gid = String(op.adminGroupId || '').trim();
  const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(gid);
  if (gid && !isGuid) notes.push(`"${gid}" is not a group ID (it should look like 00000000-0000-0000-0000-000000000000), so admin-only gating was left off.`);
  p.app.features = { loading: op.loadingScreen !== false, confirm: op.confirmationPage !== false, stepper: op.statusStepper !== false, scope: op.showMineFirst ? 'mine' : 'all', usage: !!op.trackUsage, adminGroup: isGuid ? gid : '' };
  // 2-4 sequential stages → an approval chain (one stage stays the classic approvals queue)
  let chain = null;
  const rs = Array.isArray(ap?.stages) ? ap.stages.slice(0, 4) : [];
  if (approvals && rs.length >= 2) {
    const stages = stagesFromSpec(findList(approvals.list), rs, findCol, notes);
    chain = { list: approvals.list, field: approvals.field, values: { ...guessChainValues(findList(approvals.list), approvals.field), pending: approvals.pending, approved: approvals.yes, rejected: approvals.no }, stages, reassign: ap.reassign === true, override: ap.adminOverride === true };
    if (chain.override && !p.app.features.adminGroup) notes.push('Admin override needs an admin group; add its ID on the App tab, or lint flags the chain.');
  }
  generateApp(p, { screens, approvals, chain });
  if (p.app.chain) addChainFlows(p);

  // flows
  const usedNames = new Set((p.flows || []).map((f) => flowFxName(f.name)));
  for (const rf of (plan?.flows || []).slice(0, 12)) {
    const f = flowFromSpec(p, rf, { findList, findCol, usedNames, notes });
    if (!f) continue;
    p.flows.push(f);
    if (f.kind === 'button') addRunFlowButton(p, f);
  }
  // email links open the item's detail screen
  const linked = p.flows.find((f) => /\{Link\}/.test(f.subject + f.body));
  if (linked) {
    const det = p.screens.find((s) => s.name === 'scrDetail' + slug(linked.list));
    if (det) p.app.deepLink = { screen: det.id, list: linked.list };
  }
  if (p.flows.length && !p.app.siteUrl) notes.push('Set the SharePoint site URL on the Flows tab before downloading flows.');
  return { project: p, notes };
}

/** Build a cleaned flow from a plan/edit spec. Lookups: findList(name), findCol(list, name). Returns the flow (not added). */
export function flowFromSpec(p, rf, { findList, findCol, usedNames = new Set((p.flows || []).map((f) => flowFxName(f.name))), notes = [] }) {
  if (!FLOW_KINDS[rf.kind]) return null;
  const l = findList(rf.list) || p.lists[0];
  if (!l) return null;
  const f = newFlow(p, rf.kind, l.name);
  let nm = String(rf.name || f.name).trim().slice(0, 80) || f.name;
  if (rf.kind === 'button') nm = cleanName(nm, 'RunFlow');
  let n2 = nm, i = 2; while (usedNames.has(flowFxName(n2))) n2 = `${nm} ${i++}`;
  f.name = n2; usedNames.add(flowFxName(n2));
  applyFlowFields(p, f, rf, l, { findCol, notes });
  return f;
}
/** Copy subject, body, recipients and trigger settings from a spec onto a flow, cleaning each. Empty fields are left as they are. */
export function applyFlowFields(p, f, rf, l, { findCol, notes = [] }) {
  const tokenFix = (text) => String(text || '').replace(/\{([^}]+)\}/g, (m, tok) => {
    const [n, sub] = tok.split('.');
    if (['Title', 'ID', 'Created', 'Modified', 'Author', 'Editor', 'Link', 'Me'].includes(n)) return m;
    const c = findCol(l, n);
    return c ? `{${internalName(c.name)}${sub ? '.' + sub : ''}}` : m;
  });
  if (rf.subject) f.subject = tokenFix(rf.subject);
  if (rf.body) f.body = tokenFix(rf.body);
  const rec = [];
  for (const r of rf.recipients || []) {
    if (r.kind === 'column') { const c = findCol(l, r.value); if (c?.type === 'Person') rec.push({ kind: 'column', value: c.name }); else notes.push(`Flow "${f.name}": "${r.value}" is not a Person column of ${l.name}; recipient dropped.`); }
    else if (r.kind === 'address') { const em = String(r.value || '').split(/[;,\s]+/).filter(isEmail); em.forEach((e) => rec.push({ kind: 'address', value: e })); if (!em.length) notes.push(`Flow "${f.name}": "${r.value}" is not an email address; recipient dropped.`); }
    else if (r.kind === 'me' && f.kind !== 'button') rec.push({ kind: 'author', value: '' });
    else if (['author', 'editor', 'me'].includes(r.kind)) rec.push({ kind: r.kind, value: '' });
  }
  if (rec.length) f.recipients = rec;
  if (f.kind === 'statusChanged' || f.kind === 'reminder') {
    const c = rf.statusField ? findCol(l, rf.statusField) : null;
    if (f.kind === 'statusChanged') {
      if (c?.type === 'Choice') f.statusField = c.name;
      const col = l.columns.find((x) => x.name === f.statusField);
      if ((rf.statusValues || []).length || rf.statusField) f.statusValues = (rf.statusValues || []).filter((v) => col?.choices.includes(v));
      if (!col) notes.push(`Flow "${f.name}" needs a Choice column to watch; pick one on the Flows tab.`);
    } else {
      if (c) f.statusField = c.name;
      const col = l.columns.find((x) => x.name === f.statusField);
      if (rf.reminderValue !== undefined && rf.reminderValue !== '' || rf.statusField) f.reminderValue = col?.type === 'Choice' ? (col.choices.includes(rf.reminderValue) ? rf.reminderValue : col.choices[0]) : String(rf.reminderValue ?? '');
      if (rf.days !== undefined && (rf.days || rf.statusField)) f.days = Math.max(0, Math.min(365, Math.round(Number(rf.days) || 0)));
      if (rf.hour !== undefined && (rf.hour || rf.statusField)) f.hour = Math.max(0, Math.min(23, Math.round(Number(rf.hour) || 8)));
    }
  }
  return f;
}

/** Approval stages from Claude's STAGE_SPEC objects → chain stages (cleaned; bad IDs and columns become notes). */
export function stagesFromSpec(l, rs, findCol, notes = []) {
  const sla = (st) => (Math.round(Number(st?.slaDays)) > 0 ? { slaDays: Math.min(90, Math.round(Number(st.slaDays))) } : {});
  return (rs || []).slice(0, 4).map((st, i) => {
    const name = String(st?.name || `Stage ${i + 1}`).trim();
    if (st?.approverKind === 'group') {
      const g = String(st.groupId || '').trim(), ok = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(g);
      if (g && !ok) notes.push(`Stage "${name}": "${g}" is not a group ID; add it in Finish setup.`);
      const em = String(st.notifyEmail || '').trim();
      return { name, approver: 'group', groupId: ok ? g : '', notify: isEmail(em) ? em : '', ...sla(st) };
    }
    const c = findCol(l, st?.approverColumn);
    if (st?.approverColumn && c?.type !== 'Person') notes.push(`Stage "${name}": "${st.approverColumn}" is not a Person column of ${l.name}, so a new approver column was added.`);
    return { name, approver: 'column', column: c?.type === 'Person' ? c.name : '', ...sla(st) };
  });
}
/** Add the approval chain's flows (one per stage hand-off, a resubmit notice and the final outcome email). Replaces earlier chain flows for the list. */
export function addChainFlows(project) {
  const ch = project.app.chain;
  if (!ch) return [];
  const l = project.lists.find((x) => x.name === ch.list);
  if (!l) return [];
  project.flows = (project.flows || []).filter((f) => !(f.chain && f.list === ch.list));
  const findList = () => l, findCol = (list, n) => list.columns.find((c) => c.name === n);
  const out = [];
  for (const spec of chainFlowSpecs(ch)) {
    const f = flowFromSpec(project, { ...spec, list: l.name }, { findList, findCol });
    if (!f) continue;
    f.chain = true;
    if (spec.chainStage) f.chainStage = spec.chainStage;
    project.flows.push(f); out.push(f);
  }
  // the emails' {Link} opens the item's details
  const det = project.screens.find((s) => s.name === 'scrDetail' + slug(ch.list));
  if (det && !project.app.deepLink?.screen) project.app.deepLink = { screen: det.id, list: ch.list };
  return out;
}

/** Put a "Run a flow" button on the flow's detail screen: in the header beside Edit (tablet), above the action bar (phone). */
export function addRunFlowButton(project, flow) {
  const det = project.screens.find((s) => s.name === 'scrDetail' + slug(flow.list));
  if (!det) return null;
  const phone = project.app.formFactor === 'phone';
  const W = phone ? 640 : 1366, H = phone ? 1136 : 768, HH = phone ? 64 : 72;
  const label = String(flow.name).replace(/([a-z])([A-Z])/g, '$1 $2').slice(0, 26);
  const V = 'varSel' + slug(flow.list);
  let x, y, w, props;
  if (phone) {
    const runs = det.controls.filter((c) => c.name.startsWith('btnRun')).length;
    w = W - 32; x = 16; y = H - 72 - 52 - runs * 52;
    props = { text: label, fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1, size: 12 };
  } else {
    const inHeader = det.controls.filter((c) => c.type === 'button' && c.y === (HH - 40) / 2);
    const leftMost = inHeader.reduce((m, c) => Math.min(m, c.x), W - 24);
    w = 180; x = leftMost - 10 - w; y = (HH - 40) / 2;
    props = { text: label, fill: 'transparent', color: '@OnHeader', borderColor: '@OnHeader', borderThickness: 1, size: 12 };
  }
  const btn = newControl(project, 'button', {
    name: 'btnRun' + flowFxName(flow.name), x, y, w, h: 40, props,
    actions: { OnSelect: [{ type: 'runFlow', flow: flow.id, message: 'Done', then: '' }] }, fx: { Visible: `!IsBlank(${V}.ID)` },
  });
  det.controls.push(btn);
  return btn;
}
