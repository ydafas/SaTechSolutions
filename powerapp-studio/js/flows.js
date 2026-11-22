// flows.js — Power Automate flows for the generated app, exported as legacy flow packages (.zip)
// that import with My flows → Import → Import Package (Legacy) → Create as new.
//
// Package layout (as Power Automate exports it, see the powerapps-canvas-yaml skill §12):
//   manifest.json                                   package resources: the flow, each connector, each connection
//   Microsoft.Flow/flows/manifest.json              {"packageSchemaVersion":"1.0","flowAssets":{"assetPaths":[<flowKey>]}}
//   Microsoft.Flow/flows/<flowKey>/definition.json  the Logic Apps workflow definition (minified)
//   Microsoft.Flow/flows/<flowKey>/apisMap.json     connector name → api resource key
//   Microsoft.Flow/flows/<flowKey>/connectionsMap.json  connector name → connection resource key
// On import Power Automate asks the user to pick a connection for each connection resource, so no
// connection ids from any tenant are needed. A real export can be loaded as a "tenant sample" to copy
// that tenant's exact resource details.
import { uid, slug } from './core.js';

export const internalName = (n) => String(n).replace(/[^A-Za-z0-9_]/g, '') || 'Field';
/** The name Power Apps gives a flow added as a data source: its display name without spaces or symbols. */
export const flowFxName = (name) => (String(name || '').replace(/[^A-Za-z0-9_]/g, '') || 'MyFlow').replace(/^(\d)/, '_$1');

export const FLOW_KINDS = {
  created: { label: 'When an item is created', blurb: 'Email someone each time a new item is added: a new request for the approver, a new task for the assignee.', icon: 'plus' },
  statusChanged: { label: 'When a column changes', blurb: 'Email when a Choice column such as Status changes to certain values: tell the requester it was approved or rejected.', icon: 'refresh' },
  reminder: { label: 'Daily reminder', blurb: 'Every morning, email about items that have sat in one state for too long.', icon: 'clock' },
  button: { label: 'Run from a button', blurb: 'A button in the app runs this flow for the selected item and waits for it to finish.', icon: 'send' },
};
export const TIME_ZONES = ['Eastern Standard Time', 'Central Standard Time', 'Mountain Standard Time', 'Pacific Standard Time', 'Hawaiian Standard Time', 'Alaskan Standard Time', 'UTC'];

const CONNECTORS = {
  shared_sharepointonline: { displayName: 'SharePoint', iconUri: 'https://connectoricons-prod.azureedge.net/releases/v1.0.1676/1.0.1676.3617/sharepointonline/icon.png' },
  shared_office365: { displayName: 'Office 365 Outlook', iconUri: 'https://connectoricons-prod.azureedge.net/releases/v1.0.1680/1.0.1680.3652/office365/icon.png' },
};

// ------------------------------------------------------------------ model
export function newFlow(project, kind, list) {
  const l = project.lists.find((x) => x.name === list) || project.lists[0];
  const st = l?.columns.find((c) => c.type === 'Choice' && /status/i.test(c.name)) || l?.columns.find((c) => c.type === 'Choice');
  const person = l?.columns.find((c) => c.type === 'Person');
  const single = (l?.name || 'item').replace(/ies$/, 'y').replace(/s$/, '').toLowerCase();
  const base = { id: uid(), kind, list: l?.name || '', enabled: true, recipients: [], subject: '', body: '', statusField: st?.name || '', statusValues: [], days: 3, hour: 8, timeZone: 'Central Standard Time', reminderValue: '' };
  switch (kind) {
    case 'created':
      return { ...base, name: `Notify on new ${single}`, recipients: [person ? { kind: 'column', value: person.name } : { kind: 'address', value: '' }],
        subject: `New ${single}: {Title}`, body: `Hello,\n\n{Author} submitted a new ${single}: {Title}.\n\nOpen it in the app: {Link}` };
    case 'statusChanged': {
      const done = (st?.choices || []).filter((c) => /approv|reject|done|complete|cancel/i.test(c));
      return { ...base, name: `Notify when ${st?.name || 'status'} changes`, statusValues: done.length ? done : (st?.choices || []).slice(1),
        recipients: [{ kind: 'author', value: '' }], subject: `Your ${single} is {${internalName(st?.name || 'Status')}}: {Title}`,
        body: `Hello {Author},\n\nYour ${single} "{Title}" is now {${internalName(st?.name || 'Status')}}.\n\nOpen it in the app: {Link}` };
    }
    case 'reminder':
      return { ...base, name: `Daily reminder for waiting ${l?.name?.toLowerCase() || 'items'}`, reminderValue: st?.choices?.[0] || '',
        recipients: [person ? { kind: 'column', value: person.name } : { kind: 'author', value: '' }],
        subject: `Reminder: {Title} is still {${internalName(st?.name || 'Status')}}`, body: `Hello,\n\n"{Title}" has been waiting since {Created}.\n\nOpen it in the app: {Link}` };
    case 'button':
      return { ...base, name: `Send ${single} summary`, recipients: [{ kind: 'me', value: '' }],
        subject: `${single[0].toUpperCase() + single.slice(1)} summary: {Title}`, body: `Here are the details for "{Title}" (ID {ID}).\n\nOpen it in the app: {Link}` };
  }
  return base;
}

// ------------------------------------------------------------------ expressions
/** Where the item's fields come from in each kind of flow. */
function source(kind) {
  if (kind === 'reminder') return { path: (parts) => `items('Apply_to_each')?${parts.map((p) => `['${p}']`).join('?')}` };
  if (kind === 'button') return { path: (parts) => `outputs('Get_item')?['body/${parts.join('/')}']` };
  return { path: (parts) => `triggerOutputs()?['body/${parts.join('/')}']` };
}
const BUILTIN = {
  ID: { type: 'Number', path: ['ID'] }, Created: { type: 'DateTime', path: ['Created'] }, Modified: { type: 'DateTime', path: ['Modified'] },
  Author: { type: 'Person', path: ['Author'] }, Editor: { type: 'Person', path: ['Editor'] },
};
/** Expression (without @{}) for a token like Title, Status, Approver, Approver.Email, Author, Link. */
export function tokenExpr(project, flow, token) {
  const [name, sub] = String(token).split('.');
  const src = source(flow.kind);
  if (name === 'Link') {
    const link = String(project.app.appLink || '').trim();
    const sep = link.includes('?') ? '&' : '?';
    return link ? `concat('${link.replace(/'/g, "''")}${sep}ID=', string(${src.path(['ID'])}))` : `string(${src.path(['ID'])})`;
  }
  if (name === 'Me' && flow.kind === 'button') return sub === 'Email' ? `triggerOutputs()?['headers']?['x-ms-user-email']` : `triggerOutputs()?['headers']?['x-ms-user-name']`;
  const list = project.lists.find((l) => l.name === flow.list);
  const col = BUILTIN[name] ? { type: BUILTIN[name].type, name } : list?.columns.find((c) => c.name === name || internalName(c.name) === name);
  if (!col) return null;
  const iname = BUILTIN[name] ? name : internalName(col.name);
  const ref = (parts) => src.path([iname, ...parts]);
  switch (col.type) {
    case 'Choice': case 'Lookup': return ref(['Value']);
    case 'Person': return ref([sub === 'Email' ? 'Email' : 'DisplayName']);
    case 'DateTime': return `if(empty(${ref([])}), '', formatDateTime(${ref([])}, 'MM/dd/yyyy'))`;
    case 'YesNo': return `if(equals(${ref([])}, true), 'Yes', 'No')`;
    default: return ref([]);
  }
}
const TOKEN = /\{([A-Za-z_][A-Za-z0-9_]*(?:\.Email)?)\}/g;
/** Tokens used in a text that do not resolve. */
export function badTokens(project, flow, text) {
  return [...String(text || '').matchAll(TOKEN)].map((m) => m[1]).filter((t) => tokenExpr(project, flow, t) == null);
}
const escHtml = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
/** Template text → Logic Apps string with @{...} interpolations. html=true turns lines into <p>/<br>. */
export function interpolate(project, flow, text, html = false) {
  const out = [];
  let last = 0;
  // Logic Apps: only '@{' starts an interpolation inside a string, and '@@{' is its escape; a lone @ (an email
  // address) stays literal. A string that begins with @ is an expression, so a leading literal @ is doubled below.
  const lit = (t) => { out.push((html ? escHtml(t).replace(/\n/g, '<br>') : t).replace(/@\{/g, '@@{')); };
  for (const m of String(text || '').matchAll(TOKEN)) {
    lit(text.slice(last, m.index));
    const e = tokenExpr(project, flow, m[1]);
    if (e == null) lit(m[0]);
    else if (html && m[1] === 'Link' && project.app.appLink) out.push(`<a href="@{${e}}">Open in ${escHtml(project.app.name || 'the app')}</a>`);
    else out.push(`@{${e}}`);
    last = m.index + m[0].length;
  }
  lit(String(text || '').slice(last));
  let s = out.join('');
  if (html) s = `<p>${s}</p>`;
  else if (s.startsWith('@') && !s.startsWith('@{')) s = '@' + s;
  return s;
}
/** "To" field: semicolon-separated addresses. */
export function recipientsExpr(project, flow) {
  const parts = [];
  for (const r of flow.recipients || []) {
    if (r.kind === 'address' && r.value) parts.push(String(r.value).trim());
    if (r.kind === 'author') parts.push(`@{${tokenExpr(project, flow, 'Author.Email')}}`);
    if (r.kind === 'editor') parts.push(`@{${tokenExpr(project, flow, 'Editor.Email')}}`);
    if (r.kind === 'me' && flow.kind === 'button') parts.push(`@{${tokenExpr(project, flow, 'Me.Email')}}`);
    if (r.kind === 'column' && r.value) { const e = tokenExpr(project, flow, r.value + '.Email'); if (e) parts.push(`@{${e}}`); }
  }
  return parts.join(';');
}
/** Readable recipient summary for Play mode and previews. */
export function recipientsText(flow) {
  return (flow.recipients || []).map((r) => (r.kind === 'address' ? r.value : r.kind === 'author' ? 'Created By' : r.kind === 'editor' ? 'Modified By' : r.kind === 'me' ? 'the person who pressed the button' : r.value)).filter(Boolean).join('; ') || '(nobody)';
}
const odataValue = (col, v) => (col?.type === 'YesNo' ? (/^(true|yes|1)$/i.test(String(v)) ? '1' : '0') : col?.type === 'Number' || col?.type === 'Currency' ? String(Number(v) || 0) : `'${String(v ?? '').replace(/'/g, "''")}'`);

// ------------------------------------------------------------------ definition
const SP = 'shared_sharepointonline', O365 = 'shared_office365';
const host = (conn, op) => ({ connectionName: conn, operationId: op, apiId: `/providers/Microsoft.PowerApps/apis/${conn}` });
const spParams = (project, flow, extra = {}) => ({ dataset: String(project.app.siteUrl || '').trim().replace(/\/+$/, ''), table: flow.list, ...extra });

function sendEmail(project, flow, runAfter) {
  return {
    runAfter, type: 'OpenApiConnection',
    inputs: {
      host: host(O365, 'SendEmailV2'),
      parameters: {
        'emailMessage/To': recipientsExpr(project, flow),
        'emailMessage/Subject': interpolate(project, flow, flow.subject),
        'emailMessage/Body': interpolate(project, flow, flow.body, true),
        'emailMessage/Importance': 'Normal',
      },
      authentication: "@parameters('$authentication')",
    },
  };
}

export function flowDefinition(project, flow) {
  const triggers = {}, actions = {};
  const list = project.lists.find((l) => l.name === flow.list);
  if (flow.kind === 'created') {
    triggers.When_an_item_is_created = {
      recurrence: { interval: 1, frequency: 'Minute' }, splitOn: "@triggerOutputs()?['body/value']", type: 'OpenApiConnection',
      inputs: { host: host(SP, 'GetOnNewItems'), parameters: spParams(project, flow), authentication: "@parameters('$authentication')" },
    };
    actions['Send_an_email_(V2)'] = sendEmail(project, flow, {});
  }
  if (flow.kind === 'statusChanged') {
    const iname = internalName(flow.statusField || 'Status');
    triggers.When_an_item_is_created_or_modified = {
      recurrence: { interval: 1, frequency: 'Minute' }, splitOn: "@triggerOutputs()?['body/value']", type: 'OpenApiConnection',
      inputs: { host: host(SP, 'GetOnUpdatedItems'), parameters: spParams(project, flow), authentication: "@parameters('$authentication')" },
    };
    actions.Get_changes = {
      runAfter: {}, type: 'OpenApiConnection',
      inputs: {
        host: host(SP, 'GetItemChanges'),
        parameters: spParams(project, flow, { id: "@triggerOutputs()?['body/ID']", since: "@triggerOutputs()?['body/{TriggerWindowStartToken}']", includeDrafts: false }),
        authentication: "@parameters('$authentication')",
      },
    };
    const changed = { equals: [`@body('Get_changes')?['ColumnHasChanged']?['${iname}']`, true] };
    const vals = (flow.statusValues || []).filter(Boolean);
    const valueRef = `@triggerOutputs()?['body/${iname}/Value']`;
    actions.Condition = {
      actions: { 'Send_an_email_(V2)': sendEmail(project, flow, {}) }, else: { actions: {} },
      expression: vals.length ? { and: [changed, { or: vals.map((v) => ({ equals: [valueRef, v] })) }] } : { and: [changed] },
      runAfter: { Get_changes: ['Succeeded'] }, type: 'If',
    };
  }
  if (flow.kind === 'reminder') {
    triggers.Recurrence = { recurrence: { frequency: 'Day', interval: 1, schedule: { hours: [String(Number(flow.hour) || 8)], minutes: [0] }, timeZone: flow.timeZone || 'UTC' }, type: 'Recurrence' };
    const days = Math.max(0, Number(flow.days) || 0);
    const filters = [];
    if (flow.statusField && flow.reminderValue !== '') {
      const col = list?.columns.find((c) => c.name === flow.statusField);
      filters.push(`${internalName(flow.statusField)} eq ${odataValue(col, flow.reminderValue)}`);
    }
    if (days) filters.push(`Created le '@{addDays(utcNow(), -${days}, 'yyyy-MM-ddTHH:mm:ssZ')}'`);
    actions.Get_items = {
      runAfter: {}, type: 'OpenApiConnection',
      inputs: { host: host(SP, 'GetItems'), parameters: spParams(project, flow, { ...(filters.length ? { $filter: filters.join(' and ') } : {}), $top: 500 }), authentication: "@parameters('$authentication')" },
    };
    actions.Apply_to_each = { foreach: "@outputs('Get_items')?['body/value']", actions: { 'Send_an_email_(V2)': sendEmail(project, flow, {}) }, runAfter: { Get_items: ['Succeeded'] }, type: 'Foreach' };
  }
  if (flow.kind === 'button') {
    triggers.manual = {
      type: 'Request', kind: 'PowerAppV2',
      inputs: { schema: { type: 'object', properties: { number: { title: 'ItemID', type: 'number', 'x-ms-dynamically-added': true, description: 'ID of the item to use', 'x-ms-content-hint': 'NUMBER' } }, required: ['number'] } },
    };
    actions.Get_item = {
      runAfter: {}, type: 'OpenApiConnection',
      inputs: { host: host(SP, 'GetItem'), parameters: spParams(project, flow, { id: "@triggerBody()['number']" }), authentication: "@parameters('$authentication')" },
    };
    actions['Send_an_email_(V2)'] = sendEmail(project, flow, { Get_item: ['Succeeded'] });
    actions.Respond_to_a_PowerApp_or_flow = {
      runAfter: { 'Send_an_email_(V2)': ['Succeeded'] }, type: 'Response', kind: 'PowerApp',
      inputs: { statusCode: 200, body: { result: 'Sent' }, schema: { type: 'object', properties: { result: { title: 'result', 'x-ms-dynamically-added': true, type: 'string' } } } },
    };
  }
  return {
    $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
    contentVersion: '1.0.0.0',
    parameters: { $connections: { defaultValue: {}, type: 'Object' }, $authentication: { defaultValue: {}, type: 'SecureObject' } },
    triggers, actions,
  };
}

// ------------------------------------------------------------------ package
const guid = () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => { const r = (Math.random() * 16) | 0; return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16); });
const key = () => String(Math.floor(Math.random() * 9e17) + 1e17).slice(0, 18) + String(Math.floor(Math.random() * 90 + 10));

/** All files of a legacy import package: {path: text}. sample = tenant details captured from a real export. */
export function flowPackageFiles(project, flow, sample = project.app.flowSample || null) {
  const flowKey = key(), flowGuid = guid();
  const conns = [SP, O365];
  const res = {}, apisMap = {}, connectionsMap = {}, connectionReferences = {};
  const depends = [];
  for (const c of conns) {
    const apiKey = key(), connKey = key();
    const s = sample?.[c] || {};
    res[apiKey] = { id: `/providers/Microsoft.PowerApps/apis/${c}`, name: c, type: 'Microsoft.PowerApps/apis', suggestedCreationType: 'Existing',
      details: { displayName: s.displayName || CONNECTORS[c].displayName, iconUri: s.iconUri || CONNECTORS[c].iconUri }, configurableBy: 'System', hierarchy: 'Child', dependsOn: [] };
    res[connKey] = { type: 'Microsoft.PowerApps/apis/connections', suggestedCreationType: 'Existing', creationType: 'Existing',
      details: { displayName: s.connectionDisplayName || CONNECTORS[c].displayName, iconUri: s.iconUri || CONNECTORS[c].iconUri }, configurableBy: 'User', hierarchy: 'Child', dependsOn: [apiKey] };
    apisMap[c] = apiKey; connectionsMap[c] = connKey; depends.push(apiKey, connKey);
    connectionReferences[c] = { connectionName: s.connectionName || `${c.replace('shared_', 'shared-').slice(0, 20)}-${guid()}`, source: 'Embedded', id: `/providers/Microsoft.PowerApps/apis/${c}`, tier: 'NotSpecified' };
  }
  const manifest = {
    schema: '1.0',
    details: { displayName: flow.name, description: `Generated by Power App Studio for ${project.app.name}.`, createdTime: new Date().toISOString(), packageTelemetryId: guid(), creator: 'Power App Studio', sourceEnvironment: '' },
    resources: {
      [flowKey]: { id: `/providers/Microsoft.Flow/flows/${flowGuid}`, name: flowGuid, type: 'Microsoft.Flow/flows', suggestedCreationType: 'New', creationType: 'Existing, New, Update',
        details: { displayName: flow.name }, configurableBy: 'User', hierarchy: 'Root', dependsOn: depends },
      ...res,
    },
  };
  const definition = {
    name: flowGuid, id: `/providers/Microsoft.Flow/flows/${flowGuid}`, type: 'Microsoft.Flow/flows',
    properties: {
      apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows', displayName: flow.name,
      definition: flowDefinition(project, flow), connectionReferences,
      flowFailureAlertSubscribed: false, isManaged: false,
    },
  };
  const dir = `Microsoft.Flow/flows/${flowKey}`;
  return {
    'manifest.json': JSON.stringify(manifest),
    'Microsoft.Flow/flows/manifest.json': JSON.stringify({ packageSchemaVersion: '1.0', flowAssets: { assetPaths: [flowKey] } }),
    [`${dir}/definition.json`]: JSON.stringify(definition),
    [`${dir}/apisMap.json`]: JSON.stringify(apisMap),
    [`${dir}/connectionsMap.json`]: JSON.stringify(connectionsMap),
  };
}
export const flowFileName = (flow) => `${slug(flow.name) || 'Flow'}.zip`;

/** Read tenant details (connector display names, icons, connection names) out of a real exported package. */
export function sampleFromPackage(files) {
  const man = JSON.parse(files['manifest.json'] || '{}');
  const out = {};
  const resources = man.resources || {};
  for (const [k, r] of Object.entries(resources)) {
    if (r.type !== 'Microsoft.PowerApps/apis' || !CONNECTORS[r.name]) continue;
    const conn = Object.values(resources).find((x) => x.type === 'Microsoft.PowerApps/apis/connections' && (x.dependsOn || []).includes(k));
    out[r.name] = { displayName: r.details?.displayName, iconUri: r.details?.iconUri, connectionDisplayName: conn?.details?.displayName };
  }
  for (const [p, t] of Object.entries(files)) {
    if (!/definition\.json$/.test(p)) continue;
    const refs = JSON.parse(t).properties?.connectionReferences || {};
    for (const [c, r] of Object.entries(refs)) if (CONNECTORS[c]) out[c] = { ...(out[c] || {}), connectionName: r.connectionName };
  }
  return out;
}

// ------------------------------------------------------------------ app side
/** Flows that fire for an app event, for Play mode. */
export function flowsFor(project, list, event, changed = {}) {
  return (project.flows || []).filter((f) => f.enabled !== false && f.list === list && (
    (event === 'created' && f.kind === 'created') ||
    (event === 'changed' && f.kind === 'statusChanged' && f.statusField in changed && (!(f.statusValues || []).length || f.statusValues.includes(String(changed[f.statusField]))))));
}
/** Deep link: the start screen formula and the screen that opens an item from ?ID= */
export function deepLinkInfo(project) {
  const d = project.app.deepLink;
  if (!d?.screen || !d?.list) return null;
  const scr = project.screens.find((s) => s.id === d.screen);
  if (!scr || !project.lists.some((l) => l.name === d.list)) return null;
  return { screen: scr, list: d.list };
}
export function flowLint(project, add) {
  const names = new Set();
  for (const f of project.flows || []) {
    const where = `flow ${f.name}`;
    if (names.has(flowFxName(f.name))) add('error', 'PS050', where, 'Two flows have the same name once spaces are removed; Power Apps would not tell them apart.');
    names.add(flowFxName(f.name));
    const list = project.lists.find((l) => l.name === f.list);
    if (!list) { add('error', 'PS043', where, 'Pick a list from the data model.'); continue; }
    if (f.kind !== 'button' && !String(project.app.siteUrl || '').trim()) add('error', 'PS045', where, 'Set the SharePoint site URL on the App tab; SharePoint triggers need it.');
    if (f.kind === 'button' && !String(project.app.siteUrl || '').trim()) add('error', 'PS045', where, 'Set the SharePoint site URL on the App tab; Get item needs it.');
    if (!recipientsExpr(project, f)) add('error', 'PS044', where, 'Add at least one recipient.');
    if (f.kind === 'statusChanged') {
      const col = list.columns.find((c) => c.name === f.statusField);
      if (!col) add('error', 'PS046', where, 'Pick the column to watch.');
      else if (col.type !== 'Choice') add('error', 'PS046', where, `${col.name} is not a Choice column; pick a Choice column to watch.`);
      add('info', 'PS046', where, 'Uses "Get changes for an item": keep version history on for this list (SharePoint default).');
    }
    if (f.kind === 'reminder' && f.statusField && !list.columns.some((c) => c.name === f.statusField)) add('error', 'PS046', where, 'The reminder column is not on the list.');
    const bad = [...badTokens(project, f, f.subject), ...badTokens(project, f, f.body)];
    if (bad.length) add('warn', 'PS048', where, `Unknown token${bad.length > 1 ? 's' : ''} ${[...new Set(bad)].map((b) => '{' + b + '}').join(', ')}; they will be sent as plain text.`);
    if (/\{Link\}/.test(f.body + f.subject) && !String(project.app.appLink || '').trim()) add('warn', 'PS047', where, '{Link} needs the app web link on the App tab (Apps → Details → Web link). Until then it sends only the item ID.');
    if (/\{Link\}/.test(f.body + f.subject) && project.app.appLink && !deepLinkInfo(project)) add('warn', 'PS047', where, 'Pick the screen that opens an item from an email link on the App tab, or the link will just open the start screen.');
    for (const r of f.recipients || []) if (r.kind === 'column' && !list.columns.some((c) => c.name === r.value && c.type === 'Person')) add('error', 'PS044', where, `Recipient column ${r.value || '(none)'} is not a Person column of ${list.name}.`);
  }
}
