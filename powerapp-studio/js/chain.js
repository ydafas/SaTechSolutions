// chain.js — multi-stage approval chains (1–4 ordered stages) with per-stage decisions and an append-only history.
//
// Data model, all on the item's own list (no second list, no lookups to patch):
//   <Stage>Decision   Choice  Pending / Approved / Rejected   (default Pending)  ← the stepper reads THIS per stage
//   <Stage>DecidedOn  Date
//   <Stage>DecidedBy  Person   (meUserFx claims shape, field-tested)
//   <Stage>Comments   Note
//   CurrentStage      Number   1..n while waiting, 0 when finished (default 1)   ← delegable queue filter
//   ApprovalHistory   Note     one line per decision, appended by the app
//   Status            Choice   overall: Submitted / In review / Approved / Rejected
// Approver per stage: a Person column on the item (the requester picks it on the form) or a Microsoft 365 group
// checked once in OnStart with Office365Groups.ListGroupMembers wrapped in IfError (same shape as admin gating).
// History is a Note column instead of a History list: one Patch writes the decision and its history line together,
// so there is no second write to fail halfway, no Lookup column to patch, and nothing extra to provision.
import { q, s, varSel, newColumn, IDENT, stageKey, normalizeChain } from './core.js';
import { ifErrorFx, notifyFx, meUserFx } from './fx.js';

export const DECISIONS = ['Pending', 'Approved', 'Rejected'];
export const MAX_STAGES = 4;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const isGuid = (v) => GUID.test(String(v || '').trim());

/** Column names for one stage. */
export function stageCols(stage) {
  const K = stageKey(stage);
  return { K, decision: K + 'Decision', date: K + 'DecidedOn', by: K + 'DecidedBy', comments: K + 'Comments', groupVar: 'varIsApprover' + K };
}
export { normalizeChain };
export const chainFor = (project, listName) => (project.app?.chain && project.app.chain.list === listName ? project.app.chain : null);

/**
 * Add the chain's columns to its list (idempotent). Approver Person columns stay visible (the requester fills them in);
 * every other chain column is flagged chain:true so forms, rows and detail grids leave it out.
 */
/** Pick the chain's four overall-status values from an existing Choice column, so a scenario's "Submitted / Approved / Rejected" are reused rather than duplicated. */
export function guessChainValues(list, field = 'Status') {
  const col = list?.columns.find((c) => c.name === field && c.type === 'Choice');
  const ch = col?.choices || [];
  const pick = (re, d) => ch.find((x) => re.test(x)) || d;
  return { pending: pick(/pend|submit|new|wait|open/i, 'Submitted'), review: pick(/review|progress|process/i, 'In review'), approved: pick(/approv|accept/i, 'Approved'), rejected: pick(/reject|deny|denied|declin/i, 'Rejected') };
}
export function ensureChainColumns(project, chain) {
  const list = project.lists.find((l) => l.name === chain.list);
  if (!list) return null;
  const add = (name, type, extra = {}) => {
    let c = list.columns.find((x) => x.name === name);
    if (!c) { c = newColumn(name, type, extra); list.columns.push(c); } else Object.assign(c, { ...extra, type: c.type });
    return c;
  };
  // overall status: keep the user's column and its values, add the four the chain needs
  const st = add(chain.field, 'Choice', { chainRole: 'status', chainValues: { ...chain.values } });
  if (st.type !== 'Choice') { st.type = 'Choice'; }
  const want = [chain.values.pending, chain.values.review, chain.values.approved, chain.values.rejected];
  st.choices = [...new Set([chain.values.pending, ...(st.choices || []).filter((x) => x !== chain.values.pending), ...want])];
  chain.stages.forEach((stage, i) => {
    const k = stageCols(stage), n = i + 1;
    if (stage.approver === 'column') {
      if (!stage.column) stage.column = k.K + 'Approver';
      const ap = list.columns.find((x) => x.name === stage.column);
      if (ap && ap.type !== 'Person') { stage.column = k.K + 'Approver'; }
      add(stage.column, 'Person', { required: true, chainRole: 'approver', chainStage: n });
    }
    add(k.decision, 'Choice', { choices: [...DECISIONS], chain: true, chainRole: 'decision', chainStage: n });
    add(k.date, 'DateTime', { chain: true, chainRole: 'date', chainStage: n });
    add(k.by, 'Person', { chain: true, chainRole: 'by', chainStage: n });
    add(k.comments, 'Note', { chain: true, chainRole: 'comments', chainStage: n });
  });
  add('CurrentStage', 'Number', { chain: true, chainRole: 'stage', defaultValue: 1 });
  add('ApprovalHistory', 'Note', { chain: true, chainRole: 'history' });
  return list;
}

// ------------------------------------------------------------------ formulas
const ref = (V, col) => `${V}.${q(col)}`;
const decisionOf = (V, stage) => `Coalesce(${ref(V, stageCols(stage).decision)}.Value, "Pending")`;
/** Who may decide stage n, as a test on the selected record V. */
function whoFx(stage, V) {
  return stage.approver === 'group' ? stageCols(stage).groupVar : `Lower(${ref(V, stage.column)}.Email) = Lower(User().Email)`;
}
/** True when the signed-in user may decide the selected item now (never their own request). */
export function canDecideFx(chain, V) {
  const parts = chain.stages.map((st, i) => `(${V}.CurrentStage = ${i + 1} && ${whoFx(st, V)})`);
  return `!IsBlank(${V}.ID) && Lower(${V}.'Created By'.Email) <> Lower(User().Email) && (${parts.join(' || ')})`;
}
/** Requester may resubmit a rejected item. */
export function canResubmitFx(chain, V) {
  return `!IsBlank(${V}.ID) && Lower(${V}.'Created By'.Email) = Lower(User().Email) && ${ref(V, chain.field)}.Value = ${s(chain.values.rejected)}`;
}
/**
 * Approval queue: items at a stage the user can decide. The inner Filter is delegable (Number compare, Person.Email =
 * User().Email, constant group flags); "not my own request" runs locally on its result (skill delegation shape).
 */
export function queueFx(chain) {
  const L = q(chain.list);
  const parts = chain.stages.map((st, i) => st.approver === 'group' ? `(CurrentStage = ${i + 1} && ${stageCols(st).groupVar})` : `(CurrentStage = ${i + 1} && ${q(st.column)}.Email = User().Email)`);
  return `SortByColumns(\n    Filter(\n        Filter(${L}, ${parts.join(' || ')}),\n        'Created By'.Email <> User().Email\n    ),\n    "Created",\n    SortOrder.Ascending\n)`;
}
/** OnStart lines: one membership flag per group stage. */
export function groupChecksFx(chain) {
  return chain.stages.filter((st) => st.approver === 'group' && st.groupId).map((st) =>
    `Set(${stageCols(st).groupVar}, IfError(!IsBlank(LookUp(Office365Groups.ListGroupMembers(${s(st.groupId)}).value, Lower(mail) = Lower(User().Email))), false))`);
}
const historyLine = (stageName, verb, commentCtl) =>
  `Text(Now(), "mm/dd/yyyy hh:mm") & "  " & User().FullName & "  ·  ${stageName.replace(/"/g, '""')}  ·  ${verb}" & If(IsBlank(Trim(${commentCtl}.Text)), "", "  ·  " & Trim(${commentCtl}.Text))`;
// newest line first, so a label showing the top of the history shows the latest decision
const appendHistory = (line) => `${line} & If(IsBlank(cur.ApprovalHistory), "", Char(10) & cur.ApprovalHistory)`;

/**
 * Approve or reject the selected item's current stage. One Patch per stage branch writes the stage's decision, date,
 * approver, comment, the next CurrentStage, the overall status and the history line together, on a fresh LookUp.
 */
export function decideFx(chain, decision, commentCtl) {
  const L = q(chain.list), V = varSel(chain.list), n = chain.stages.length, approve = decision === 'Approved';
  const branches = chain.stages.map((stage, i) => {
    const k = stageCols(stage), last = i === n - 1;
    const next = approve && !last ? i + 2 : 0;
    const overall = approve ? (last ? chain.values.approved : chain.values.review) : chain.values.rejected;
    const rec = [
      `${q(k.decision)}: {Value: ${s(decision)}}`,
      `${q(k.date)}: Today()`,
      `${q(k.by)}: ${meUserFx()}`,
      `${q(k.comments)}: Trim(${commentCtl}.Text)`,
      `CurrentStage: ${next}`,
      `${q(chain.field)}: {Value: ${s(overall)}}`,
      `ApprovalHistory: ${appendHistory(historyLine(stage.name, decision, commentCtl))}`,
    ];
    const msg = approve ? (last ? 'Approved' : `Approved. Sent to ${chain.stages[i + 1].name}.`) : 'Rejected';
    const ok = [`Set(${V}, LookUp(${L}, ID = ${V}.ID))`, notifyFx(msg, 'Success'), `Reset(${commentCtl})`];
    return `${i + 1}, ${ifErrorFx(`Patch(${L}, cur, {\n        ${rec.join(',\n        ')}\n    })`, approve ? 'Approve' : 'Reject', ok).replace(/\n/g, '\n        ')}`;
  });
  return `With({cur: LookUp(${L}, ID = ${V}.ID)},\n    Switch(cur.CurrentStage,\n        ${branches.join(',\n        ')}\n    )\n)`;
}
/** Requester resubmits a rejected item: every stage back to Pending, CurrentStage 1, status Submitted, history line. */
export function resubmitFx(chain) {
  const L = q(chain.list), V = varSel(chain.list);
  const rec = [];
  chain.stages.forEach((stage) => { const k = stageCols(stage); rec.push(`${q(k.decision)}: {Value: "Pending"}`, `${q(k.date)}: Blank()`, `${q(k.comments)}: ""`); });
  rec.push('CurrentStage: 1', `${q(chain.field)}: {Value: ${s(chain.values.pending)}}`,
    `ApprovalHistory: ${appendHistory(`Text(Now(), "mm/dd/yyyy hh:mm") & "  " & User().FullName & "  ·  Resubmitted"`)}`);
  return `With({cur: LookUp(${L}, ID = ${V}.ID)},\n    ${ifErrorFx(`Patch(${L}, cur, {\n        ${rec.join(',\n        ')}\n    })`, 'Resubmit', [`Set(${V}, LookUp(${L}, ID = ${V}.ID))`, notifyFx('Resubmitted for approval', 'Success')]).replace(/\n/g, '\n    ')}\n)`;
}
/** New items start at stage 1 (column default 1 covers lists made by the script; this covers hand-made lists). */
export const newItemStageFx = (chain) => `Coalesce(LookUp(${q(chain.list)}, ID = ${varSel(chain.list)}.ID).CurrentStage, 1)`;

/** Stepper HTML (template for templates.hx): one step per stage, coloured from THAT stage's own Decision field. */
export function stepperTpl(chain, V, h, esc) {
  const cells = chain.stages.map((stage, i) => {
    const k = stageCols(stage), n = i + 1, d = decisionOf(V, stage);
    const color = `Switch(${d}, "Approved", appTheme.SuccessHex, "Rejected", appTheme.DangerHex, If(${V}.CurrentStage = ${n}, appTheme.PrimaryHex, appTheme.BorderHex))`;
    const mark = `Switch(${d}, "Approved", "&#10003;", "Rejected", "&#10005;", "${n}")`;
    const who = stage.approver === 'group' ? s(stage.groupLabel || `${stage.name} group`) : `Coalesce(${ref(V, stage.column)}.DisplayName, "Not set")`;
    const sub = `If(${d} = "Pending", If(${V}.CurrentStage = ${n}, "Waiting on " & ${esc(who)}, "Not reached yet"), ${d} & " " & Text(${ref(V, k.date)}, "mm/dd") & " · " & ${esc(`${ref(V, k.by)}.DisplayName`)})`;
    return `<td style='text-align:center;vertical-align:top;padding:0 4px'>`
      + `<div style='margin:0 auto;width:30px;height:30px;border-radius:50%;line-height:30px;font-size:13px;font-weight:700;color:#FFFFFF;background:{{=${color}}}'>{{=${mark}}}</div>`
      + `<div style='font-size:12px;margin-top:6px;font-weight:{{=If(${V}.CurrentStage = ${n}, "700", "600")}};color:{{Text}}'>${stage.name.replace(/[<&'"]/g, '')}</div>`
      + `<div style='font-size:11px;margin-top:2px;color:{{TextMuted}};white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{{=${sub}}}</div></td>`;
  }).join('');
  return `<div style='margin:3px;box-sizing:border-box;height:${h - 16}px;overflow:hidden;background:{{Surface}};border:1px solid {{Border}};border-radius:{{Radius}}px;box-shadow:0 1px 3px rgba(15,23,42,0.08);font-family:{{Font}};padding:14px 10px 0'>`
    + `<div style='font-size:12px;font-weight:700;color:{{TextMuted}};padding:0 8px 10px'>Approval</div>`
    + `<table style='width:100%;border-collapse:collapse;table-layout:fixed'><tr>${cells}</tr></table></div>`;
}

// ------------------------------------------------------------------ flows
/** Flow specs for the chain: notify each stage's approver as the item reaches it, and the requester at the end. */
export function chainFlowSpecs(chain) {
  const out = [];
  const to = (stage) => stage.approver === 'column' ? [{ kind: 'column', value: stage.column }] : (stage.notify ? [{ kind: 'address', value: stage.notify }] : []);
  chain.stages.forEach((stage, i) => {
    const r = to(stage);
    if (!r.length) return;
    if (i === 0) out.push({ kind: 'created', name: `Notify ${stage.name} of new item`, recipients: r, subject: `Approval needed: {Title}`, body: `Hello,\n\n{Author} submitted "{Title}" and it needs your ${stage.name} approval.\n\nOpen it in the app: {Link}` });
    if (i === 0 && chain.resubmit !== false) out.push({ kind: 'statusChanged', name: `Notify ${stage.name} of resubmitted item`, recipients: r, statusField: chain.field, statusValues: [chain.values.pending], subject: `Approval needed again: {Title}`, body: `Hello,\n\n{Author} changed and resubmitted "{Title}". It needs your ${stage.name} approval again.\n\nOpen it in the app: {Link}` });
    else {
      const prev = stageCols(chain.stages[i - 1]);
      out.push({ kind: 'statusChanged', name: `Notify ${stage.name} after ${chain.stages[i - 1].name}`, recipients: r, statusField: prev.decision, statusValues: ['Approved'],
        subject: `Approval needed: {Title}`, body: `Hello,\n\n"{Title}" was approved by ${chain.stages[i - 1].name} and now needs your ${stage.name} approval.\n\nOpen it in the app: {Link}` });
    }
  });
  out.push({ kind: 'statusChanged', name: 'Tell requester the outcome', recipients: [{ kind: 'author', value: '' }], statusField: chain.field, statusValues: [chain.values.approved, chain.values.rejected],
    subject: `Your request is {${chain.field}}: {Title}`, body: `Hello {Author},\n\nYour request "{Title}" is now {${chain.field}}.\n\nOpen it in the app: {Link}` });
  return out;
}

// ------------------------------------------------------------------ Play
/** Apply a decision to a Play row (same fields the Patch writes). Returns {msg, changed} or {error}. */
export function playDecide(chain, row, decision, comment, user = { DisplayName: 'Alex Rivera', Mail: 'alex.rivera@contoso.com', Email: 'alex.rivera@contoso.com' }) {
  const n = chain.stages.length, cur = Number(row.CurrentStage) || 0;
  if (cur < 1 || cur > n) return { error: 'This item is not waiting for a decision.' };
  if (decision === 'Rejected' && !String(comment || '').trim()) return { error: 'Add a comment to reject.' };
  const stage = chain.stages[cur - 1], k = stageCols(stage), last = cur === n, approve = decision === 'Approved';
  const before = { ...row };
  row[k.decision] = decision; row[k.date] = new Date().toISOString().slice(0, 10); row[k.by] = user; row[k.comments] = String(comment || '').trim();
  row.CurrentStage = approve && !last ? cur + 1 : 0;
  row[chain.field] = approve ? (last ? chain.values.approved : chain.values.review) : chain.values.rejected;
  const stamp = new Date();
  const line = `${String(stamp.getMonth() + 1).padStart(2, '0')}/${String(stamp.getDate()).padStart(2, '0')}/${stamp.getFullYear()} ${String(stamp.getHours()).padStart(2, '0')}:${String(stamp.getMinutes()).padStart(2, '0')}  ${user.DisplayName}  ·  ${stage.name}  ·  ${decision}${row[k.comments] ? '  ·  ' + row[k.comments] : ''}`;
  row.ApprovalHistory = line + (row.ApprovalHistory ? '\n' + row.ApprovalHistory : '');
  const changed = {}; for (const key of Object.keys(row)) if (JSON.stringify(before[key]) !== JSON.stringify(row[key])) changed[key] = row[key];
  return { msg: approve ? (last ? 'Approved' : `Approved. Sent to ${chain.stages[cur].name}.`) : 'Rejected', changed };
}
export function playResubmit(chain, row, user = { DisplayName: 'Alex Rivera' }) {
  for (const st of chain.stages) { const k = stageCols(st); row[k.decision] = 'Pending'; row[k.date] = ''; row[k.comments] = ''; }
  row.CurrentStage = 1; row[chain.field] = chain.values.pending;
  row.ApprovalHistory = `${new Date().toLocaleDateString('en-US')}  ${user.DisplayName}  ·  Resubmitted` + (row.ApprovalHistory ? '\n' + row.ApprovalHistory : '');
}
/** Can this (Play) user decide this row now? */
export function playCanDecide(chain, row, vars = {}, me = 'alex.rivera@contoso.com') {
  const cur = Number(row.CurrentStage) || 0;
  if (cur < 1 || cur > chain.stages.length) return false;
  if (String(row.Author?.Mail || row.Author?.Email || '').toLowerCase() === me) return false;
  const st = chain.stages[cur - 1];
  return st.approver === 'group' ? vars[stageCols(st).groupVar] !== false : String(row[st.column]?.Mail || row[st.column]?.Email || '').toLowerCase() === me;
}

// ------------------------------------------------------------------ lint
export function chainLint(project, add) {
  const raw = project.app?.chain;
  if (!raw) return;
  const c = raw, where = 'Approval chain';
  const list = project.lists.find((l) => l.name === c.list);
  if (!list) { add('error', 'PS053', where, `The approval chain's list "${c.list}" is not in the data model.`); return; }
  if (!c.stages?.length || c.stages.length > MAX_STAGES) add('error', 'PS053', where, `An approval chain needs 1 to ${MAX_STAGES} stages.`);
  for (const [i, st] of (c.stages || []).entries()) {
    const k = stageCols(st), w = `${where} · stage ${i + 1} (${st.name})`;
    if (st.approver === 'group') {
      if (!isGuid(st.groupId)) add('error', 'PS054', w, 'Approver group needs a Microsoft 365 group ID (a GUID). Finish setup → Approval groups.');
      if (!st.notify) add('info', 'PS054', w, 'Group stages get no "approval needed" email unless you give the group a mailbox address.');
    } else {
      const col = list.columns.find((x) => x.name === st.column);
      if (!col || col.type !== 'Person') add('error', 'PS055', w, `Approver column ${st.column || '(none)'} must be a Person column of ${list.name}.`);
    }
    for (const cn of [k.decision, k.date, k.by, k.comments]) if (!list.columns.some((x) => x.name === cn)) add('error', 'PS056', w, `Column ${list.name}.${cn} is missing; the chain writes it. Turn the chain off and on again to recreate it.`);
  }
  for (const cn of ['CurrentStage', 'ApprovalHistory', c.field]) if (!list.columns.some((x) => x.name === cn)) add('error', 'PS056', where, `Column ${list.name}.${cn} is missing; the chain writes it.`);
  void IDENT;
}
