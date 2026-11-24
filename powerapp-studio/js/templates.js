// templates.js — requirement scenarios, whole-screen templates and insertable blocks.
import { newControl, newScreen, newList, newColumn, FORM_FACTORS, allNames, slug, varSel, uid } from './core.js';
import { readFx, missingFx, usageFx } from './emit.js';
import { newFlow } from './flows.js';
import { placeholderHtml } from './attach.js';
import { guessChainValues, chainFor, canDecideFx, canResubmitFx, queueFx, stepperTpl, ensureChainColumns, normalizeChain, stageCols, chainSla, overdueFx, canReassignFx } from './chain.js';


// ---------------------------------------------------------------- scenarios (the requirements wizard)
const C = newColumn;
export const SCENARIOS = {
  requests: {
    label: 'Request & approval tracker', icon: 'send', blurb: 'People submit requests; approvers approve or reject; everyone tracks status.',
    app: 'Request Tracker',
    lists: () => [newList('Requests', [
      C('Details', 'Note'), C('Category', 'Choice', { choices: ['Hardware', 'Software', 'Access', 'Other'], required: true }),
      C('Priority', 'Choice', { choices: ['Low', 'Medium', 'High'] }),
      C('Status', 'Choice', { choices: ['Submitted', 'Approved', 'Rejected', 'Completed'] }),
      C('NeededBy', 'DateTime'), C('Approver', 'Person'), C('EstimatedCost', 'Currency')])],
    approvals: { list: 'Requests', field: 'Status', pending: 'Submitted', yes: 'Approved', no: 'Rejected' },
  },
  assets: {
    label: 'Asset inventory', icon: 'box', blurb: 'Track equipment, who has it, where it is and what condition it is in.',
    app: 'Asset Inventory',
    lists: () => [newList('Assets', [
      C('AssetTag', 'Text', { required: true }), C('Category', 'Choice', { choices: ['Laptop', 'Monitor', 'Phone', 'Furniture', 'Other'] }),
      C('AssignedTo', 'Person'), C('AssetLocation', 'Text'), C('Condition', 'Choice', { choices: ['New', 'Good', 'Fair', 'Retired'] }),
      C('PurchaseDate', 'DateTime'), C('PurchasePrice', 'Currency'), C('InService', 'YesNo')])],
  },
  events: {
    label: 'Event registration', icon: 'calendar', blurb: 'Admins post events; people register, bring guests and cancel.',
    app: 'Event Registration',
    lists: () => [
      newList('Events', [C('EventDate', 'DateTime', { required: true }), C('Venue', 'Text'), C('Capacity', 'Number'), C('Details', 'Note')]),
      newList('Registrations', [C('Event', 'Lookup', { lookupList: 'Events', required: true }), C('Attendee', 'Person'), C('Guests', 'Number'),
        C('RegStatus', 'Choice', { choices: ['Registered', 'Waitlist', 'Cancelled'] })]),
    ],
  },
  tasks: {
    label: 'Team task tracker', icon: 'check', blurb: 'Assign tasks, set due dates and priorities, move them to done.',
    app: 'Task Tracker',
    lists: () => [newList('Tasks', [
      C('Details', 'Note'), C('AssignedTo', 'Person'), C('DueDate', 'DateTime'),
      C('Priority', 'Choice', { choices: ['Low', 'Medium', 'High', 'Urgent'] }),
      C('Status', 'Choice', { choices: ['Not started', 'In progress', 'Blocked', 'Done'] })])],
  },
  contacts: {
    label: 'Contact directory', icon: 'users', blurb: 'A searchable list of people and companies with notes.',
    app: 'Contact Directory',
    lists: () => [newList('Contacts', [C('Company', 'Text'), C('Email', 'Text'), C('Phone', 'Text'),
      C('ContactType', 'Choice', { choices: ['Customer', 'Vendor', 'Partner', 'Internal'] }), C('Notes', 'Note')])],
  },
  workorders: {
    label: 'Work orders (IT / Facilities)', icon: 'settings', cat: 'Operations', blurb: 'Report a problem, assign a technician, track priority and status to completion.',
    app: 'Work Orders',
    lists: () => [newList('WorkOrders', [
      C('Details', 'Note', { required: true }), C('Category', 'Choice', { choices: ['Facilities', 'IT', 'Electrical', 'Plumbing', 'Other'], required: true }),
      C('Priority', 'Choice', { choices: ['Normal', 'Low', 'High', 'Critical'] }),
      C('Status', 'Choice', { choices: ['New', 'Assigned', 'In progress', 'On hold', 'Completed', 'Cancelled'] }),
      C('Technician', 'Person'), C('Building', 'Text'), C('DueDate', 'DateTime'), C('Resolution', 'Note')])],
    features: { loading: true, confirm: true, stepper: true, scope: 'all' },
    flows: [
      { kind: 'statusChanged', name: 'Tell the technician', list: 'WorkOrders', recipients: [{ kind: 'column', value: 'Technician' }], statusField: 'Status', statusValues: ['Assigned'], subject: 'Work order assigned: {Title}', body: 'Hello,\n\nWork order "{Title}" ({Priority}, {Building}) is assigned to you.\n\n{Details}\n\nOpen it in the app: {Link}' },
      { kind: 'statusChanged', name: 'Tell the requester it is done', list: 'WorkOrders', recipients: [{ kind: 'author', value: '' }], statusField: 'Status', statusValues: ['Completed'], subject: 'Work order completed: {Title}', body: 'Hello {Author},\n\nYour work order "{Title}" is completed.\n\n{Resolution}\n\nOpen it in the app: {Link}' },
      { kind: 'reminder', name: 'Daily reminder for open work orders', list: 'WorkOrders', recipients: [{ kind: 'column', value: 'Technician' }], statusField: 'Status', reminderValue: 'In progress', days: 3, hour: 7, subject: 'Still open: {Title}', body: 'Work order "{Title}" has been in progress since {Created}.\n\nOpen it in the app: {Link}' },
    ],
  },
  checkout: {
    label: 'Equipment check-out', icon: 'box', cat: 'Operations', blurb: 'Lend laptops, radios or tools: who has what, when it is due back, and what came back.',
    app: 'Equipment Check-out',
    lists: () => [
      newList('Equipment', [C('AssetTag', 'Text', { required: true }), C('Category', 'Choice', { choices: ['Laptop', 'Projector', 'Radio', 'Camera', 'Tool', 'Other'] }),
        C('Availability', 'Choice', { choices: ['Available', 'Checked out', 'In repair', 'Retired'] }), C('Notes', 'Note')]),
      newList('Checkouts', [C('Item', 'Lookup', { lookupList: 'Equipment', required: true }), C('Borrower', 'Person', { required: true }),
        C('DueBack', 'DateTime', { required: true }), C('ReturnedOn', 'DateTime'),
        C('LoanStatus', 'Choice', { choices: ['Checked out', 'Returned', 'Lost'] }), C('ReturnNotes', 'Note')]),
    ],
    features: { loading: true, confirm: true, stepper: false, scope: 'all' },
    flows: [
      { kind: 'created', name: 'Confirm the loan to the borrower', list: 'Checkouts', recipients: [{ kind: 'column', value: 'Borrower' }], subject: 'You checked out {Item}', body: 'Hello {Borrower},\n\nYou checked out {Item}. Please return it by {DueBack}.\n\nOpen the loan in the app: {Link}' },
      { kind: 'reminder', name: 'Weekly return reminder', list: 'Checkouts', recipients: [{ kind: 'column', value: 'Borrower' }], statusField: 'LoanStatus', reminderValue: 'Checked out', days: 7, hour: 8, subject: 'Reminder: {Item} is due back {DueBack}', body: 'Hello {Borrower},\n\n{Item} is still checked out to you (due back {DueBack}).\n\nOpen the loan in the app: {Link}' },
    ],
  },
  incidents: {
    label: 'Incident & safety reporting', icon: 'alert', cat: 'Compliance', blurb: 'Anyone reports an incident or near miss; an investigator records the cause and the corrective action.',
    app: 'Incident Reporting',
    lists: () => [newList('Incidents', [
      C('Details', 'Note', { required: true }), C('IncidentDate', 'DateTime', { required: true }),
      C('Severity', 'Choice', { choices: ['Low', 'Medium', 'High', 'Critical'], required: true }),
      C('IncidentType', 'Choice', { choices: ['Near miss', 'Injury', 'Property damage', 'Environmental', 'Security'] }),
      C('Site', 'Text'), C('Investigator', 'Person'),
      C('Status', 'Choice', { choices: ['Reported', 'Under investigation', 'Corrective action', 'Closed'] }),
      C('RootCause', 'Choice', { choices: ['Not yet known', 'Equipment', 'Procedure', 'Training', 'Environment', 'Human factors'] }), C('CorrectiveAction', 'Note')])],
    features: { loading: true, confirm: true, stepper: true, scope: 'mine' },
    flows: [
      { kind: 'statusChanged', name: 'Tell the investigator', list: 'Incidents', recipients: [{ kind: 'column', value: 'Investigator' }], statusField: 'Status', statusValues: ['Under investigation'], subject: 'Incident to investigate: {Title} ({Severity})', body: 'Hello,\n\nYou are the investigator for "{Title}" ({IncidentType}, {Severity}) at {Site} on {IncidentDate}.\n\n{Details}\n\nOpen it in the app: {Link}' },
      { kind: 'statusChanged', name: 'Tell the reporter it is closed', list: 'Incidents', recipients: [{ kind: 'author', value: '' }], statusField: 'Status', statusValues: ['Closed'], subject: 'Incident closed: {Title}', body: 'Hello {Author},\n\nThe incident you reported, "{Title}", is closed.\n\nCorrective action: {CorrectiveAction}\n\nOpen it in the app: {Link}' },
    ],
  },
  onboarding: {
    label: 'Employee onboarding', icon: 'users', cat: 'People', blurb: 'One record per new hire and a checklist of tasks for IT, facilities, HR and the manager.',
    app: 'Onboarding',
    lists: () => [
      newList('NewHires', [C('StartOn', 'DateTime', { required: true }), C('Department', 'Choice', { choices: ['Operations', 'Finance', 'IT', 'HR', 'Engineering', 'Other'] }),
        C('HiringManager', 'Person', { required: true }), C('Buddy', 'Person'),
        C('HireStatus', 'Choice', { choices: ['Pre-start', 'First week', 'First 90 days', 'Complete'] })]),
      newList('OnboardingTasks', [C('Hire', 'Lookup', { lookupList: 'NewHires', required: true }), C('TaskOwner', 'Person', { required: true }),
        C('DueDate', 'DateTime'), C('TaskArea', 'Choice', { choices: ['IT', 'Facilities', 'HR', 'Training', 'Manager'] }),
        C('TaskStatus', 'Choice', { choices: ['Not started', 'In progress', 'Done'] }), C('Notes', 'Note')]),
    ],
    features: { loading: true, confirm: true, stepper: true, scope: 'all' },
    flows: [
      { kind: 'created', name: 'Tell the task owner', list: 'OnboardingTasks', recipients: [{ kind: 'column', value: 'TaskOwner' }], subject: 'Onboarding task for {Hire}: {Title}', body: 'Hello {TaskOwner},\n\nYou have an onboarding task for {Hire}: {Title} (due {DueDate}).\n\nOpen it in the app: {Link}' },
      { kind: 'reminder', name: 'Daily reminder for tasks not started', list: 'OnboardingTasks', recipients: [{ kind: 'column', value: 'TaskOwner' }], statusField: 'TaskStatus', reminderValue: 'Not started', days: 2, hour: 8, subject: 'Not started yet: {Title} for {Hire}', body: 'Hello {TaskOwner},\n\n"{Title}" for {Hire} has not been started (added {Created}).\n\nOpen it in the app: {Link}' },
      { kind: 'created', name: 'Welcome note to the hiring manager', list: 'NewHires', recipients: [{ kind: 'column', value: 'HiringManager' }], subject: 'New hire added: {Title}', body: 'Hello {HiringManager},\n\n{Title} starts on {StartOn}. Their onboarding tasks will appear in the app.\n\n{Link}' },
    ],
  },
  leave: {
    label: 'Leave requests (2-step approval)', icon: 'calendar', cat: 'People', blurb: 'Employees request leave; their supervisor, then HR, approve or reject, with a full decision history.',
    app: 'Leave Requests',
    lists: () => [newList('LeaveRequests', [
      C('LeaveType', 'Choice', { choices: ['Annual', 'Sick', 'Personal', 'Training', 'Unpaid'], required: true }),
      C('FromDate', 'DateTime', { required: true }), C('ToDate', 'DateTime', { required: true }), C('Hours', 'Number', { required: true }),
      C('Reason', 'Note'), C('Supervisor', 'Person', { required: true }),
      C('Status', 'Choice', { choices: ['Submitted', 'In review', 'Approved', 'Rejected'] })])],
    approvals: { list: 'LeaveRequests', field: 'Status' },
    chain: { stages: [{ name: 'Supervisor', approver: 'column', column: 'Supervisor', groupId: '', notify: '' }, { name: 'HR', approver: 'group', column: '', groupId: '', notify: '' }] },
    features: { loading: true, confirm: true, stepper: true, scope: 'mine' },
  },
  policies: {
    label: 'Document & policy review', icon: 'file', cat: 'Compliance', blurb: 'Owners send policies for review; reviewers approve or request changes before the review due date.',
    app: 'Policy Review',
    lists: () => { const l = newList('Policies', [
      C('PolicyNumber', 'Text', { required: true }), C('PolicyOwner', 'Person', { required: true }), C('Reviewer', 'Person', { required: true }),
      C('ReviewDue', 'DateTime', { required: true }), C('PolicyVersion', 'Text'), C('Summary', 'Note'),
      C('ReviewStatus', 'Choice', { choices: ['Draft', 'In review', 'Approved', 'Changes requested', 'Published'] })]); l.attachments = true; return [l]; },
    approvals: { list: 'Policies', field: 'ReviewStatus', pending: 'In review', yes: 'Approved', no: 'Changes requested' },
    features: { loading: true, confirm: true, stepper: true, scope: 'all' },
    flows: [
      { kind: 'statusChanged', name: 'Ask the reviewer', list: 'Policies', recipients: [{ kind: 'column', value: 'Reviewer' }], statusField: 'ReviewStatus', statusValues: ['In review'], subject: 'Please review {PolicyNumber}: {Title}', body: 'Hello {Reviewer},\n\n{PolicyOwner} sent {PolicyNumber} "{Title}" (version {PolicyVersion}) for review. Please review it by {ReviewDue}.\n\nOpen it in the app: {Link}' },
      { kind: 'statusChanged', name: 'Tell the owner the outcome', list: 'Policies', recipients: [{ kind: 'column', value: 'PolicyOwner' }], statusField: 'ReviewStatus', statusValues: ['Approved', 'Changes requested'], subject: '{PolicyNumber} review: {ReviewStatus}', body: 'Hello {PolicyOwner},\n\nThe review of {PolicyNumber} "{Title}" is {ReviewStatus}.\n\nOpen it in the app: {Link}' },
      { kind: 'reminder', name: 'Daily reminder for reviews waiting', list: 'Policies', recipients: [{ kind: 'column', value: 'Reviewer' }], statusField: 'ReviewStatus', reminderValue: 'In review', days: 5, hour: 8, subject: 'Review waiting: {PolicyNumber}', body: 'Hello {Reviewer},\n\n{PolicyNumber} "{Title}" is still waiting for your review (due {ReviewDue}).\n\nOpen it in the app: {Link}' },
    ],
  },
  blank: { label: 'Start from my own lists', icon: 'grid', blurb: 'Design your lists first, then generate screens for them.', app: 'My App', lists: () => [newList('Items', [C('Status', 'Choice', { choices: ['Open', 'Closed'] })])] },
};

// ---------------------------------------------------------------- design system
// One spacing and type scale for every generated screen, so apps look designed rather than assembled.
// Building blocks are the field-tested ones from the powerapps-canvas-yaml skill:
//  * HtmlViewer cards and gallery rows (box-sizing:border-box + explicit height; content + margins <= Height - 10)
//  * Image icons between a card and its transparent Classic/Button tap target, which goes last
//  * two-click confirm (§27), per-field red borders + "still needed" summary (§5), confirmation overlay (§4)
//  * self-starting loading screen (timer + SMIL SVG in an Image, original emblem, separate progress ring)
//  * role check via Office365Groups.ListGroupMembers(...).value wrapped in IfError (§6)

/** Layout helpers for a screen: size, padding and an add() that keeps names unique app-wide. */
export function layoutCtx(project, screen) { return ctx(project, screen); }
function ctx(project, screen) {
  const ff = FORM_FACTORS[project.app.formFactor || 'tablet'];
  const used = allNames(project);
  const nm = (base) => { let n = base, i = 2; while (used.has(n)) n = base + i++; used.add(n); return n; };
  const add = (type, over, into) => { const c = newControl(project, type, { ...over, name: nm(over.name || type) }); (into || screen.controls).push(c); return c; };
  const phone = project.app.formFactor === 'phone';
  return { W: ff.w, H: ff.h, phone, nm, add, pad: phone ? 16 : 32, gap: phone ? 12 : 16, HH: phone ? 64 : 72 };
}
const features = (project) => ({ loading: false, usage: false, confirm: true, stepper: true, scope: 'all', adminGroup: '', ...(project.app.features || {}) });
export const statusCol = (list) => list.columns.find((c) => c.type === 'Choice' && /status|state|stage/i.test(c.name)) || list.columns.find((c) => c.type === 'Choice');
const subtitleCol = (list, skip) => list.columns.find((c) => !c.builtin && !c.chain && c !== skip && ['Person', 'Text', 'Lookup'].includes(c.type));
const dateCol = (list) => list.columns.find((c) => c.type === 'DateTime' && !c.chain);
const tag = (screen) => screen.name.replace(/^scr/, '');
const pretty = (n) => String(n).replace(/([a-z])([A-Z])/g, '$1 $2');
const singular = (n) => pretty(n).replace(/ies$/, 'y').replace(/s$/, '');
const lower = (n) => singular(n).toLowerCase();
const NEG = /reject|cancel|declin|denied|block|retired|lost|closed - no/i;
const tone = (c, i) => (/approv|done|complete|closed|registered|good|new|resolved|active|accepted/i.test(c) ? 'Success' : NEG.test(c) || /urgent/i.test(c) ? 'Danger' : /progress|pending|submit|wait|medium|fair|review|open|draft|not started|high/i.test(c) ? 'Warning' : i === 0 ? 'Primary' : 'TextMuted');
const toneIcon = (t) => ({ Success: 'check', Danger: 'close', Warning: 'clock', Primary: 'flag', TextMuted: 'list' }[t] || 'list');
const badgeMap = (st) => (st?.choices || []).map((c, i) => `${c}=@${tone(c, i)}`).join('\n');

// ---------------------------------------------------------------- HTML formula helpers
/** HTML template → Power Fx text expression. {{Token}} → appTheme.TokenHex, {{Radius}}, {{Font}}, {{=expr}} → (expr). */
export function hx(tpl) {
  const parts = String(tpl).split(/\{\{(=[^}]*|[A-Za-z]+)\}\}/);
  const out = [];
  parts.forEach((p, i) => {
    if (i % 2 === 0) { if (p) out.push('"' + p.replace(/"/g, '""') + '"'); return; }
    if (p[0] === '=') out.push('(' + p.slice(1) + ')');
    else if (p === 'Radius') out.push('appTheme.Radius');
    else if (p === 'Font') out.push('appTheme.FontCss');
    else out.push(`appTheme.${p}Hex`);
  });
  return out.join(' & ') || '""';
}
/** Escape text for HtmlText (user data can contain < and &). */
const esc = (e) => `Substitute(Substitute(Substitute(${e}, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")`;
/** Status colour hex as a Switch over the status text. */
const statusHex = (ref, st) => `Switch(${ref}, ${(st.choices || []).map((c, i) => `"${c.replace(/"/g, '""')}", appTheme.${tone(c, i)}Hex`).join(', ')}, appTheme.TextMutedHex)`;
const pill = (ref, st) => `<span style='display:inline-block;padding:3px 11px;border-radius:999px;font-size:11px;font-weight:600;color:#FFFFFF;background:{{=${statusHex(ref, st)}}}'>{{=${esc(ref)}}}</span>`;
/** Card shell with a 3px margin (room for the shadow): content height = control height - 16. */
const cardOpen = (h, extra = '') => `<div style='margin:3px;box-sizing:border-box;height:${h - 16}px;overflow:hidden;background:{{Surface}};border:1px solid {{Border}};border-radius:{{Radius}}px;box-shadow:0 1px 3px rgba(15,23,42,0.08);font-family:{{Font}};${extra}'>`;

// ---------------------------------------------------------------- shared blocks
export function headerBlock(project, screen, { title, titleFx, back = false, right = [], avatar = false, button = null } = {}) {
  const { W, add, phone, HH } = ctx(project, screen), T = tag(screen), t = project.theme;
  const out = [];
  out.push(t.header === 'gradient'
    ? add('card', { name: 'cardHeader' + T, x: 0, y: 0, w: W, h: HH, props: { fill: '@HeaderBg', fill2: '@Primary', gradAngle: 100, borderThickness: 0, radius: 0, shadow: false } })
    : add('rect', { name: 'rectHeader' + T, x: 0, y: 0, w: W, h: HH, props: { fill: '@HeaderBg', borderThickness: 0 } }));
  let x = phone ? 16 : 24;
  if (back) { out.push(add('icon', { name: 'icoBack' + T, x: phone ? 12 : 20, y: (HH - 32) / 2, w: 32, h: 32, props: { icon: 'back', color: '@OnHeader', stroke: 2.2 }, actions: { OnSelect: [{ type: 'back' }] } })); x = phone ? 52 : 64; }
  if (button) right = [...right, { text: button.text, actions: button.actions }];
  let rx = W - (phone ? 12 : 24);
  const btns = [];
  if (avatar) {
    rx -= 40;
    out.push(add('html', { name: 'htmlAvatar' + T, x: rx, y: (HH - 40) / 2, w: 40, h: 40, props: { html: '' },
      fx: { HtmlText: hx(`<div style='box-sizing:border-box;height:30px;width:30px;border-radius:50%;background:rgba(255,255,255,0.18);color:{{OnHeader}};font:600 14px {{Font}};text-align:center;line-height:30px'>{{=Upper(Left(User().FullName, 1))}}</div>`) } }));
    rx -= 8;
  }
  for (const r of [...right].reverse()) {
    const w = r.w || (phone ? 96 : 150);
    rx -= w;
    btns.push(add('button', { name: r.name || 'btnHeader' + T, x: rx, y: (HH - 40) / 2, w, h: 40, props: { text: r.text, fill: r.style === 'ghost' ? 'transparent' : '@Surface', color: r.style === 'ghost' ? '@OnHeader' : '@Primary', borderColor: r.style === 'ghost' ? '@OnHeader' : '@Surface', borderThickness: r.style === 'ghost' ? 1 : 0, size: 12, weight: 'Semibold' }, actions: { OnSelect: r.actions }, fx: r.fx || {} }));
    rx -= 10;
  }
  const lbl = add('label', { name: 'lblTitle' + T, x, y: (HH - 44) / 2, w: Math.max(120, rx - x - 8), h: 44, props: { text: title || pretty(screen.name.replace(/^scr/, '')), size: phone ? 16 : 18, weight: 'Semibold', color: '@OnHeader' } });
  if (titleFx) lbl.fx.Text = titleFx;
  out.push(lbl, ...btns);
  return out;
}
/** Icon in a rounded coloured chip (card + icon). */
function iconChip(add, name, x, y, size, icon, fill, fill2 = '') {
  add('card', { name: 'cardChip' + name, x, y, w: size, h: size, props: { fill, fill2, gradAngle: 135, borderThickness: 0, radius: Math.round(size / 3.2), shadow: false } });
  const i = Math.round(size * 0.5);
  add('icon', { name: 'icoChip' + name, x: x + (size - i) / 2, y: y + (size - i) / 2, w: i, h: i, props: { icon, color: '#FFFFFF', stroke: 2.2 } });
}
/** Centered empty state: icon chip, title, hint and an optional button. Visible when the gallery is empty. */
function emptyState(project, screen, gal, { x, y, w, title, hint, button }) {
  const { add } = ctx(project, screen), T = tag(screen) + gal.name.replace(/^gal/, '');
  const vis = { Visible: `CountRows(${gal.name}.AllItems) = 0` };
  const cx = x + w / 2;
  add('card', { name: 'cardEmptyChip' + T, x: cx - 28, y, w: 56, h: 56, props: { fill: '@SurfaceAlt', borderColor: '@Border', borderThickness: 1, radius: 28, shadow: false }, fx: { ...vis } });
  add('icon', { name: 'icoEmpty' + T, x: cx - 14, y: y + 14, w: 28, h: 28, props: { icon: 'search', color: '@TextMuted', stroke: 2 }, fx: { ...vis } });
  add('label', { name: 'lblEmpty' + T, x, y: y + 68, w, h: 28, props: { text: title, size: 13, weight: 'Semibold', color: '@Text', align: 'Center' }, fx: { ...vis } });
  if (hint) add('label', { name: 'lblEmptyHint' + T, x, y: y + 96, w, h: 24, props: { text: hint, size: 11, color: '@TextMuted', align: 'Center' }, fx: { ...vis } });
  if (button) add('button', { name: 'btnEmpty' + T, x: cx - 90, y: y + 132, w: 180, h: 40, props: { text: button.text, size: 12 }, actions: { OnSelect: button.actions }, fx: { ...vis } });
}

/** A gallery row: HtmlViewer card (title, meta line, status pill) + chevron icon. The row tap button is added by the emitter, last. */
export function galleryRows(project, scr, gal, list, { compact = false } = {}) {
  const { add } = ctx(project, scr), T = tag(scr) + gal.name.replace(/^gal/, '');
  const st = statusCol(list), sub = subtitleCol(list, st), dc = dateCol(list);
  const rowH = gal.props.templateSize = compact ? 72 : 84;
  gal.props.templateFill = 'transparent'; gal.props.selectedFill = 'transparent'; gal.props.emptyText = '';
  const title = readFx(project, list.name, list.columns[0].name, 'ThisItem');
  const metaParts = [];
  if (sub) metaParts.push(readFx(project, list.name, sub.name, 'ThisItem'));
  for (const mc of gal.data?.metaCols || []) {
    const c = list.columns.find((x) => x.name === mc);
    if (c && c !== sub && c !== dc) metaParts.push(`If(IsBlank(${readFx(project, list.name, c.name, 'ThisItem')}), "", "${pretty(c.name)} " & ${readFx(project, list.name, c.name, 'ThisItem')})`);
  }
  if (dc) metaParts.push(`If(IsBlank(ThisItem.${/^[A-Za-z_]\w*$/.test(dc.name) ? dc.name : `'${dc.name}'`}), "", "${pretty(dc.name)} " & ${readFx(project, list.name, dc.name, 'ThisItem')})`);
  else metaParts.push(`"Updated " & Text(ThisItem.Modified, "mm/dd/yyyy")`);
  const meta = metaParts.map((m) => esc(m)).join(` & "  ·  " & `);
  const W = gal.w;
  const pillCell = st ? `<td style='width:150px;text-align:right;vertical-align:middle;padding-right:30px'>${pill(`ThisItem.${/^[A-Za-z_]\w*$/.test(st.name) ? st.name : `'${st.name}'`}.Value`, st)}</td>` : `<td style='width:30px'></td>`;
  add('html', { name: 'htmlRow' + T, x: 0, y: 0, w: W, h: rowH, props: { html: '' },
    fx: { HtmlText: hx(`<div style='margin:4px 2px;box-sizing:border-box;height:${rowH - 18}px;padding:0 16px;background:{{Surface}};border:1px solid {{Border}};border-radius:{{Radius}}px;box-shadow:0 1px 3px rgba(15,23,42,0.08);font-family:{{Font}}'>`
      + `<table style='width:100%;height:${rowH - 20}px;border-collapse:collapse'><tr><td style='vertical-align:middle'>`
      + `<div style='font-size:${compact ? 13 : 14}px;font-weight:600;color:{{Text}};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:${Math.max(120, W - 240)}px'>{{=${esc(title)}}}</div>`
      + `<div style='font-size:12px;color:{{TextMuted}};margin-top:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:${Math.max(120, W - 240)}px'>{{=${meta}}}</div>`
      + `</td>${pillCell}</tr></table></div>`) } }, gal.children);
  add('icon', { name: 'icoRowGo' + T, x: W - 38, y: Math.round(rowH / 2 - 11), w: 22, h: 22, props: { icon: 'chevron', color: '@TextMuted', stroke: 2 } }, gal.children);
}

/** Label + input for one column. */
export function inputFor(project, screen, list, col, x, y, w) {
  const { add } = ctx(project, screen), T = tag(screen);
  const type = col.type === 'YesNo' ? 'toggle' : col.type === 'Person' ? 'person' : col.type === 'DateTime' ? 'datepicker' : (col.type === 'Choice' || col.type === 'Lookup') ? 'dropdown' : 'textinput';
  const prefix = { textinput: 'txt', dropdown: 'dd', datepicker: 'dtp', toggle: 'tgl', person: 'pp' }[type];
  const h = col.type === 'Note' ? 88 : type === 'toggle' ? 32 : 40;
  const lab = add('label', { name: 'lbl' + slug(col.name) + T, x, y, w, h: 22, props: { text: pretty(col.name) + (col.required ? ' *' : ''), size: 11, weight: 'Semibold', color: '@TextMuted' } });
  const inp = add(type, { name: prefix + slug(col.name) + T, x, y: y + 24, w: type === 'toggle' ? 80 : w, h, field: { list: list.name, column: col.name },
    props: type === 'textinput' ? { hint: col.type === 'Currency' ? '0.00' : col.type === 'Number' ? '0' : 'Enter ' + pretty(col.name).toLowerCase(), defaultText: '', multiline: col.type === 'Note', numeric: col.type === 'Number' || col.type === 'Currency', size: 12 } : undefined });
  return { lab, inp, bottom: y + 24 + h };
}
/** Section a column belongs to on generated forms and details screens. */
export const sectionOf = (col) => (col.type === 'Note' ? 'Notes' : ['Person', 'DateTime'].includes(col.type) ? 'People and dates' : 'Details');
/** Split columns into Details / People and dates / Notes. */
function sections(cols) {
  const groups = [
    { title: 'Details', cols: cols.filter((c) => !['Person', 'DateTime', 'Note'].includes(c.type)) },
    { title: 'People and dates', cols: cols.filter((c) => ['Person', 'DateTime'].includes(c.type)) },
    { title: 'Notes', cols: cols.filter((c) => c.type === 'Note') },
  ];
  return groups.filter((g) => g.cols.length);
}
/** Section card with a title (returns the card so callers can size it). */
export function sectionCard(add, name, x, y, w, title) {
  const card = add('card', { name: 'card' + name, x, y, w, h: 100, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  add('label', { name: 'lblSec' + name, x: x + 20, y: y + 14, w: w - 40, h: 26, props: { text: title, size: 12, weight: 'Semibold', color: '@Text' } });
  return card;
}
/** Make room by shortening multi-line inputs (bottom first, never below 40px) and moving what sits below them up. */
function shrinkNotes(scr, need, listName) {
  const notes = scr.controls.filter((c) => c.type === 'textinput' && c.props.multiline && c.field?.list === listName).sort((a, b) => b.y - a.y);
  let left = need;
  for (const n of notes) {
    const d = Math.min(left, n.h - 40);
    if (d <= 0) continue;
    const edge = n.y + n.h;
    n.h -= d;
    for (const c of scr.controls) if (c !== n) { if (c.y >= edge) c.y -= d; else if (c.y < n.y && c.y + c.h >= edge) c.h -= d; }
    if ((left -= d) <= 0) break;
  }
  return need - Math.max(0, left);
}
/** Attachments placeholder (a dashed box naming the form to insert by hand; see attach.js). */
export const ATTACH_H = { edit: 100, view: 84 };
export function attachSlot(add, list, mode, name, x, y, w) {
  const h = ATTACH_H[mode];
  return add('html', { name, x, y, w, h, props: { html: placeholderHtml(list.name, mode, h) }, attachSlot: { list: list.name, mode } });
}
const hideStatus = (project, list) => {
  const ch = chainFor(project, list.name);
  if (ch) return ch.field;
  const a = project.app.approvalsCfg;
  return a && a.list === list.name ? a.field : null;
};

// ---------------------------------------------------------------- screen templates
export const SCREEN_TEMPLATES = {
  blank: { label: 'Blank', needsList: false },
  home: { label: 'Home dashboard', needsList: false },
  browse: { label: 'Browse list (search + gallery)', needsList: true },
  detail: { label: 'Item details', needsList: true },
  edit: { label: 'New / edit form', needsList: true },
  approvals: { label: 'Approvals queue', needsList: true },
  loading: { label: 'Loading screen (animated)', needsList: false },
};

export function genBlank(project, name) {
  const scr = newScreen(project, name || 'scrScreen');
  project.screens.push(scr);
  headerBlock(project, scr, { title: pretty(scr.name.replace(/^scr/, '')) });
  return scr;
}

export function genEdit(project, list, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrEdit' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add, pad, HH } = ctx(project, scr), T = tag(scr), V = varSel(list.name), ft = features(project);
  const L1 = lower(list.name);
  headerBlock(project, scr, { title: 'Edit ' + L1, titleFx: `If(IsBlank(${V}.ID), "New ${L1}", "Edit ${L1}")`, back: true });
  const hidden = hideStatus(project, list);
  const cols = list.columns.filter((c) => c.name !== hidden && !c.chain);
  const groups = sections(cols);
  const BAR = 72, top = HH + 16, bottom = H - BAR - 12;
  const FIELD = 76, NOTE = 124, SEC_HEAD = 48, SEC_GAP = 14;
  const height = (n) => groups.reduce((h, g) => h + SEC_HEAD + (g.title === 'Notes' ? g.cols.length * NOTE : Math.ceil(g.cols.length / n) * FIELD) + 8 + SEC_GAP, 0);
  const att = !!list.attachments, attH = att && !(!phone && groups.some((g) => g.title === 'Notes')) ? ATTACH_H.edit + SEC_GAP : 0;
  const colsN = phone ? 1 : height(2) <= bottom - top - attH ? 2 : 3;
  let y = top;
  const cw0 = W - pad * 2;
  // tablet: the attachments box sits beside the Notes card
  const beside = att && !phone && groups.some((g) => g.title === 'Notes') ? Math.floor((cw0 - 12) * 0.62) : 0;
  for (const g of groups) {
    const gw = beside && g.title === 'Notes' ? beside : cw0;
    const card = sectionCard(add, 'Sec' + slug(g.title) + T, pad, y, gw, g.title);
    if (gw !== cw0) attachSlot(add, list, 'edit', 'htmlAttach' + T, pad + gw + 12, y, cw0 - gw - 12);
    const inner = gw - 40, n = g.title === 'Notes' ? 1 : colsN, gap = 20, cw = Math.floor((inner - gap * (n - 1)) / n);
    let fy = y + SEC_HEAD, i = 0, rowBottom = fy;
    for (const c of g.cols) {
      const r = inputFor(project, scr, list, c, pad + 20 + (i % n) * (cw + gap), fy, cw);
      rowBottom = Math.max(rowBottom, r.bottom);
      i++;
      if (i % n === 0) fy = rowBottom + 12;
    }
    card.h = rowBottom + 20 - y;
    y += card.h + SEC_GAP;
  }
  // a long form on a phone: shorten the multi-line boxes rather than run under the Save bar
  const under = y - SEC_GAP - bottom;
  if (under > 0) y -= shrinkNotes(scr, under, list.name);
  if (att && !beside) {
    const over = y + ATTACH_H.edit - bottom;
    if (over > 0) y -= shrinkNotes(scr, over, list.name);
    attachSlot(add, list, 'edit', 'htmlAttach' + T, pad, y, cw0);
  }
  // bottom action bar with the "still needed" summary
  add('rect', { name: 'rectBar' + T, x: 0, y: H - BAR, w: W, h: BAR, props: { fill: '@Surface', borderThickness: 0 } });
  add('rect', { name: 'rectBarLine' + T, x: 0, y: H - BAR, w: W, h: 1, props: { fill: '@Border', borderThickness: 0 } });
  const bw = phone ? Math.floor((W - pad * 2 - 12) / 2) : 150;
  const save = add('button', { name: 'btnSave' + T, x: W - pad - bw, y: H - BAR + 16, w: bw, h: 40, props: { text: 'Save', size: 12 }, actions: { OnSelect: [{ type: 'save', list: list.name, then: ft.confirm ? 'confirm' : 'back', validate: true }] } });
  add('button', { name: 'btnCancel' + T, x: W - pad - bw - 12 - bw, y: H - BAR + 16, w: bw, h: 40, props: { text: 'Cancel', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1, size: 12 }, actions: { OnSelect: [{ type: 'back' }] } });
  const miss = missingFx(project, scr, list.name);
  if (miss) {
    const nx = phone ? pad : pad, nw = phone ? W - pad * 2 : W - pad * 2 - bw * 2 - 36;
    const ny = phone ? H - BAR - 44 : H - BAR + 16;
    add('label', { name: 'lblNeeded' + T, x: nx, y: ny, w: nw, h: 40, props: { text: 'Still needed', size: 11, weight: 'Semibold', color: '#92400E', fill: '#FEF3C7' },
      fx: { Text: `"  Still needed: " & ${miss.list}`, Visible: `varTriedSave${slug(list.name)} && (${miss.any})` } });
  }
  void save;
  // confirmation overlay for new items (front-most: added last)
  if (ft.confirm) {
    const vis = { var: 'varDone' + slug(list.name), op: 'true' };
    add('rect', { name: 'rectDoneBg' + T, x: 0, y: 0, w: W, h: H, props: { fill: '@Bg', borderThickness: 0 }, visibleWhen: vis });
    const cw = phone ? W - 48 : 520, cx = (W - cw) / 2, cy = phone ? 260 : 170;
    add('card', { name: 'cardDone' + T, x: cx, y: cy, w: cw, h: 330, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true }, visibleWhen: vis });
    add('card', { name: 'cardDoneChip' + T, x: cx + cw / 2 - 36, y: cy + 36, w: 72, h: 72, props: { fill: '@Success', borderThickness: 0, radius: 36, shadow: false }, visibleWhen: vis });
    add('icon', { name: 'icoDone' + T, x: cx + cw / 2 - 20, y: cy + 52, w: 40, h: 40, props: { icon: 'check', color: '#FFFFFF', stroke: 2.6 }, visibleWhen: vis });
    add('label', { name: 'lblDoneTitle' + T, x: cx + 24, y: cy + 128, w: cw - 48, h: 40, props: { text: `${singular(list.name)} submitted`, size: 18, weight: 'Semibold', color: '@Text', align: 'Center' }, visibleWhen: vis });
    add('label', { name: 'lblDoneSub' + T, x: cx + 32, y: cy + 168, w: cw - 64, h: 48, props: { text: `It is saved in ${pretty(list.name)}. Anyone with access can find it there.`, size: 12, color: '@TextMuted', align: 'Center', valign: 'Top' }, visibleWhen: vis });
    const b2 = (cw - 64 - 12) / 2;
    add('button', { name: 'btnDoneAnother' + T, x: cx + 32, y: cy + 250, w: b2, h: 44, props: { text: 'Add another', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1, size: 12 }, visibleWhen: vis,
      actions: { OnSelect: [{ type: 'setVar', name: 'varDone' + slug(list.name), value: 'false' }, { type: 'reset' }] } });
    add('button', { name: 'btnDoneClose' + T, x: cx + 32 + b2 + 12, y: cy + 250, w: b2, h: 44, props: { text: 'Done', size: 12 }, visibleWhen: vis,
      actions: { OnSelect: [{ type: 'setVar', name: 'varDone' + slug(list.name), value: 'false' }, { type: 'back' }] } });
  }
  scr.onVisible = [`Set(varTriedSave${slug(list.name)}, false)`, ...(ft.confirm ? [`Set(varDone${slug(list.name)}, false)`] : [])].join(';\n');
  return scr;
}

/** Status stepper: one circle per positive stage, coloured from the item's own status value. */
export function stepperFx(list, st, V) {
  const ref = `${V}.${/^[A-Za-z_]\w*$/.test(st.name) ? st.name : `'${st.name}'`}.Value`;
  const steps = st.choices.filter((c) => !NEG.test(c));
  const negs = st.choices.filter((c) => NEG.test(c));
  const ix = `Switch(${ref}, ${steps.map((c, i) => `"${c.replace(/"/g, '""')}", ${i + 1}`).join(', ')}, 0)`;
  const bad = negs.length ? negs.map((c) => `${ref} = "${c.replace(/"/g, '""')}"`).join(' || ') : 'false';
  const cells = steps.map((c, i) => {
    const k = i + 1;
    return `<td style='text-align:center;vertical-align:top'>`
      + `<div style='margin:0 auto;width:30px;height:30px;border-radius:50%;line-height:30px;font-size:13px;font-weight:700;color:#FFFFFF;background:{{=If(ix > ${k}, appTheme.SuccessHex, ix = ${k}, If(bad, appTheme.DangerHex, appTheme.PrimaryHex), appTheme.BorderHex)}}'>{{=If(ix > ${k}, "&#10003;", "${k}")}}</div>`
      + `<div style='font-size:12px;margin-top:7px;font-weight:{{=If(ix = ${k}, "700", "400")}};color:{{=If(ix >= ${k}, appTheme.TextHex, appTheme.TextMutedHex)}}'>${c.replace(/'/g, '&#39;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}</div></td>`;
  }).join('');
  const body = `<table style='width:100%;border-collapse:collapse;table-layout:fixed;margin-top:14px'><tr>${cells}</tr></table>`
    + `{{=If(bad, "<div style='text-align:center;margin-top:6px;font-size:12px;font-weight:600;color:" & appTheme.DangerHex & "'>" & ${esc(ref)} & "</div>", "")}}`;
  return (h) => `With({ix: ${ix}, bad: ${bad}}, ${hx(cardOpen(h) + body + '</div>')})`;
}

/** Details summary card: title, status pill, created by / on, ID. */
export function summaryFx(project, list, sw, sh) {
  const phone = project.app.formFactor === 'phone', V = varSel(list.name), st = statusCol(list);
  const tRef = readFx(project, list.name, list.columns[0].name, V);
  const stRef = st ? `${V}.${/^[A-Za-z_]\w*$/.test(st.name) ? st.name : `'${st.name}'`}.Value` : null;
  return hx(cardOpen(sh, 'padding:18px 22px;')
      + `<div style='font-size:${phone ? 18 : 22}px;font-weight:700;color:{{Text}};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:${phone ? sw - 60 : sw - 420}px'>{{=${esc(tRef)}}}</div>`
      + `<div style='margin-top:10px;font-size:12px;color:{{TextMuted}}'>${st ? pill(stRef, st) + '&nbsp;&nbsp;' : ''}Created by {{=${esc(`${V}.'Created By'.DisplayName`)}}} on {{=Text(${V}.Created, "mm/dd/yyyy")}}  ·  #{{=Text(${V}.ID)}}</div></div>`);
}

export function genDetail(project, list, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrDetail' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add, pad, HH } = ctx(project, scr), T = tag(scr), V = varSel(list.name), ft = features(project);
  const L = /^[A-Za-z_]\w*$/.test(list.name) ? list.name : `'${list.name}'`;
  const titleCol = list.columns[0], st = statusCol(list);
  const has = `!IsBlank(${V}.ID)`, vis = { Visible: has };
  const editScr = opts.editScreen;
  headerBlock(project, scr, { title: singular(list.name) + ' details', back: true,
    right: editScr && !phone ? [{ name: 'btnEdit' + T, text: 'Edit', actions: [{ type: 'navigate', screen: editScr.id }], fx: { Visible: has } }] : [] });
  // nothing selected (opened directly or after a delete)
  const ey = HH + 120;
  add('card', { name: 'cardNoSel' + T, x: W / 2 - 28, y: ey, w: 56, h: 56, props: { fill: '@SurfaceAlt', borderColor: '@Border', borderThickness: 1, radius: 28, shadow: false }, fx: { Visible: `!(${has})` } });
  add('icon', { name: 'icoNoSel' + T, x: W / 2 - 14, y: ey + 14, w: 28, h: 28, props: { icon: 'file', color: '@TextMuted', stroke: 2 }, fx: { Visible: `!(${has})` } });
  add('label', { name: 'lblNoSel' + T, x: pad, y: ey + 70, w: W - pad * 2, h: 30, props: { text: `No ${lower(list.name)} selected`, size: 13, weight: 'Semibold', color: '@Text', align: 'Center' }, fx: { Visible: `!(${has})` } });
  add('button', { name: 'btnNoSel' + T, x: W / 2 - 90, y: ey + 110, w: 180, h: 40, props: { text: 'Back', size: 12 }, actions: { OnSelect: [{ type: 'back' }] }, fx: { Visible: `!(${has})` } });

  // summary card: title, status pill, who/when
  const sy = HH + 16, sh = phone ? 132 : 112, sw = W - pad * 2;
  const tRef = readFx(project, list.name, titleCol.name, V);
  const stRef = st ? `${V}.${/^[A-Za-z_]\w*$/.test(st.name) ? st.name : `'${st.name}'`}.Value` : null;
  add('html', { name: 'htmlSummary' + T, x: pad, y: sy, w: sw, h: sh, props: { html: '' }, fx: { ...vis, HtmlText: summaryFx(project, list, sw, sh) } });
  // actions: delete (two-click), approve/reject when the item is waiting
  const ap = project.app.approvalsCfg && project.app.approvalsCfg.list === list.name ? project.app.approvalsCfg : null;
  const admin = ft.adminGroup ? ' && varIsAdmin' : '';
  const confirmVar = 'varConfirmDelete' + slug(list.name);
  const delBtn = (x, y, w) => add('button', { name: 'btnDelete' + T, x, y, w, h: 40, props: { text: 'Delete', fill: 'transparent', color: '@Danger', borderColor: '@Danger', borderThickness: 1, size: 12 },
    actions: { OnSelect: [{ type: 'delete', list: list.name, then: 'back', confirm: true }] },
    fx: { Visible: has + (ft.adminGroup ? ' && varIsAdmin' : ''), Text: `If(${confirmVar}, "Confirm delete", "Delete")`, Fill: `If(${confirmVar}, appTheme.Danger, RGBA(0, 0, 0, 0))`, Color: `If(${confirmVar}, RGBA(255, 255, 255, 1), appTheme.Danger)` } });
  const decide = (name, text, fill, value, x, y, w) => add('button', { name: name + T, x, y, w, h: 40, props: { text, fill, borderColor: fill, size: 12 },
    actions: { OnSelect: [{ type: 'setField', list: list.name, field: ap.field, value, message: text === 'Approve' ? 'Approved' : 'Rejected' }] },
    fx: { Visible: `${has} && ${stRef} = "${String(ap.pending).replace(/"/g, '""')}"${admin}` } });
  if (!phone) {
    let bx = pad + sw - 22 - 130;
    delBtn(bx, sy + (sh - 40) / 2, 130);
    if (ap) { bx -= 12 + 120; decide('btnReject', 'Reject', '@Danger', ap.no, bx, sy + (sh - 40) / 2, 120); bx -= 12 + 120; decide('btnApprove', 'Approve', '@Success', ap.yes, bx, sy + (sh - 40) / 2, 120); }
  } else if (ap) {
    const bw = (sw - 44 - 12) / 2;
    decide('btnApprove', 'Approve', '@Success', ap.yes, pad + 22, sy + sh + 8, bw);
    decide('btnReject', 'Reject', '@Danger', ap.no, pad + 22 + bw + 12, sy + sh + 8, bw);
  }
  let y = sy + sh + (phone && ap ? 60 : 12);
  const ch = chainFor(project, list.name);
  let chainOverlay = null;
  if (ch) ({ y, overlay: chainOverlay } = chainDetailBlock(project, scr, list, ch, { x: pad, y, w: sw, T, has }));
  // stepper
  else if (ft.stepper && st && st.choices.filter((c) => !NEG.test(c)).length >= 2) {
    const h = 112;
    add('html', { name: 'htmlStepper' + T, x: pad, y, w: sw, h, props: { html: '' }, fx: { ...vis, HtmlText: stepperFx(list, st, V)(h) } });
    y += h + 12;
  }
  // sections
  const cols = list.columns.filter((c) => !c.builtin && !c.chain && c !== st);
  const groups = sections(cols);
  const besideV = !!list.attachments && !phone && cols.some((c) => c.type === 'Note');
  const attV = list.attachments && !besideV ? ATTACH_H.view + 12 : 0;
  const bottomLimit = (phone ? H - 84 : H - 16) - attV;
  const people = groups.find((g) => g.title === 'People and dates');
  const details = groups.find((g) => g.title === 'Details');
  const notes = groups.find((g) => g.title === 'Notes');
  const ROW = 50;
  const fieldGrid = (g, x, gy, w, n) => {
    const card = sectionCard(add, 'Sec' + slug(g.title) + T, x, gy, w, g.title);
    card.fx.Visible = has;
    // the title hides with its card, or it floats over the "nothing selected" state
    const ttl = scr.controls.find((c) => c.name === 'lblSecSec' + slug(g.title) + T);
    if (ttl) ttl.fx = { ...(ttl.fx || {}), Visible: has };
    const inner = w - 40, gap = 16, cw = Math.floor((inner - gap * (n - 1)) / n);
    let i = 0, fy = gy + 46, bottom = fy;
    for (const c of g.cols) {
      const fx0 = x + 20 + (i % n) * (cw + gap), noteH = c.type === 'Note' ? Math.max(40, Math.min(110, bottomLimit - fy - 50)) : 22;
      add('label', { name: 'lblCap' + slug(c.name) + T, x: fx0, y: fy, w: cw, h: 18, props: { text: pretty(c.name), size: 10, weight: 'Semibold', color: '@TextMuted' }, fx: { ...vis } });
      add('label', { name: 'lblVal' + slug(c.name) + T, x: fx0, y: fy + 18, w: cw, h: noteH, props: { text: '-', size: 12, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: list.name, field: c.name }, fx: { ...vis } });
      bottom = Math.max(bottom, fy + 18 + noteH);
      i++;
      if (i % n === 0) fy = bottom + 10;
    }
    card.h = bottom + 16 - gy;
    return card.h;
  };
  if (!phone && details && people) {
    const lw = Math.floor((sw - 12) * 0.6), rw = sw - 12 - lw;
    const h1 = fieldGrid(details, pad, y, lw, details.cols.length > 4 ? 3 : 2);
    const h2 = fieldGrid(people, pad + lw + 12, y, rw, 1 + (people.cols.length > 3 ? 1 : 0));
    y += Math.max(h1, h2) + 12;
  } else {
    for (const g of [details, people].filter(Boolean)) y += fieldGrid(g, pad, y, sw, phone ? 2 : 3) + 12;
  }
  if (besideV) {
    const nw = Math.floor((sw - 12) * 0.6);   // lines up with the People card above
    fieldGrid(notes, pad, y, nw, 1);
    attachSlot(add, list, 'view', 'htmlAttach' + T, pad + nw + 12, y, sw - nw - 12).fx.Visible = has;
  } else {
    if (notes) y += fieldGrid(notes, pad, y, sw, 1) + 12;
    if (list.attachments) attachSlot(add, list, 'view', 'htmlAttach' + T, pad, y, sw).fx.Visible = has;
  }
  // phone: bottom action bar
  if (phone) {
    add('rect', { name: 'rectBar' + T, x: 0, y: H - 72, w: W, h: 72, props: { fill: '@Surface', borderThickness: 0 }, fx: { ...vis } });
    add('rect', { name: 'rectBarLine' + T, x: 0, y: H - 72, w: W, h: 1, props: { fill: '@Border', borderThickness: 0 }, fx: { ...vis } });
    const bw = Math.floor((W - pad * 2 - 12) / 2);
    if (editScr) add('button', { name: 'btnEdit' + T, x: pad + bw + 12, y: H - 56, w: bw, h: 40, props: { text: 'Edit', size: 12 }, actions: { OnSelect: [{ type: 'navigate', screen: editScr.id }] }, fx: { ...vis } });
    delBtn(pad, H - 56, editScr ? bw : W - pad * 2);
  }
  scr.onVisible = `Set(${confirmVar}, false)`;
  // approval options overlay (reassign / admin override): front-most, so added last
  if (chainOverlay) { chainOverlay(); scr.onVisible += `;\nSet(varChainMore${slug(list.name)}, false)`; }
  void L;
  return scr;
}

/**
 * Approval-chain part of the details screen: stepper (each step coloured from its own stage's Decision column),
 * newest-first history, and a decision bar that always shows something (decide / waiting on / outcome + resubmit).
 */
function chainDetailBlock(project, scr, list, ch, { x, y, w, T, has }) {
  const { add, phone } = ctx(project, scr), V = varSel(list.name), vis = { Visible: has };
  const sh = 132;
  const stW = phone ? w : Math.floor((w - 12) * 0.62);
  add('html', { name: 'htmlChain' + T, x, y, w: stW, h: sh, props: { html: '' }, fx: { ...vis, HtmlText: hx(stepperTpl(ch, V, sh, esc)) } });
  const hx0 = phone ? x : x + stW + 12, hy = phone ? y + sh + 12 : y, hw = phone ? w : w - stW - 12, hh = phone ? 120 : sh;
  const card = sectionCard(add, 'SecHistory' + T, hx0, hy, hw, 'History');
  card.h = hh; card.fx.Visible = has;
  const ttl = scr.controls.find((c) => c.name === 'lblSecSecHistory' + T);
  if (ttl) ttl.fx = { ...(ttl.fx || {}), Visible: has };
  add('label', { name: 'lblHistory' + T, x: hx0 + 18, y: hy + 40, w: hw - 36, h: hh - 50, props: { text: '', size: 10, color: '@TextMuted', valign: 'Top' },
    fx: { ...vis, Text: `Coalesce(${V}.ApprovalHistory, "No decisions yet.")` } });
  y = (phone ? hy + hh : y + sh) + 12;
  // decision bar
  const bh = phone ? 116 : 64;
  const can = canDecideFx(ch, V), resub = canResubmitFx(ch, V);
  const bar = add('card', { name: 'cardDecision' + T, x, y, w, h: bh, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true }, fx: { ...vis } });
  void bar;
  const who = (st) => st.approver === 'group' ? `"${(st.groupLabel || st.name + ' group').replace(/"/g, '""')}"` : `Coalesce(${V}.${/^[A-Za-z_]\w*$/.test(st.column) ? st.column : `'${st.column}'`}.DisplayName, "the ${st.name.replace(/"/g, '""')} approver")`;
  const stField = `${V}.${/^[A-Za-z_]\w*$/.test(ch.field) ? ch.field : `'${ch.field}'`}.Value`;
  const late = chainSla(ch) ? ` & If(${overdueFx(ch, V)}, "  ·  overdue", "")` : '';
  const who2 = (st) => ch.reassign ? `Coalesce(${V}.${stageCols(st).delegate}.DisplayName, ${who(st)})` : who(st);
  const stateFx = `Switch(${V}.CurrentStage, ${ch.stages.map((st, i) => `${i + 1}, "Waiting on ${st.name.replace(/"/g, '""')} approval: " & ${who2(st)}${late}`).join(', ')}, If(${stField} = "${ch.values.rejected.replace(/"/g, '""')}", "Rejected. The requester can change it and resubmit.", If(${stField} = "${ch.values.approved.replace(/"/g, '""')}", "Approved by every stage.", ${stField})))`;
  add('label', { name: 'lblDecisionState' + T, x: x + 18, y: y + (phone ? 12 : 12), w: w - 36 - (phone ? 0 : 170), h: phone ? 44 : 40, props: { text: '', size: 12, weight: 'Semibold', color: '@Text' },
    fx: { Visible: `${has} && !(${can})`, Text: stateFx } });
  if (ch.resubmit) add('button', { name: 'btnResubmit' + T, x: phone ? x + 18 : x + w - 18 - 150, y: phone ? y + 60 : y + 12, w: phone ? w - 36 : 150, h: 40, props: { text: 'Resubmit', size: 12 },
    actions: { OnSelect: [{ type: 'resubmit', list: list.name }] }, fx: { Visible: resub } });
  const bw = 112;
  const extras = !!(ch.reassign || ch.override), admin = !!features(project).adminGroup;
  const cx = x + 18, cw = phone ? w - 36 : w - 36 - bw * 2 - 24 - (extras ? bw + 12 : 0);
  const cy = phone ? y + 12 : y + 12;
  const comment = add('textinput', { name: 'txtDecisionComment' + T, x: cx, y: cy, w: cw, h: 40, props: { hint: 'Comment (required to reject)', defaultText: '', size: 12 }, fx: { Visible: can } });
  const by = phone ? y + 64 : cy, bx = phone ? x + 18 : cx + cw + 12, pbw = phone ? Math.floor((w - 36 - 12 * (extras ? 2 : 1)) / (extras ? 3 : 2)) : bw;
  add('button', { name: 'btnChainApprove' + T, x: bx, y: by, w: pbw, h: 40, props: { text: 'Approve', fill: '@Success', borderColor: '@Success', size: 12 },
    actions: { OnSelect: [{ type: 'decide', list: list.name, decision: 'Approved', commentCtl: comment.name }] }, fx: { Visible: can } });
  add('button', { name: 'btnChainReject' + T, x: bx + pbw + 12, y: by, w: pbw, h: 40, props: { text: 'Reject', fill: '@Danger', borderColor: '@Danger', size: 12 },
    actions: { OnSelect: [{ type: 'decide', list: list.name, decision: 'Rejected', commentCtl: comment.name }] },
    fx: { Visible: can, DisplayMode: `If(IsBlank(Trim(${comment.name}.Text)), DisplayMode.Disabled, DisplayMode.Edit)` } });
  if (!extras) return { y: y + bh + 12 };
  // round 6: reassign / admin override live in an overlay opened from here
  const moreVar = 'varChainMore' + slug(list.name);
  const reVis = ch.reassign ? canReassignFx(ch, V, admin) : '';
  const ovVis = ch.override && admin ? `!IsBlank(${V}.ID) && ${V}.CurrentStage > 0 && varIsAdmin` : '';
  const optVis = [reVis, ovVis].filter(Boolean).map((v) => `(${v})`).join(' || ') || 'false';
  add('button', { name: 'btnChainMore' + T, x: phone ? x + 18 + (pbw + 12) * 2 : x + w - 18 - bw, y: phone ? y + 64 : cy, w: phone ? pbw : bw, h: 40,
    props: { text: 'Options', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1, size: 12 },
    actions: { OnSelect: [{ type: 'setVar', name: moreVar, value: 'true' }] }, fx: { Visible: optVis } });
  const overlay = () => chainOptionsOverlay(project, scr, list, ch, { T, moreVar, reVis, ovVis });
  return { y: y + bh + 12, overlay };
}

/** "Approval options" overlay: reassign the waiting stage (picker + button) and/or admin override (reason + approve/reject). */
function chainOptionsOverlay(project, scr, list, ch, { T, moreVar, reVis, ovVis }) {
  const { W, H, phone, add } = ctx(project, scr);
  const on = (v) => (v ? `${moreVar} && (${v})` : 'false');
  const cw = phone ? W - 32 : 560, secR = ch.reassign ? 112 : 0, secO = ovVis ? 156 : 0;
  const ch0 = 64 + secR + secO + 8, cx = (W - cw) / 2, cy = Math.max(80, Math.round((H - ch0) / 2));
  add('rect', { name: 'rectMoreDim' + T, x: 0, y: 0, w: W, h: H, props: { fill: '#0F172A73', borderThickness: 0 }, fx: { Visible: moreVar } });
  add('card', { name: 'cardMore' + T, x: cx, y: cy, w: cw, h: ch0, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true }, fx: { Visible: moreVar } });
  add('label', { name: 'lblMoreTitle' + T, x: cx + 24, y: cy + 18, w: cw - 124, h: 30, props: { text: 'Approval options', size: 16, weight: 'Semibold', color: '@Text' }, fx: { Visible: moreVar } });
  add('button', { name: 'btnMoreClose' + T, x: cx + cw - 24 - 76, y: cy + 16, w: 76, h: 34, props: { text: 'Close', fill: 'transparent', color: '@Primary', borderColor: '@Border', borderThickness: 1, size: 11 },
    actions: { OnSelect: [{ type: 'setVar', name: moreVar, value: 'false' }] }, fx: { Visible: moreVar } });
  const ry = cy + 64, oy = ry + secR;
  // admin override first in the layer order, so the people picker's results list (added after) draws over it
  if (ovVis) {
    add('label', { name: 'lblOverride' + T, x: cx + 24, y: oy, w: cw - 48, h: 24, props: { text: 'Admin override', size: 13, weight: 'Semibold', color: '@Text' }, fx: { Visible: on(ovVis) } });
    add('label', { name: 'lblOverrideHint' + T, x: cx + 24, y: oy + 26, w: cw - 48, h: 20, props: { text: 'Decide for the approver. The reason is kept in the history.', size: 11, color: '@TextMuted' }, fx: { Visible: on(ovVis) } });
    const reason = add('textinput', { name: 'txtOverrideReason' + T, x: cx + 24, y: oy + 52, w: cw - 48, h: 40, props: { hint: 'Reason (required)', defaultText: '', size: 12 }, fx: { Visible: on(ovVis) } });
    const bw = Math.floor((cw - 48 - 12) / 2), dis = `If(IsBlank(Trim(${reason.name}.Text)), DisplayMode.Disabled, DisplayMode.Edit)`;
    add('button', { name: 'btnOverrideApprove' + T, x: cx + 24, y: oy + 100, w: bw, h: 40, props: { text: 'Approve (override)', fill: '@Success', borderColor: '@Success', size: 12 },
      actions: { OnSelect: [{ type: 'override', list: list.name, decision: 'Approved', commentCtl: reason.name }] }, fx: { Visible: on(ovVis), DisplayMode: dis } });
    add('button', { name: 'btnOverrideReject' + T, x: cx + 24 + bw + 12, y: oy + 100, w: bw, h: 40, props: { text: 'Reject (override)', fill: '@Danger', borderColor: '@Danger', size: 12 },
      actions: { OnSelect: [{ type: 'override', list: list.name, decision: 'Rejected', commentCtl: reason.name }] }, fx: { Visible: on(ovVis), DisplayMode: dis } });
  }
  if (ch.reassign) {
    add('label', { name: 'lblReassign' + T, x: cx + 24, y: ry, w: cw - 48, h: 24, props: { text: 'Reassign this stage', size: 13, weight: 'Semibold', color: '@Text' }, fx: { Visible: on(reVis) } });
    add('label', { name: 'lblReassignHint' + T, x: cx + 24, y: ry + 26, w: cw - 48, h: 20, props: { text: 'The person you pick can decide it instead.', size: 11, color: '@TextMuted' }, fx: { Visible: on(reVis) } });
    const pw = cw - 48 - 12 - 130;
    const pick = add('person', { name: 'ppReassign' + T, x: cx + 24, y: ry + 52, w: pw, h: 40, props: { hint: 'Search people', size: 12 }, fx: { Visible: on(reVis) } });
    add('button', { name: 'btnReassign' + T, x: cx + 24 + pw + 12, y: ry + 52, w: 130, h: 40, props: { text: 'Reassign', size: 12 },
      actions: { OnSelect: [{ type: 'reassign', list: list.name, picker: pick.name }] }, fx: { Visible: on(reVis), DisplayMode: `If(IsBlank(varPick${pick.name}), DisplayMode.Disabled, DisplayMode.Edit)` } });
  }
}


/** Approval queue for a chain: items at a stage the signed-in user may decide, with a decision panel (tablet). */
export function genChainApprovals(project, list, ch, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrApprovals' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add, pad, HH } = ctx(project, scr), T = tag(scr), V = varSel(list.name);
  headerBlock(project, scr, { title: 'Approvals', back: !!opts.back });
  const qw = phone ? W - pad * 2 : 500, gy = HH + 56;
  const gal = add('gallery', { name: 'galQueue' + T, x: pad - 2, y: gy, w: qw + 4, h: H - gy - 12, props: { templateSize: 84 },
    data: { list: list.name, searchCtl: '', searchField: '', filterField: '', filterValue: '', sortField: '', sortDir: 'Ascending', chainQueue: true },
    fx: { Items: queueFx(ch) },
    actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: phone ? (opts.detailScreen?.id || '') : '', fresh: true }] } });
  galleryRows(project, scr, gal, list);
  if (chainSla(ch)) add('label', { name: 'lblOverdue' + T, x: qw - 104, y: 47, w: 64, h: 17, props: { text: 'Overdue', size: 9, weight: 'Semibold', color: '#FFFFFF', fill: '@Danger', align: 'Center' },
    fx: { Visible: overdueFx(ch, 'ThisItem') } }, gal.children);
  add('label', { name: 'lblQueue' + T, x: pad, y: HH + 14, w: qw, h: 32, props: { text: 'Waiting for your decision', size: 13, weight: 'Semibold', color: '@Text' },
    fx: { Text: `"Waiting for your decision (" & CountRows(${gal.name}.AllItems) & ")"` } });
  emptyState(project, scr, gal, { x: pad, y: gy + 60, w: qw, title: 'All caught up', hint: 'Nothing is waiting for your decision.' });
  if (phone) return scr;
  const px = pad + qw + 20, pw = W - px - pad;
  const has = canDecideFx(ch, V);
  const panel = add('card', { name: 'cardPanel' + T, x: px, y: HH + 16, w: pw, h: H - HH - 28, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  add('icon', { name: 'icoPick' + T, x: px + pw / 2 - 16, y: panel.y + panel.h / 2 - 50, w: 32, h: 32, props: { icon: 'eye', color: '@TextMuted', stroke: 2 }, fx: { Visible: `!(${has})` } });
  add('label', { name: 'lblPick' + T, x: px + 20, y: panel.y + panel.h / 2 - 10, w: pw - 40, h: 30, props: { text: `Select a ${lower(list.name)} to decide it`, size: 13, color: '@TextMuted', align: 'Center' }, fx: { Visible: `!(${has})` } });
  add('label', { name: 'lblPanelTitle' + T, x: px + 24, y: panel.y + 20, w: pw - 48, h: 36, props: { text: 'Title', size: 17, weight: 'Semibold', color: '@Text' }, bind: { src: 'sel', list: list.name, field: list.columns[0].name }, fx: { Visible: has } });
  add('label', { name: 'lblPanelStage' + T, x: px + 24, y: panel.y + 56, w: pw - 48, h: 24, props: { text: '', size: 11, weight: 'Semibold', color: '@Primary' },
    fx: { Visible: has, Text: `Switch(${V}.CurrentStage, ${ch.stages.map((st, i) => `${i + 1}, "Stage ${i + 1} of ${ch.stages.length}: ${st.name.replace(/"/g, '""')}"`).join(', ')}, "")` } });
  add('label', { name: 'lblPanelMeta' + T, x: px + 24, y: panel.y + 80, w: pw - 48, h: 24, props: { text: '', size: 11, color: '@TextMuted' },
    fx: { Visible: has, Text: `"Submitted by " & ${V}.'Created By'.DisplayName & " on " & Text(${V}.Created, "mm/dd/yyyy")` } });
  if ((ch.reassign || ch.override) && opts.detailScreen) {
    // reassign and admin override live on the details screen
    scr.controls.find((c) => c.name === 'lblPanelTitle' + T).w = pw - 48 - 142;
    add('button', { name: 'btnPanelDetails' + T, x: px + pw - 24 - 130, y: panel.y + 20, w: 130, h: 36, props: { text: 'Options', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1, size: 11 },
      actions: { OnSelect: [{ type: 'navigate', screen: opts.detailScreen.id }] }, fx: { Visible: has } });
  }
  let y = panel.y + 116;
  const pcols = list.columns.filter((c) => !c.builtin && !c.chain && c.name !== ch.field);
  const ordered = [...pcols.filter((c) => c.type !== 'Note'), ...pcols.filter((c) => c.type === 'Note')];
  const maxY = panel.y + panel.h - 150;
  for (const col of ordered) {
    const h = col.type === 'Note' ? 54 : 24;
    if (y + h > maxY) break;
    add('label', { name: 'lblPCap' + slug(col.name) + T, x: px + 24, y, w: 170, h: 24, props: { text: pretty(col.name), size: 11, weight: 'Semibold', color: '@TextMuted' }, fx: { Visible: has } });
    add('label', { name: 'lblPVal' + slug(col.name) + T, x: px + 200, y, w: pw - 224, h, props: { text: '-', size: 12, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: list.name, field: col.name }, fx: { Visible: has } });
    y += h + 10;
  }
  const by = panel.y + panel.h - 64, bw = 150;
  const comment = add('textinput', { name: 'txtDecisionComment' + T, x: px + 24, y: by - 56, w: pw - 48, h: 40, props: { hint: 'Comment (required to reject)', defaultText: '', size: 12 }, fx: { Visible: has } });
  add('button', { name: 'btnChainApprove' + T, x: px + pw - 24 - bw * 2 - 12, y: by, w: bw, h: 44, props: { text: 'Approve', fill: '@Success', borderColor: '@Success', size: 12 }, fx: { Visible: has },
    actions: { OnSelect: [{ type: 'decide', list: list.name, decision: 'Approved', commentCtl: comment.name }] } });
  add('button', { name: 'btnChainReject' + T, x: px + pw - 24 - bw, y: by, w: bw, h: 44, props: { text: 'Reject', fill: '@Danger', borderColor: '@Danger', size: 12 },
    fx: { Visible: has, DisplayMode: `If(IsBlank(Trim(${comment.name}.Text)), DisplayMode.Disabled, DisplayMode.Edit)` },
    actions: { OnSelect: [{ type: 'decide', list: list.name, decision: 'Rejected', commentCtl: comment.name }] } });
  return scr;
}

export function genBrowse(project, list, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrBrowse' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add, pad, HH } = ctx(project, scr), T = tag(scr), ft = features(project);
  const newAct = opts.editScreen ? [{ type: 'newItem', list: list.name, screen: opts.editScreen.id }] : null;
  headerBlock(project, scr, { title: pretty(list.name), back: !!opts.back, right: newAct && !phone ? [{ name: 'btnNew' + T, text: '+ New ' + lower(list.name), actions: newAct, w: 170 }] : [] });
  const st = statusCol(list);
  const scopeVar = 'varScope' + slug(list.name);
  // toolbar: search, All/Mine, status filter, count
  const ty = HH + 16;
  const sw = phone ? W - pad * 2 : 340;
  const search = add('textinput', { name: 'txtSearch' + T, x: pad, y: ty, w: sw, h: 40, props: { hint: 'Search by ' + pretty(list.columns[0].name).toLowerCase(), defaultText: '', size: 12 } });
  add('icon', { name: 'icoSearch' + T, x: pad + sw - 32, y: ty + 9, w: 22, h: 22, props: { icon: 'search', color: '@TextMuted', stroke: 2 } });
  const ry = phone ? ty + 52 : ty;
  let rx = phone ? pad : pad + sw + 12;
  const seg = (name, text, val, x) => add('button', { name: name + T, x, y: ry, w: 84, h: 40, props: { text, size: 12, weight: 'Semibold', borderColor: '@Border', borderThickness: 1 },
    actions: { OnSelect: [{ type: 'setVar', name: scopeVar, value: `"${val}"` }] },
    fx: { Fill: `If(${scopeVar} ${val === 'mine' ? '=' : '<>'} "mine", appTheme.Primary, appTheme.Surface)`, Color: `If(${scopeVar} ${val === 'mine' ? '=' : '<>'} "mine", appTheme.OnPrimary, appTheme.Text)` } });
  seg('btnScopeAll', 'All', 'all', rx); seg('btnScopeMine', 'Mine', 'mine', rx + 88);
  rx += 88 + 84 + 12;
  let dd = null;
  if (st) dd = add('dropdown', { name: 'ddStatus' + T, x: rx, y: ry, w: phone ? W - pad - rx : 200, h: 40, props: { items: ['All statuses', ...st.choices].join('\n'), size: 12 } });
  const gy = phone ? ry + 56 : ty + 56;
  const gal = add('gallery', { name: 'galList' + T, x: pad - 2, y: gy, w: W - pad * 2 + 4, h: H - gy - (phone ? 12 : 12), props: { templateSize: 84 },
    data: { list: list.name, searchCtl: search.name, searchField: list.columns[0].name, sortField: 'Modified', sortDir: 'Descending', filterField: '', filterValue: '',
      scopeVar, ...(dd ? { statusCtl: dd.name, statusField: st.name, statusAll: 'All statuses' } : {}) },
    actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: opts.detailScreen?.id || '', fresh: true }] } });
  galleryRows(project, scr, gal, list);
  if (!phone) add('label', { name: 'lblCount' + T, x: W - pad - 160, y: ty, w: 160, h: 40, props: { text: '0 items', size: 11, color: '@TextMuted', align: 'Right' },
    fx: { Text: `CountRows(${gal.name}.AllItems) & If(CountRows(${gal.name}.AllItems) = 1, " ${lower(list.name)}", " ${pretty(list.name).toLowerCase()}")` } });
  emptyState(project, scr, gal, { x: pad, y: gy + 60, w: W - pad * 2, title: `No ${pretty(list.name).toLowerCase()} here`, hint: 'Clear the search, switch to All, or add a new one.', button: newAct ? { text: '+ New ' + lower(list.name), actions: newAct } : null });
  if (phone && newAct) add('button', { name: 'btnFab' + T, x: W - pad - 60, y: H - pad - 60, w: 60, h: 60, props: { text: '+', size: 22, weight: 'Semibold', radius: 30, borderThickness: 0 }, actions: { OnSelect: newAct } });
  scr.onVisible = `If(IsBlank(${scopeVar}), Set(${scopeVar}, "${ft.scope === 'mine' ? 'mine' : 'all'}"))`;
  return scr;
}

export function genApprovals(project, list, cfg = {}) {
  const st = cfg.field ? list.columns.find((c) => c.name === cfg.field) : statusCol(list);
  const scr = newScreen(project, cfg.name || 'scrApprovals' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add, pad, HH } = ctx(project, scr), T = tag(scr), V = varSel(list.name);
  const pending = cfg.pending || st?.choices?.[0] || '';
  const yes = cfg.yes || st?.choices?.find((c) => /approv|done|complete/i.test(c)) || st?.choices?.[1] || '';
  const no = cfg.no || st?.choices?.find((c) => /reject|cancel|declin/i.test(c)) || st?.choices?.[2] || '';
  headerBlock(project, scr, { title: 'Approvals', back: !!cfg.back });
  const qw = phone ? W - pad * 2 : 500;
  const gy = HH + 56;
  const gal = add('gallery', { name: 'galQueue' + T, x: pad - 2, y: gy, w: qw + 4, h: H - gy - 12, props: { templateSize: 84 },
    data: { list: list.name, searchCtl: '', searchField: '', filterField: st?.name || '', filterValue: pending, sortField: 'Created', sortDir: 'Ascending' },
    actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: phone ? (cfg.detailScreen?.id || '') : '', fresh: true }] } });
  galleryRows(project, scr, gal, list);
  add('label', { name: 'lblQueue' + T, x: pad, y: HH + 14, w: qw, h: 32, props: { text: 'Waiting for a decision', size: 13, weight: 'Semibold', color: '@Text' },
    fx: { Text: `"Waiting for a decision (" & CountRows(${gal.name}.AllItems) & ")"` } });
  emptyState(project, scr, gal, { x: pad, y: gy + 60, w: qw, title: 'All caught up', hint: `Nothing is waiting for a decision.` });
  if (phone || !st) return scr;
  const px = pad + qw + 20, pw = W - px - pad;
  const has = `!IsBlank(${V}.ID) && ${V}.${/^[A-Za-z_]\w*$/.test(st.name) ? st.name : `'${st.name}'`}.Value = "${pending.replace(/"/g, '""')}"`;
  const panel = add('card', { name: 'cardPanel' + T, x: px, y: HH + 16, w: pw, h: H - HH - 28, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  add('icon', { name: 'icoPick' + T, x: px + pw / 2 - 16, y: panel.y + panel.h / 2 - 50, w: 32, h: 32, props: { icon: 'eye', color: '@TextMuted', stroke: 2 }, fx: { Visible: `!(${has})` } });
  add('label', { name: 'lblPick' + T, x: px + 20, y: panel.y + panel.h / 2 - 10, w: pw - 40, h: 30, props: { text: `Select a ${lower(list.name)} to review it`, size: 13, color: '@TextMuted', align: 'Center' }, fx: { Visible: `!(${has})` } });
  add('label', { name: 'lblPanelTitle' + T, x: px + 24, y: panel.y + 20, w: pw - 48, h: 36, props: { text: 'Title', size: 17, weight: 'Semibold', color: '@Text' }, bind: { src: 'sel', list: list.name, field: list.columns[0].name }, fx: { Visible: has } });
  add('label', { name: 'lblPanelMeta' + T, x: px + 24, y: panel.y + 56, w: pw - 48, h: 24, props: { text: '', size: 11, color: '@TextMuted' },
    fx: { Visible: has, Text: `"Submitted by " & ${V}.'Created By'.DisplayName & " on " & Text(${V}.Created, "mm/dd/yyyy")` } });
  let y = panel.y + 96;
  const pcols = list.columns.filter((c) => !c.builtin && !c.chain && c !== st);
  const ordered = [...pcols.filter((c) => c.type !== 'Note'), ...pcols.filter((c) => c.type === 'Note')];
  const maxY = panel.y + panel.h - 84;
  for (const col of ordered) {
    const h = col.type === 'Note' ? 54 : 24;
    if (y + h > maxY) break;
    add('label', { name: 'lblPCap' + slug(col.name) + T, x: px + 24, y, w: 170, h: 24, props: { text: pretty(col.name), size: 11, weight: 'Semibold', color: '@TextMuted' }, fx: { Visible: has } });
    add('label', { name: 'lblPVal' + slug(col.name) + T, x: px + 200, y, w: pw - 224, h, props: { text: '-', size: 12, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: list.name, field: col.name }, fx: { Visible: has } });
    y += h + 10;
  }
  const by = panel.y + panel.h - 64, bw = 150;
  const L = /^[A-Za-z_]\w*$/.test(list.name) ? list.name : `'${list.name}'`;
  const admin = features(project).adminGroup ? ' && varIsAdmin' : '';
  add('button', { name: 'btnApprove' + T, x: px + pw - 24 - bw * 2 - 12, y: by, w: bw, h: 44, props: { text: 'Approve', fill: '@Success', borderColor: '@Success', size: 12 }, fx: { Visible: has + admin },
    actions: { OnSelect: [{ type: 'setField', list: list.name, field: st.name, value: yes, message: 'Approved' }, { type: 'custom', fx: `Set(${V}, Defaults(${L}))` }] } });
  add('button', { name: 'btnReject' + T, x: px + pw - 24 - bw, y: by, w: bw, h: 44, props: { text: 'Reject', fill: '@Danger', borderColor: '@Danger', size: 12 }, fx: { Visible: has + admin },
    actions: { OnSelect: [{ type: 'setField', list: list.name, field: st.name, value: no, message: 'Rejected' }, { type: 'custom', fx: `Set(${V}, Defaults(${L}))` }] } });
  return scr;
}

export function genHome(project, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrHome');
  project.screens.unshift(scr);
  const { W, H, phone, add, pad, HH } = ctx(project, scr), T = tag(scr), ft = features(project);
  headerBlock(project, scr, { title: project.app.name, avatar: true });
  // hero: greeting + primary action
  const hy = HH + 16, hh = phone ? 150 : 116, hw = W - pad * 2;
  add('html', { name: 'htmlHero' + T, x: pad, y: hy, w: hw, h: hh, props: { html: '' },
    fx: { HtmlText: hx(`<div style='margin:0;box-sizing:border-box;height:${hh - 10}px;padding:${phone ? 18 : 22}px 26px;border-radius:{{Radius}}px;background:linear-gradient(115deg,{{Primary}},{{Accent}});font-family:{{Font}};color:#FFFFFF'>`
      + `<div style='font-size:${phone ? 20 : 24}px;font-weight:700'>{{=If(Hour(Now()) < 12, "Good morning", Hour(Now()) < 17, "Good afternoon", "Good evening") & ", " & ${esc('First(Split(User().FullName, " ")).Value')}}}</div>`
      + `<div style='font-size:13px;margin-top:6px;opacity:0.9'>${(opts.tagline || 'Here is what needs your attention today.').replace(/'/g, '&#39;')}</div></div>`) } });
  const prim = opts.primary;
  if (prim) add('button', { name: 'btnHeroNew' + T, x: phone ? pad + 26 : pad + hw - 26 - 190, y: phone ? hy + hh - 56 : hy + (hh - 44) / 2, w: phone ? 190 : 190, h: 44, props: { text: prim.text, fill: '#FFFFFF', color: '@Primary', borderThickness: 0, size: 12, weight: 'Semibold' }, actions: { OnSelect: prim.actions } });
  let y = hy + hh + 16;
  // KPIs: one per status value of the first list that has one
  const kl = project.lists.find((l) => statusCol(l) && !l.system);
  if (kl) {
    const st = statusCol(kl), vals = st.choices.slice(0, 4), perRow = phone ? 2 : vals.length;
    const gap = phone ? 12 : 16, w = Math.floor((hw - gap * (perRow - 1)) / perRow), h = phone ? 96 : 100;
    const L = /^[A-Za-z_]\w*$/.test(kl.name) ? kl.name : `'${kl.name}'`;
    const S = /^[A-Za-z_]\w*$/.test(st.name) ? st.name : `'${st.name}'`;
    vals.forEach((v, i) => {
      const x = pad + (i % perRow) * (w + gap), ky = y + Math.floor(i / perRow) * (h + gap), t = tone(v, st.choices.indexOf(v)), k = slug(v) + T;
      add('html', { name: 'htmlKpi' + k, x, y: ky, w, h, props: { html: '' }, fx: { HtmlText: hx(cardOpen(h, `padding:${phone ? 14 : 16}px 16px 0 ${phone ? 62 : 70}px;`)
        + `<div style='font-size:${phone ? 22 : 26}px;font-weight:700;color:{{Text}};line-height:1.15'>{{=Text(CountRows(Filter(${L}, ${S}.Value = "${v.replace(/"/g, '""')}")))}}</div>`
        + `<div style='font-size:12px;color:{{TextMuted}};margin-top:4px;white-space:nowrap'>${v.replace(/'/g, '&#39;').replace(/</g, '&lt;')}</div></div>`) } });
      iconChip(add, 'Kpi' + k, x + (phone ? 16 : 18), ky + (h - 40) / 2, 40, toneIcon(t), '@' + t);
    });
    y += Math.ceil(vals.length / perRow) * (h + gap);
  }
  // quick actions + lists
  const tiles = opts.tiles || [];
  const mine = opts.mine, waiting = opts.waiting;
  const colGap = 16;
  const leftW = phone ? hw : Math.floor((hw - colGap) * 0.6), rightW = phone ? hw : hw - colGap - leftW;
  const rightX = phone ? pad : pad + leftW + colGap;
  const sectionTitle = (name, text, x, sy, w, link) => {
    add('label', { name: 'lblSec' + name + T, x, y: sy, w: w - 100, h: 28, props: { text, size: 13, weight: 'Semibold', color: '@Text' } });
    if (link) add('button', { name: 'btnSee' + name + T, x: x + w - 96, y: sy, w: 96, h: 28, props: { text: link.text, fill: 'transparent', color: '@Primary', borderThickness: 0, size: 11, weight: 'Semibold' }, actions: { OnSelect: link.actions } });
  };
  const listBlock = (key, cfg, x, sy, w, h) => {
    sectionTitle(key, cfg.title, x, sy, w, cfg.link);
    const gal = add('gallery', { name: 'gal' + key + T, x: x - 2, y: sy + 34, w: w + 4, h: h - 34, props: { templateSize: 72 },
      data: { list: cfg.list.name, searchCtl: '', searchField: '', filterField: cfg.filterField || '', filterValue: cfg.filterValue || '', extraFilter: cfg.extraFilter || '', sortField: 'Modified', sortDir: 'Descending' },
      actions: { OnSelect: [{ type: 'openItem', list: cfg.list.name, screen: cfg.detailScreen?.id || '', fresh: true }] } });
    if (cfg.chainQueue) { gal.data.chainQueue = true; gal.data.sortField = ''; gal.fx.Items = queueFx(chainFor(project, cfg.list.name)); }
    galleryRows(project, scr, gal, cfg.list, { compact: true });
    emptyState(project, scr, gal, { x, y: sy + 60, w, title: cfg.emptyTitle, hint: cfg.emptyHint, button: cfg.emptyButton });
  };
  if (phone) {
    // two quick tiles in a row, then my items to the bottom
    const qt = tiles.slice(0, 2), tw = Math.floor((hw - 12) / 2);
    qt.forEach((t, i) => tile(add, T, t, pad + i * (tw + 12), y, tw, 72, true));
    if (qt.length) y += 72 + 16;
    if (mine) listBlock('Mine', mine, pad, y, hw, H - y - 12);
  } else {
    const h = H - y - 12;
    if (mine) listBlock('Mine', mine, pad, y, leftW, h);
    if (waiting) listBlock('Waiting', waiting, rightX, y, rightW, h);
    else {
      sectionTitle('Quick', 'Quick actions', rightX, y, rightW, null);
      let ty = y + 34;
      for (const t of tiles.slice(0, Math.max(1, Math.floor((h - 34) / 84)))) { tile(add, T, t, rightX, ty, rightW, 72, false); ty += 84; }
    }
  }
  return scr;
}
/** Quick-action tile: card, icon chip, title, subtitle, transparent button on top. */
function tile(add, T, t, x, y, w, h, small) {
  const k = slug(t.label) + T;
  add('card', { name: 'cardTile' + k, x, y, w, h, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  iconChip(add, 'Tile' + k, x + 16, y + (h - 40) / 2, 40, t.icon || 'list', '@Primary', '@Accent');
  add('label', { name: 'lblTile' + k, x: x + 68, y: y + (small ? 14 : 12), w: w - 84, h: 24, props: { text: t.label, size: 13, weight: 'Semibold', color: '@Text' } });
  add('label', { name: 'lblTileSub' + k, x: x + 68, y: y + (small ? 36 : 36), w: w - 84, h: 22, props: { text: t.sub || '', size: 11, color: '@TextMuted' } });
  add('button', { name: 'btnTile' + k, x, y, w, h, props: { text: '', fill: 'transparent', color: 'transparent', borderColor: 'transparent', borderThickness: 0, radius: 'theme' }, actions: { OnSelect: t.actions }, fx: t.fx || {} });
}

/** Animated loading screen: self-starting timer, SMIL emblem (original monogram, no seal), separate progress ring. */
export function genLoading(project, { next, lists = [] } = {}) {
  const scr = newScreen(project, 'scrLoading');
  scr.role = 'loading';
  project.screens.unshift(scr);
  const { W, H, phone, add } = ctx(project, scr), T = tag(scr), ft = features(project);
  add('card', { name: 'cardLoadBg' + T, x: 0, y: 0, w: W, h: H, props: { fill: '@HeaderBg', fill2: '@Primary', gradAngle: 135, borderThickness: 0, radius: 0, shadow: false } });
  const initials = (String(project.app.name || 'App').match(/\b[A-Za-z]/g) || ['A']).slice(0, 2).join('').toUpperCase();
  const E = phone ? 200 : 180, ex = (W - E) / 2, ey = phone ? 300 : 170;
  // emblem: hexagon frame, monogram, orbiting dot; fly-in on first render (gated on locInited so it is not cut off by the load splash)
  const emblem = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 200 200'>`
    + `<g opacity='0'><animate attributeName='opacity' from='0' to='1' dur='0.7s' fill='freeze'/>`
    + `<animateTransform attributeName='transform' type='scale' values='0.6;1.04;1' keyTimes='0;0.7;1' dur='0.8s' additive='sum' fill='freeze'/>`
    + `<path d='M100 14 L174 57 L174 143 L100 186 L26 143 L26 57 Z' fill='rgba(255,255,255,0.10)' stroke='{{Accent}}' stroke-width='5' stroke-linejoin='round'/>`
    + `<text x='100' y='120' text-anchor='middle' font-family='Segoe UI, Arial, sans-serif' font-size='58' font-weight='700' fill='#FFFFFF'>${initials}</text></g>`
    + `<g><animateTransform attributeName='transform' type='rotate' from='0 100 100' to='360 100 100' dur='3.2s' repeatCount='indefinite'/>`
    + `<circle cx='100' cy='8' r='6' fill='{{Accent}}'/></g></svg>`;
  add('image', { name: 'imgEmblem' + T, x: ex, y: ey, w: E, h: E, props: { src: '' }, fx: { Image: `If(!locInited, "", "data:image/svg+xml;utf8," & EncodeUrl(${hx(emblem)}))` } });
  const R = 64, ry = ey + E + 36, rx = (W - R) / 2;
  const ring = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><circle cx='32' cy='32' r='26' fill='none' stroke='rgba(255,255,255,0.18)' stroke-width='6'/>`
    + `<circle cx='32' cy='32' r='26' fill='none' stroke='#FFFFFF' stroke-width='6' stroke-linecap='round' transform='rotate(-90 32 32)' stroke-dasharray='{{=Text(Round(163 * Min(locStep, 3) / 3, 0))}} 163'/></svg>`;
  add('image', { name: 'imgRing' + T, x: rx, y: ry, w: R, h: R, props: { src: '' }, fx: { Image: `"data:image/svg+xml;utf8," & EncodeUrl(${hx(ring)})` } });
  add('label', { name: 'lblLoadApp' + T, x: 0, y: ry + R + 20, w: W, h: 40, props: { text: project.app.name, size: phone ? 18 : 20, weight: 'Semibold', color: '#FFFFFF', align: 'Center' } });
  add('label', { name: 'lblLoadStep' + T, x: 0, y: ry + R + 60, w: W, h: 28, props: { text: 'Starting…', size: 12, color: '#FFFFFF', align: 'Center' },
    fx: { Text: `Switch(locStep, 1, "Connecting to your lists…", 2, "${ft.adminGroup ? 'Checking your access…' : 'Getting things ready…'}", 3, "Ready", "Starting…")` } });
  const refresh = lists.filter((l) => !l.system).slice(0, 4).map((l) => `Refresh(${/^[A-Za-z_]\w*$/.test(l.name) ? l.name : `'${l.name}'`})`);
  const nextName = next?.name || 'scrHome';
  const step2 = [ft.usage ? usageFx(project) : ''].filter(Boolean);
  const onEnd = [
    `// ===== EDIT ME: timings are in ticks of 350 ms =====`,
    `UpdateContext({locTicks: locTicks + 1})`,
    `If(!locInited, UpdateContext({locInited: true, locStep: 0, locTicks: 1}))`,
    `If(locStep = 0 && locTicks >= 4, ${refresh.length ? refresh.join('; ') + '; ' : ''}UpdateContext({locStep: 1}))`,
    `If(locStep = 1 && locTicks >= 6, ${step2.length ? step2.join('; ') + '; ' : ''}UpdateContext({locStep: 2}))`,
    `If(locStep = 2 && locTicks >= 8, UpdateContext({locStep: 3, locClockStop: true}); Navigate(${nextName}, ScreenTransition.Fade))`,
  ].join(';\n');
  add('timer', { name: 'tmrLoad' + T, x: 8, y: 8, w: 60, h: 30, props: { duration: 350, repeat: true, autoStart: true }, actions: { OnTimerEnd: [{ type: 'custom', fx: onEnd }] }, fx: { Start: '!locClockStop' } });
  scr.onVisible = 'UpdateContext({locInited: false, locClockStop: false, locTicks: 0, locStep: 0})';
  return scr;
}

/** Build a complete app for the project's lists (requirements wizard / Claude plan → app). */
export function generateApp(project, { screens = ['home', 'browse', 'detail', 'edit'], approvals = null, chain = null } = {}) {
  project.screens = [];
  project.app.features = features(project);
  const ft = project.app.features;
  project.lists = project.lists.filter((l) => !l.system);
  // an approval chain replaces the single-step approvals for its list
  const chList = chain && project.lists.find((l) => l.name === chain.list);
  const ch = chain ? normalizeChain(chain.values ? chain : { ...chain, values: guessChainValues(chList, chain.field || 'Status') }) : null;
  if (ch && project.lists.some((l) => l.name === ch.list)) { ensureChainColumns(project, ch); project.app.chain = ch; approvals = null; }
  else delete project.app.chain;
  const apList = approvals && project.lists.find((l) => l.name === approvals.list);
  const apCol = apList && apList.columns.find((c) => c.name === (approvals.field || statusCol(apList)?.name) && c.type === 'Choice');
  project.app.approvalsCfg = apCol ? { list: apList.name, field: apCol.name, pending: approvals.pending || apCol.choices[0], yes: approvals.yes || apCol.choices.find((c) => /approv|accept|done/i.test(c)) || apCol.choices[1], no: approvals.no || apCol.choices.find((c) => NEG.test(c)) || apCol.choices[2] } : null;
  const tiles = [];
  const byList = [];
  for (const list of project.lists) {
    const has = (k) => screens.includes(k);
    const s = {};
    if (has('edit')) s.edit = genEdit(project, list);
    if (has('detail')) s.detail = genDetail(project, list, { editScreen: s.edit });
    if (has('browse')) s.browse = genBrowse(project, list, { editScreen: s.edit, detailScreen: s.detail, back: has('home') });
    byList.push([list, s]);
    if (s.browse) tiles.push({ label: pretty(list.name), sub: 'Search and open ' + pretty(list.name).toLowerCase(), icon: 'list', actions: [{ type: 'navigate', screen: s.browse.id }] });
    if (s.edit) tiles.push({ label: 'New ' + lower(list.name), sub: 'Fill in the form', icon: 'plus', actions: [{ type: 'newItem', list: list.name, screen: s.edit.id }] });
  }
  let apScr = null;
  if (project.app.chain) {
    const list = project.lists.find((l) => l.name === project.app.chain.list);
    const det = byList.find(([l]) => l === list)?.[1].detail;
    apScr = genChainApprovals(project, list, project.app.chain, { back: screens.includes('home'), detailScreen: det });
    tiles.unshift({ label: 'Approvals', sub: 'Waiting for your decision', icon: 'check', actions: [{ type: 'navigate', screen: apScr.id }] });
  } else if (project.app.approvalsCfg) {
    const a = project.app.approvalsCfg, list = project.lists.find((l) => l.name === a.list);
    const det = byList.find(([l]) => l === list)?.[1].detail;
    apScr = genApprovals(project, list, { ...a, back: screens.includes('home'), detailScreen: det });
    tiles.unshift({ label: 'Approvals', sub: 'Waiting for a decision', icon: 'check', actions: [{ type: 'navigate', screen: apScr.id }], fx: ft.adminGroup ? { Visible: 'varIsAdmin' } : {} });
  }
  const ordered = [];
  for (const [, s] of byList) ['browse', 'detail', 'edit'].forEach((k) => s[k] && ordered.push(s[k]));
  project.screens = [...ordered, ...project.screens.filter((x) => !ordered.includes(x))];
  let home = null;
  if (screens.includes('home')) {
    const [fl, fs] = byList[0] || [];
    const person = fl && project.app.approvalsCfg?.list === fl.name ? fl.columns.find((c) => c.type === 'Person') : null;
    const scopeVar = fl ? 'varScope' + slug(fl.name) : '';
    home = genHome(project, {
      tiles: tiles.slice(0, 6),
      primary: fs?.edit ? { text: '+ New ' + lower(fl.name), actions: [{ type: 'newItem', list: fl.name, screen: fs.edit.id }] } : null,
      mine: fl && fs?.detail ? { list: fl, title: 'My ' + pretty(fl.name).toLowerCase(), filterField: 'CreatedByMe', detailScreen: fs.detail,
        link: fs.browse ? { text: 'See all', actions: [{ type: 'setVar', name: scopeVar, value: '"mine"' }, { type: 'navigate', screen: fs.browse.id }] } : null,
        emptyTitle: `You have no ${pretty(fl.name).toLowerCase()} yet`, emptyHint: 'New ones you create show up here.',
        emptyButton: fs.edit ? { text: '+ New ' + lower(fl.name), actions: [{ type: 'newItem', list: fl.name, screen: fs.edit.id }] } : null } : null,
      waiting: fl && fs?.detail && project.app.chain?.list === fl.name ? { list: fl, title: 'Waiting on me', chainQueue: true, detailScreen: fs.detail,
        link: apScr ? { text: 'Open queue', actions: [{ type: 'navigate', screen: apScr.id }] } : null, emptyTitle: 'Nothing waiting on you', emptyHint: `${pretty(fl.name)} you need to decide appear here.` }
      : person && fs?.detail ? { list: fl, title: 'Waiting on me', filterField: person.name, extraFilter: `${/^[A-Za-z_]\w*$/.test(project.app.approvalsCfg.field) ? project.app.approvalsCfg.field : `'${project.app.approvalsCfg.field}'`}.Value = "${String(project.app.approvalsCfg.pending).replace(/"/g, '""')}"`,
        detailScreen: fs.detail, link: apScr ? { text: 'Open queue', actions: [{ type: 'navigate', screen: apScr.id }] } : null, emptyTitle: 'Nothing waiting on you', emptyHint: `${pretty(fl.name)} you need to decide appear here.` } : null,
    });
  }
  if (ft.loading) genLoading(project, { next: home || project.screens[0], lists: project.lists });
  if (ft.usage && !project.lists.some((l) => l.name === 'AppUsage')) {
    const u = newList('AppUsage', [newColumn('DisplayName', 'Text'), newColumn('AppName', 'Text'), newColumn('FirstSeen', 'DateTime', { withTime: true }), newColumn('LastSeen', 'DateTime', { withTime: true }), newColumn('VisitCount', 'Number')]);
    u.system = true; u.columns[0].required = false;
    project.lists.push(u);
  }
  project.app.startScreen = project.screens[0]?.id;
  return project;
}

// ---------------------------------------------------------------- first-visit demo
/** A finished Request Tracker with every option and two flows, set up completely, for "Explore a finished demo". */
export function buildDemo(newProject) {
  const p = newProject();
  p.app.name = 'Request Tracker';
  p.app.formFactor = 'tablet';
  p.lists = SCENARIOS.requests.lists();
  p.app.features = { loading: true, usage: true, confirm: true, stepper: true, scope: 'all', adminGroup: '11111111-2222-4333-8444-555555555555' };
  p.app.siteUrl = 'https://contoso.sharepoint.com/sites/Requests';
  p.app.appLink = 'https://apps.powerapps.com/play/e/default-demo/a/00000000-demo?tenantId=00000000-demo';
  p.app.setup = { usageConfirmed: true };
  generateApp(p, { screens: ['home', 'browse', 'detail', 'edit'], approvals: SCENARIOS.requests.approvals });
  p.flows = [newFlow(p, 'created', 'Requests'), newFlow(p, 'statusChanged', 'Requests')];
  const det = p.screens.find((s) => s.name === 'scrDetailRequests');
  if (det) p.app.deepLink = { screen: det.id, list: 'Requests' };
  p.app.demo = true;
  return p;
}

// ---------------------------------------------------------------- insertable blocks (Insert tab)
export const BLOCKS = {
  header: { label: 'Header bar', icon: 'menu' },
  formFields: { label: 'Form fields from a list', icon: 'edit', needsList: true },
  detailFields: { label: 'Detail fields from a list', icon: 'file', needsList: true },
  listGallery: { label: 'Gallery of a list (card rows)', icon: 'list', needsList: true },
  kpi: { label: 'KPI card', icon: 'chart' },
  buttonRow: { label: 'Save + Cancel buttons', icon: 'check', needsList: true },
};
export function insertBlock(project, screen, key, list, at = { x: 40, y: 96 }) {
  const { W, add } = ctx(project, screen), T = tag(screen) + uid().slice(0, 3);
  const before = screen.controls.length;
  if (key === 'header') headerBlock(project, screen, { title: pretty(screen.name.replace(/^scr/, '')) });
  if (key === 'kpi') {
    add('card', { name: 'cardKpi' + T, x: at.x, y: at.y, w: 260, h: 104, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
    add('label', { name: 'lblKpiNum' + T, x: at.x + 20, y: at.y + 14, w: 220, h: 48, props: { text: '0', size: 26, weight: 'Bold', color: '@Primary' }, bind: list ? { kind: 'count', list: list.name, field: '', value: '' } : undefined });
    add('label', { name: 'lblKpiCap' + T, x: at.x + 20, y: at.y + 62, w: 220, h: 26, props: { text: list ? 'Total ' + pretty(list.name).toLowerCase() : 'Metric', size: 11, weight: 'Semibold', color: '@TextMuted' } });
  }
  if (key === 'formFields' && list) {
    let y = at.y;
    for (const col of list.columns.filter((c) => !c.chain)) { const r = inputFor(project, screen, list, col, at.x, y, Math.min(420, W - at.x - 40)); y = r.bottom + 16; }
  }
  if (key === 'detailFields' && list) {
    let y = at.y;
    for (const col of list.columns.filter((c) => !c.chain)) {
      add('label', { name: 'lblCap' + slug(col.name) + T, x: at.x, y, w: 180, h: 28, props: { text: pretty(col.name), size: 11, weight: 'Semibold', color: '@TextMuted' } });
      add('label', { name: 'lblVal' + slug(col.name) + T, x: at.x + 190, y, w: 360, h: 28, props: { text: '-', size: 12, color: '@Text' }, bind: { src: 'sel', list: list.name, field: col.name } });
      y += 36;
    }
  }
  if (key === 'listGallery' && list) {
    const gal = add('gallery', { name: 'galList' + T, x: at.x, y: at.y, w: Math.min(640, W - at.x - 40), h: 400, props: { templateSize: 84 },
      data: { list: list.name, searchCtl: '', searchField: list.columns[0].name, sortField: 'Modified', sortDir: 'Descending', filterField: '', filterValue: '' },
      actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: '', fresh: true }] } });
    galleryRows(project, screen, gal, list);
  }
  if (key === 'buttonRow' && list) {
    add('button', { name: 'btnSave' + T, x: at.x + 156, y: at.y, w: 140, h: 40, props: { text: 'Save' }, actions: { OnSelect: [{ type: 'save', list: list.name, then: 'back' }] } });
    add('button', { name: 'btnCancel' + T, x: at.x, y: at.y, w: 140, h: 40, props: { text: 'Cancel', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1 }, actions: { OnSelect: [{ type: 'back' }] } });
  }
  return screen.controls.slice(before);
}
