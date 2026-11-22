// templates.js — requirement scenarios, whole-screen templates and insertable blocks.
import { newControl, newScreen, newList, newColumn, FORM_FACTORS, allNames, slug, varSel, uid } from './core.js';

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
  blank: { label: 'Start from my own lists', icon: 'grid', blurb: 'Design your lists first, then generate screens for them.', app: 'My App', lists: () => [newList('Items', [C('Status', 'Choice', { choices: ['Open', 'Closed'] })])] },
};

// ---------------------------------------------------------------- helpers
function ctx(project, screen) {
  const ff = FORM_FACTORS[project.app.formFactor || 'tablet'];
  const used = allNames(project);
  const nm = (base) => { let n = base, i = 2; while (used.has(n)) n = base + i++; used.add(n); return n; };
  const add = (type, over, into) => { const c = newControl(project, type, { ...over, name: nm(over.name || type) }); (into || screen.controls).push(c); return c; };
  return { W: ff.w, H: ff.h, phone: project.app.formFactor === 'phone', nm, add };
}
const statusCol = (list) => list.columns.find((c) => c.type === 'Choice' && /status/i.test(c.name)) || list.columns.find((c) => c.type === 'Choice');
const subtitleCol = (list, skip) => list.columns.find((c) => !c.builtin && c !== skip && ['Text', 'Person', 'DateTime', 'Lookup'].includes(c.type));
const tag = (screen) => screen.name.replace(/^scr/, '');
const pretty = (n) => String(n).replace(/([a-z])([A-Z])/g, '$1 $2');
const singular = (n) => pretty(n).replace(/ies$/, 'y').replace(/s$/, '');

export function headerBlock(project, screen, { title, titleFx, back = false, button = null, y = 0 } = {}) {
  const { W, add } = ctx(project, screen), T = tag(screen), H = 72, t = project.theme;
  const out = [];
  out.push(t.header === 'gradient'
    ? add('card', { name: 'cardHeader' + T, x: 0, y, w: W, h: H, props: { fill: '@HeaderBg', fill2: '@Primary', gradAngle: 100, borderThickness: 0, radius: 0, shadow: false } })
    : add('rect', { name: 'rectHeader' + T, x: 0, y, w: W, h: H, props: { fill: '@HeaderBg', borderThickness: 0 } }));
  let x = 24;
  if (back) { out.push(add('icon', { name: 'icoBack' + T, x: 20, y: y + 20, w: 32, h: 32, props: { icon: 'back', color: '@OnHeader', stroke: 2.2 }, actions: { OnSelect: [{ type: 'back' }] } })); x = 64; }
  const lbl = add('label', { name: 'lblTitle' + T, x, y: y + 14, w: W - x - (button ? 220 : 24), h: 44, props: { text: title || pretty(screen.name.replace(/^scr/, '')), size: 18, weight: 'Semibold', color: '@OnHeader' } });
  if (titleFx) lbl.fx.Text = titleFx;
  out.push(lbl);
  if (button) out.push(add('button', { name: 'btnHeader' + T, x: W - 184, y: y + 16, w: 160, h: 40, props: { text: button.text, fill: '@Surface', color: '@Primary', borderThickness: 0, size: 12 }, actions: { OnSelect: button.actions } }));
  return out;
}

function inputFor(project, screen, list, col, x, y, w) {
  const { add } = ctx(project, screen), T = tag(screen);
  const type = col.type === 'YesNo' ? 'toggle' : col.type === 'Person' ? 'person' : col.type === 'DateTime' ? 'datepicker' : (col.type === 'Choice' || col.type === 'Lookup') ? 'dropdown' : 'textinput';
  const prefix = { textinput: 'txt', dropdown: 'dd', datepicker: 'dtp', toggle: 'tgl', person: 'pp' }[type];
  const h = col.type === 'Note' ? 96 : type === 'toggle' ? 32 : 40;
  const lab = add('label', { name: 'lbl' + slug(col.name) + T, x, y, w, h: 24, props: { text: pretty(col.name) + (col.required ? ' *' : ''), size: 11, weight: 'Semibold', color: '@TextMuted' } });
  const inp = add(type, { name: prefix + slug(col.name) + T, x, y: y + 26, w: type === 'toggle' ? 80 : w, h, field: { list: list.name, column: col.name },
    props: type === 'textinput' ? { hint: col.type === 'Currency' ? '0.00' : col.type === 'Number' ? '0' : 'Enter ' + pretty(col.name).toLowerCase(), defaultText: '', multiline: col.type === 'Note', numeric: col.type === 'Number' || col.type === 'Currency', size: 12 } : undefined });
  return { lab, inp, bottom: y + 26 + h };
}

// ---------------------------------------------------------------- screen templates
export const SCREEN_TEMPLATES = {
  blank: { label: 'Blank', needsList: false },
  home: { label: 'Home dashboard', needsList: false },
  browse: { label: 'Browse list (search + gallery)', needsList: true },
  detail: { label: 'Item details', needsList: true },
  edit: { label: 'New / edit form', needsList: true },
  approvals: { label: 'Approvals queue', needsList: true },
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
  const { W, H, phone, add } = ctx(project, scr), T = tag(scr), V = varSel(list.name);
  headerBlock(project, scr, { title: 'Edit ' + singular(list.name).toLowerCase(), titleFx: `If(IsBlank(${V}.ID), "New ${singular(list.name).toLowerCase()}", "Edit ${singular(list.name).toLowerCase()}")`, back: true });
  // short fields first, multi-line notes last at full width
  const cols = [...list.columns.filter((c) => c.type !== 'Note'), ...list.columns.filter((c) => c.type === 'Note')];
  const nNotes = cols.filter((c) => c.type === 'Note').length, nShort = cols.length - nNotes;
  const cardY = 96, FIELD = 84, NOTE = 140;
  const tallness = (n) => cardY + 24 + Math.ceil(nShort / n) * FIELD + nNotes * NOTE + 28 + 40 + 24;
  // tablet: two columns, or three when two would run off the bottom of the screen
  const colsN = phone ? 1 : tallness(2) > H - 8 ? 3 : 2;
  const pad = phone ? 24 : 48, gap = 32;
  const cw = Math.floor((W - pad * 2 - gap * (colsN - 1)) / colsN);
  const card = add('card', { name: 'cardForm' + T, x: pad - 20, y: cardY, w: W - (pad - 20) * 2, h: 200, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  const ys = Array(colsN).fill(cardY + 24);
  let colI = 0, rowBottom = cardY + 24;
  for (const col of cols) {
    const wide = col.type === 'Note' && colsN > 1;
    if (wide) { colI = 0; ys.fill(Math.max(...ys)); }
    const r = inputFor(project, scr, list, col, pad + colI * (cw + gap), ys[colI], wide ? W - pad * 2 : cw);
    if (wide) ys.fill(r.bottom + 18);
    else { ys[colI] = r.bottom + 18; colI = (colI + 1) % colsN; }
    rowBottom = Math.max(rowBottom, r.bottom);
  }
  const by = rowBottom + 28;
  card.h = by + 40 + 24 - cardY;
  add('button', { name: 'btnSave' + T, x: W - pad - 160, y: by, w: 160, h: 40, props: { text: 'Save' }, actions: { OnSelect: [{ type: 'save', list: list.name, then: 'back' }] } });
  add('button', { name: 'btnCancel' + T, x: W - pad - 160 - 16 - 140, y: by, w: 140, h: 40, props: { text: 'Cancel', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1 }, actions: { OnSelect: [{ type: 'back' }] } });
  return scr;
}

export function genDetail(project, list, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrDetail' + slug(list.name));
  project.screens.push(scr);
  const { W, phone, add } = ctx(project, scr), T = tag(scr), V = varSel(list.name);
  const titleCol = list.columns[0];
  headerBlock(project, scr, { title: singular(list.name), titleFx: `${V}.${titleCol.name}`, back: true });
  const pad = phone ? 24 : 48, colsN = phone ? 1 : 3;
  const cols = [...list.columns.filter((c) => !c.builtin && c.type !== 'Note'), ...list.columns.filter((c) => !c.builtin && c.type === 'Note')];
  const cw = Math.floor((W - pad * 2 - 24 * (colsN - 1)) / colsN);
  const st = statusCol(list);
  const card = add('card', { name: 'cardDetail' + T, x: pad - 20, y: 96, w: W - (pad - 20) * 2, h: 200, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  add('label', { name: 'lblBig' + T, x: pad, y: 116, w: W - pad * 2 - 140, h: 40, props: { text: 'Title', size: 20, weight: 'Semibold', color: '@Text' }, bind: { src: 'sel', list: list.name, field: titleCol.name } });
  if (st) add('badge', { name: 'pillStatus' + T, x: W - pad - 120, y: 122, w: 120, h: 28, bind: { src: 'sel', list: list.name, field: st.name }, props: { colorMap: badgeMap(st) } });
  let y = 176, i = 0, bottom = y;
  for (const col of cols) {
    if (col === st) continue;
    const x = pad + (i % colsN) * (cw + 24);
    const h = col.type === 'Note' ? 72 : 30;
    if (col.type === 'Note' && colsN > 1) {
      if (i % colsN) { y = bottom + 14; i = 0; }
      add('label', { name: 'lblCap' + slug(col.name) + T, x: pad, y, w: W - pad * 2, h: 22, props: { text: pretty(col.name), size: 10, weight: 'Semibold', color: '@TextMuted' } });
      add('label', { name: 'lblVal' + slug(col.name) + T, x: pad, y: y + 22, w: W - pad * 2, h, props: { text: '-', size: 12, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: list.name, field: col.name } });
      bottom = y + 22 + h; y = bottom + 14; i = 0; continue;
    }
    add('label', { name: 'lblCap' + slug(col.name) + T, x, y, w: cw, h: 22, props: { text: pretty(col.name), size: 10, weight: 'Semibold', color: '@TextMuted' } });
    add('label', { name: 'lblVal' + slug(col.name) + T, x, y: y + 22, w: cw, h, props: { text: '-', size: 13, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: list.name, field: col.name } });
    bottom = Math.max(bottom, y + 22 + h);
    i++;
    if (i % colsN === 0) y = bottom + 14;
  }
  const by = bottom + 28;
  card.h = by + 64 - 96;
  const editScr = opts.editScreen;
  if (editScr) add('button', { name: 'btnEdit' + T, x: W - pad - 140, y: by, w: 140, h: 40, props: { text: 'Edit' }, actions: { OnSelect: [{ type: 'navigate', screen: editScr.id }] } });
  add('button', { name: 'btnDelete' + T, x: W - pad - 140 - (editScr ? 156 : 0), y: by, w: 140, h: 40, props: { text: 'Delete', fill: 'transparent', color: '@Danger', borderColor: '@Danger', borderThickness: 1 }, actions: { OnSelect: [{ type: 'delete', list: list.name, then: 'back' }] } });
  return scr;
}
const badgeMap = (st) => (st?.choices || []).map((c, i) => `${c}=${/approv|done|complete|closed|registered|good|new/i.test(c) ? '@Success' : /reject|cancel|block|retired|urgent/i.test(c) ? '@Danger' : /progress|pending|submit|wait|medium|fair/i.test(c) ? '@Warning' : i === 0 ? '@Primary' : '@TextMuted'}`).join('\n');

export function galleryRows(project, scr, gal, list) {
  const { add } = ctx(project, scr), T = tag(scr);
  const st = statusCol(list), sub = subtitleCol(list, st), rowH = gal.props.templateSize, W = gal.w;
  const into = gal.children;
  add('label', { name: 'lblRowTitle' + T, x: 20, y: 14, w: W - (st ? 220 : 80), h: 28, props: { text: 'Title', size: 13, weight: 'Semibold', color: '@Text' }, bind: { src: 'item', list: list.name, field: list.columns[0].name } }, into);
  if (sub) add('label', { name: 'lblRowSub' + T, x: 20, y: 42, w: W - (st ? 220 : 80), h: 24, props: { text: 'Subtitle', size: 11, color: '@TextMuted' }, bind: { src: 'item', list: list.name, field: sub.name } }, into);
  if (st) add('badge', { name: 'pillRow' + T, x: W - 172, y: Math.round(rowH / 2 - 14), w: 116, h: 28, bind: { src: 'item', list: list.name, field: st.name }, props: { colorMap: badgeMap(st) } }, into);
  add('icon', { name: 'icoRowGo' + T, x: W - 44, y: Math.round(rowH / 2 - 12), w: 24, h: 24, props: { icon: 'chevron', color: '@TextMuted', stroke: 2 } }, into);
  add('rect', { name: 'rectRowLine' + T, x: 0, y: rowH - 1, w: W, h: 1, props: { fill: '@Border', borderThickness: 0 } }, into);
}

export function genBrowse(project, list, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrBrowse' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add } = ctx(project, scr), T = tag(scr);
  headerBlock(project, scr, { title: pretty(list.name), back: !!opts.back, button: opts.editScreen ? { text: '+ New', actions: [{ type: 'newItem', list: list.name, screen: opts.editScreen.id }] } : null });
  const pad = phone ? 16 : 40;
  const search = add('textinput', { name: 'txtSearch' + T, x: pad, y: 96, w: phone ? W - pad * 2 : 420, h: 40, props: { hint: 'Search ' + pretty(list.name).toLowerCase() + ' by title', defaultText: '', size: 12 } });
  add('icon', { name: 'icoSearch' + T, x: pad + (phone ? W - pad * 2 : 420) - 34, y: 104, w: 24, h: 24, props: { icon: 'search', color: '@TextMuted', stroke: 2 } });
  add('card', { name: 'cardList' + T, x: pad - 6, y: 150, w: W - (pad - 6) * 2, h: H - 150 - 20, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  const gal = add('gallery', { name: 'galList' + T, x: pad, y: 158, w: W - pad * 2, h: H - 158 - 30, props: { templateSize: 76, templateFill: 'transparent', selectedFill: '@SurfaceAlt' },
    data: { list: list.name, searchCtl: search.name, searchField: list.columns[0].name, sortField: 'Modified', sortDir: 'Descending', filterField: '', filterValue: '' },
    actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: opts.detailScreen?.id || '' }] } });
  galleryRows(project, scr, gal, list);
  return scr;
}

export function genApprovals(project, list, cfg = {}) {
  const st = cfg.field ? list.columns.find((c) => c.name === cfg.field) : statusCol(list);
  const scr = newScreen(project, cfg.name || 'scrApprovals' + slug(list.name));
  project.screens.push(scr);
  const { W, H, phone, add } = ctx(project, scr), T = tag(scr), V = varSel(list.name);
  headerBlock(project, scr, { title: 'Approvals', back: !!cfg.back });
  const pending = cfg.pending || st?.choices?.[0] || '';
  const yes = cfg.yes || st?.choices?.find((c) => /approv|done|complete/i.test(c)) || st?.choices?.[1] || '';
  const no = cfg.no || st?.choices?.find((c) => /reject|cancel|declin/i.test(c)) || st?.choices?.[2] || '';
  const listW = phone ? W - 32 : 520;
  add('label', { name: 'lblQueue' + T, x: 24, y: 88, w: listW, h: 32, props: { text: `Waiting for a decision (${pending})`, size: 12, weight: 'Semibold', color: '@TextMuted' } });
  add('card', { name: 'cardQueue' + T, x: 18, y: 124, w: listW + 12, h: H - 124 - 18, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  const gal = add('gallery', { name: 'galQueue' + T, x: 24, y: 130, w: listW, h: H - 130 - 30, props: { templateSize: 76, templateFill: 'transparent', selectedFill: '@SurfaceAlt' },
    data: { list: list.name, searchCtl: '', searchField: '', filterField: st?.name || '', filterValue: pending, sortField: 'Created', sortDir: 'Ascending' },
    actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: '' }] } });
  galleryRows(project, scr, gal, list);
  if (phone) return scr;
  const px = listW + 56, pw = W - px - 24;
  const panelVis = { Visible: `!IsBlank(${V}.ID)` };
  add('card', { name: 'cardPanel' + T, x: px - 20, y: 124, w: pw + 20, h: H - 124 - 18, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
  add('label', { name: 'lblPick' + T, x: px, y: 300, w: pw - 20, h: 40, props: { text: 'Select an item to review it', size: 13, color: '@TextMuted', align: 'Center' }, fx: { Visible: `IsBlank(${V}.ID)` } });
  add('label', { name: 'lblPanelTitle' + T, x: px, y: 140, w: pw - 20, h: 40, props: { text: 'Title', size: 18, weight: 'Semibold', color: '@Text' }, bind: { src: 'sel', list: list.name, field: list.columns[0].name }, fx: { ...panelVis } });
  let y = 192;
  const pcols = list.columns.filter((c) => !c.builtin && c !== st);
  for (const col of [...pcols.filter((c) => c.type !== 'Note'), ...pcols.filter((c) => c.type === 'Note')].slice(0, 7)) {
    add('label', { name: 'lblPCap' + slug(col.name) + T, x: px, y, w: 180, h: 28, props: { text: pretty(col.name), size: 11, weight: 'Semibold', color: '@TextMuted' }, fx: { ...panelVis } });
    add('label', { name: 'lblPVal' + slug(col.name) + T, x: px + 190, y, w: pw - 210, h: col.type === 'Note' ? 56 : 28, props: { text: '-', size: 12, color: '@Text', valign: 'Top' }, bind: { src: 'sel', list: list.name, field: col.name }, fx: { ...panelVis } });
    y += col.type === 'Note' ? 64 : 36;
  }
  if (st) {
    const by = Math.max(y + 24, H - 100);
    add('button', { name: 'btnApprove' + T, x: px, y: by, w: 160, h: 44, props: { text: 'Approve', fill: '@Success', borderColor: '@Success' }, fx: { ...panelVis }, actions: { OnSelect: [{ type: 'setField', list: list.name, field: st.name, value: yes, message: 'Approved' }, { type: 'custom', fx: `Set(${V}, Defaults(${list.name.match(/^[A-Za-z_]\w*$/) ? list.name : "'" + list.name + "'"}))` }] } });
    add('button', { name: 'btnReject' + T, x: px + 176, y: by, w: 160, h: 44, props: { text: 'Reject', fill: '@Danger', borderColor: '@Danger' }, fx: { ...panelVis }, actions: { OnSelect: [{ type: 'setField', list: list.name, field: st.name, value: no, message: 'Rejected' }, { type: 'custom', fx: `Set(${V}, Defaults(${list.name.match(/^[A-Za-z_]\w*$/) ? list.name : "'" + list.name + "'"}))` }] } });
  }
  return scr;
}

export function genHome(project, opts = {}) {
  const scr = newScreen(project, opts.name || 'scrHome');
  project.screens.unshift(scr);
  const { W, phone, add } = ctx(project, scr), T = tag(scr);
  headerBlock(project, scr, { title: project.app.name });
  const pad = phone ? 24 : 48;
  add('label', { name: 'lblWelcome' + T, x: pad, y: 96, w: W - pad * 2, h: 40, props: { text: 'Welcome', size: 20, weight: 'Semibold', color: '@Text' }, bind: { kind: 'user', field: 'FullName', prefix: 'Welcome, ' } });
  add('label', { name: 'lblIntro' + T, x: pad, y: 136, w: W - pad * 2, h: 28, props: { text: 'Pick where you want to go.', size: 12, color: '@TextMuted' } });
  let y = 184;
  // KPI row from the first list with a status-like choice column
  const kl = project.lists.find((l) => statusCol(l));
  if (kl) {
    const st = statusCol(kl), vals = st.choices.slice(0, phone ? 2 : 4);
    const n = vals.length, gap = 20, w = Math.floor((W - pad * 2 - gap * (n - 1)) / n);
    vals.forEach((v, i) => {
      const x = pad + i * (w + gap);
      add('card', { name: 'cardKpi' + slug(v) + T, x, y, w, h: 104, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
      add('label', { name: 'lblKpiNum' + slug(v) + T, x: x + 20, y: y + 14, w: w - 40, h: 48, props: { text: '0', size: 26, weight: 'Bold', color: '@Primary' }, bind: { kind: 'count', list: kl.name, field: st.name, value: v } });
      add('label', { name: 'lblKpiCap' + slug(v) + T, x: x + 20, y: y + 62, w: w - 40, h: 26, props: { text: v, size: 11, weight: 'Semibold', color: '@TextMuted' } });
    });
    y += 128;
  }
  // navigation tiles
  const tiles = opts.tiles || [];
  const per = phone ? 1 : Math.min(3, Math.max(1, tiles.length)), gap = 20;
  const tw = Math.floor((W - pad * 2 - gap * (per - 1)) / per), th = phone ? 96 : 132;
  tiles.forEach((t, i) => {
    const x = pad + (i % per) * (tw + gap), ty = y + Math.floor(i / per) * (th + gap), k = slug(t.label) + T;
    add('card', { name: 'cardTile' + k, x, y: ty, w: tw, h: th, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
    add('card', { name: 'cardTileIco' + k, x: x + 20, y: ty + 22, w: 48, h: 48, props: { fill: '@Primary', fill2: '@Accent', gradAngle: 135, borderThickness: 0, radius: 12, shadow: false } });
    add('icon', { name: 'icoTile' + k, x: x + 32, y: ty + 34, w: 24, h: 24, props: { icon: t.icon || 'list', color: '#FFFFFF', stroke: 2 } });
    add('label', { name: 'lblTile' + k, x: x + 84, y: ty + 20, w: tw - 104, h: 30, props: { text: t.label, size: 14, weight: 'Semibold', color: '@Text' } });
    add('label', { name: 'lblTileSub' + k, x: x + 84, y: ty + 50, w: tw - 104, h: phone ? 30 : 56, props: { text: t.sub || '', size: 11, color: '@TextMuted', valign: 'Top' } });
    // transparent button last, so it sits in front of the labels and gets the click
    add('button', { name: 'btnTile' + k, x, y: ty, w: tw, h: th, props: { text: '', fill: 'transparent', color: 'transparent', borderColor: 'transparent', borderThickness: 0, radius: 'theme' }, actions: { OnSelect: t.actions } });
  });
  // recent items fill the rest of the screen
  const rows = Math.ceil(tiles.length / per), ry = y + rows * (th + gap) + (tiles.length ? 8 : 0);
  const { H } = ctx(project, scr);
  const rl = opts.recent && project.lists.find((l) => l.name === opts.recent.list);
  if (rl && H - ry > 200) {
    add('label', { name: 'lblRecent' + T, x: pad, y: ry, w: W - pad * 2, h: 32, props: { text: 'Recent ' + pretty(rl.name).toLowerCase(), size: 13, weight: 'Semibold', color: '@Text' } });
    add('card', { name: 'cardRecent' + T, x: pad - 6, y: ry + 38, w: W - (pad - 6) * 2, h: H - ry - 38 - 20, props: { fill: '@Surface', borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } });
    const gal = add('gallery', { name: 'galRecent' + T, x: pad, y: ry + 46, w: W - pad * 2, h: H - ry - 46 - 32, props: { templateSize: 76, templateFill: 'transparent', selectedFill: '@SurfaceAlt' },
      data: { list: rl.name, searchCtl: '', searchField: rl.columns[0].name, sortField: 'Modified', sortDir: 'Descending', filterField: '', filterValue: '' },
      actions: { OnSelect: [{ type: 'openItem', list: rl.name, screen: opts.recent.detailScreen?.id || '' }] } });
    galleryRows(project, scr, gal, rl);
  }
  return scr;
}

/** Build a complete app for the project's lists (requirements wizard → app). */
export function generateApp(project, { screens = ['home', 'browse', 'detail', 'edit'], approvals = null } = {}) {
  project.screens = [];
  const tiles = [];
  const byList = [];
  for (const list of project.lists) {
    const has = (k) => screens.includes(k);
    const s = {};
    if (has('edit')) s.edit = genEdit(project, list);
    if (has('detail')) s.detail = genDetail(project, list, { editScreen: s.edit });
    if (has('browse')) s.browse = genBrowse(project, list, { editScreen: s.edit, detailScreen: s.detail, back: has('home') });
    byList.push([list, s]);
    if (s.browse) tiles.push({ label: pretty(list.name), sub: 'Browse, search and open ' + pretty(list.name).toLowerCase(), icon: 'list', actions: [{ type: 'navigate', screen: s.browse.id }] });
    if (s.edit) tiles.push({ label: 'New ' + singular(list.name).toLowerCase(), sub: 'Fill in the form to add one', icon: 'plus', actions: [{ type: 'newItem', list: list.name, screen: s.edit.id }] });
  }
  if (approvals && project.lists.find((l) => l.name === approvals.list)) {
    const list = project.lists.find((l) => l.name === approvals.list);
    const a = genApprovals(project, list, { ...approvals, back: screens.includes('home') });
    tiles.push({ label: 'Approvals', sub: 'Review items waiting for a decision', icon: 'check', actions: [{ type: 'navigate', screen: a.id }] });
  }
  // screens order: browse → detail → edit per list
  const ordered = [];
  for (const [, s] of byList) ['browse', 'detail', 'edit'].forEach((k) => s[k] && ordered.push(s[k]));
  project.screens = [...ordered, ...project.screens.filter((x) => !ordered.includes(x))];
  const first = byList[0];
  if (screens.includes('home')) genHome(project, { tiles: tiles.slice(0, 6), recent: first && first[1].detail ? { list: first[0].name, detailScreen: first[1].detail } : null });
  project.app.startScreen = project.screens[0]?.id;
  return project;
}

// ---------------------------------------------------------------- insertable blocks (Insert tab)
export const BLOCKS = {
  header: { label: 'Header bar', icon: 'menu' },
  formFields: { label: 'Form fields from a list', icon: 'edit', needsList: true },
  detailFields: { label: 'Detail fields from a list', icon: 'file', needsList: true },
  listGallery: { label: 'Gallery of a list', icon: 'list', needsList: true },
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
    for (const col of list.columns) { const r = inputFor(project, screen, list, col, at.x, y, Math.min(420, W - at.x - 40)); y = r.bottom + 16; }
  }
  if (key === 'detailFields' && list) {
    let y = at.y;
    for (const col of list.columns) {
      add('label', { name: 'lblCap' + slug(col.name) + T, x: at.x, y, w: 180, h: 28, props: { text: pretty(col.name), size: 11, weight: 'Semibold', color: '@TextMuted' } });
      add('label', { name: 'lblVal' + slug(col.name) + T, x: at.x + 190, y, w: 360, h: 28, props: { text: '-', size: 12, color: '@Text' }, bind: { src: 'sel', list: list.name, field: col.name } });
      y += 36;
    }
  }
  if (key === 'listGallery' && list) {
    const gal = add('gallery', { name: 'galList' + T, x: at.x, y: at.y, w: Math.min(640, W - at.x - 40), h: 400, props: { templateSize: 76, templateFill: '@Surface', selectedFill: '@SurfaceAlt' },
      data: { list: list.name, searchCtl: '', searchField: list.columns[0].name, sortField: 'Modified', sortDir: 'Descending', filterField: '', filterValue: '' },
      actions: { OnSelect: [{ type: 'openItem', list: list.name, screen: '' }] } });
    galleryRows(project, screen, gal, list);
  }
  if (key === 'buttonRow' && list) {
    add('button', { name: 'btnSave' + T, x: at.x + 156, y: at.y, w: 140, h: 40, props: { text: 'Save' }, actions: { OnSelect: [{ type: 'save', list: list.name, then: 'back' }] } });
    add('button', { name: 'btnCancel' + T, x: at.x, y: at.y, w: 140, h: 40, props: { text: 'Cancel', fill: 'transparent', color: '@Primary', borderColor: '@Primary', borderThickness: 1 }, actions: { OnSelect: [{ type: 'back' }] } });
  }
  return screen.controls.slice(before);
}
