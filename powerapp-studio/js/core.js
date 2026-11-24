// core.js — project model, control catalog, themes, column types, sample data.
// Everything here is plain data + small helpers; no DOM.

/**
 * @typedef {'@Primary'|'@OnPrimary'|string} ColorValue  '@Token' (theme), '#RRGGBB[AA]' or 'transparent'
 * @typedef {{type: string, screen?: string, list?: string, field?: string, value?: string, then?: string,
 *   name?: string, fx?: string, text?: string, kind?: string, url?: string, message?: string, transition?: string}} Action
 * @typedef {{id: string, type: string, name: string, x: number, y: number, w: number, h: number,
 *   props: Object<string, any>, fx: Object<string, string>, actions: Object<string, Action[]>,
 *   field?: {list: string, column: string}, bind?: Object, visibleWhen?: {var: string, op: string, value?: string},
 *   data?: Object, children?: Control[], paControl?: string, variant?: string, rawProps?: Object, rawChildren?: any[]}} Control
 * @typedef {{id: string, name: string, fill: ColorValue, fillFx?: string, onVisible: string, controls: Control[]}} Screen
 * @typedef {{id: string, name: string, type: string, required: boolean, choices: string[], lookupList: string, builtin?: boolean}} Column
 * @typedef {{id: string, name: string, columns: Column[]}} List
 * @typedef {{version: number, app: {name: string, formFactor: string, startScreen: string, onStart: string, formulas: string},
 *   theme: {name: string, colors: Object<string, string>, font: string, radius: number, header: string, dark: boolean},
 *   lists: List[], screens: Screen[]}} Project
 */

export const uid = () => Math.random().toString(36).slice(2, 10);
export const clone = (o) => JSON.parse(JSON.stringify(o));
export const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

export const FORM_FACTORS = {
  tablet: { label: 'Tablet 1366 × 768', w: 1366, h: 768 },
  phone: { label: 'Phone 640 × 1136', w: 640, h: 1136 },
};

// ---------- names ----------
export const IDENT = /^[A-Za-z_][A-Za-z0-9_]*$/;
export const RESERVED = new Set(('App Self Parent ThisItem ThisRecord If Switch Set Navigate Back Patch Filter LookUp ' +
  'Sort SortByColumns Text Value User Today Now true false Blank IsBlank Collect ClearCollect Clear Remove ' +
  'Notify Launch Reset Select Refresh Concat CountRows First Last Color RGBA Defaults Choices Table Screen').split(' '));

/** Quote a SharePoint list/column name for Power Fx when it is not a plain identifier. */
export const q = (n) => (IDENT.test(n) ? n : "'" + String(n).replace(/'/g, "''") + "'");
/** Power Fx string literal. */
export const s = (t) => '"' + String(t ?? '').replace(/"/g, '""') + '"';
/** Turn any text into an identifier fragment. */
export const slug = (t) => {
  const x = String(t || '').replace(/[^A-Za-z0-9]+(.)?/g, (m, c) => (c ? c.toUpperCase() : '')).replace(/^[^A-Za-z]+/, '');
  return x ? x[0].toUpperCase() + x.slice(1) : 'Item';
};
export const varSel = (listName) => 'varSel' + slug(listName);

// ---------- SharePoint column types ----------
export const COLTYPES = {
  Text: { label: 'Single line of text', input: 'textinput' },
  Note: { label: 'Multiple lines of text', input: 'textinput' },
  Number: { label: 'Number', input: 'textinput' },
  Currency: { label: 'Currency', input: 'textinput' },
  DateTime: { label: 'Date', input: 'datepicker' },
  Choice: { label: 'Choice', input: 'dropdown' },
  YesNo: { label: 'Yes/No', input: 'toggle' },
  Person: { label: 'Person', input: 'person' },
  Lookup: { label: 'Lookup', input: 'dropdown' },
};
export const SP_RESERVED_COLS = new Set(['ID', 'Created', 'Modified', 'Author', 'Editor', 'StartDate', 'Location',
  'EndDate', 'Status0', 'Attachments', 'ContentType', 'Order', 'Path', 'Type', 'Name', 'Description0']);

export const newColumn = (name = 'NewColumn', type = 'Text', extra = {}) =>
  ({ id: uid(), name, type, required: false, choices: type === 'Choice' ? ['Option A', 'Option B'] : [], lookupList: '', ...extra });

export const newList = (name = 'NewList', cols = []) => ({
  id: uid(), name,
  columns: [{ id: uid(), name: 'Title', type: 'Text', required: true, choices: [], lookupList: '', builtin: true }, ...cols],
});

// ---------- theme ----------
export const THEME_COLORS = [
  ['Primary', 'Buttons, links, focus ring'],
  ['OnPrimary', 'Text on primary'],
  ['Accent', 'Highlights, gradient end'],
  ['Bg', 'Screen background'],
  ['Surface', 'Cards and panels'],
  ['SurfaceAlt', 'Alternate rows, raised areas'],
  ['Border', 'Lines and card borders'],
  ['Text', 'Main text'],
  ['TextMuted', 'Secondary text'],
  ['InputFill', 'Input background'],
  ['InputBorder', 'Input border'],
  ['HeaderBg', 'Header bar'],
  ['OnHeader', 'Text on header'],
  ['Success', 'Approved / done'],
  ['Warning', 'Pending / attention'],
  ['Danger', 'Rejected / delete'],
];
export const FONTS = ['Segoe UI', 'Open Sans', 'Lato', 'Arial', 'Verdana', 'Georgia', 'Courier New'];

const T = (name, c, o = {}) => ({ name, colors: c, font: 'Segoe UI', radius: 8, header: 'gradient', dark: false, ...o });
export const THEME_PRESETS = [
  T('SaTech Cyan', { Primary: '#0891B2', OnPrimary: '#FFFFFF', Accent: '#22D3EE', Bg: '#F1F5F9', Surface: '#FFFFFF', SurfaceAlt: '#F8FAFC', Border: '#E2E8F0', Text: '#0F172A', TextMuted: '#64748B', InputFill: '#FFFFFF', InputBorder: '#CBD5E1', HeaderBg: '#0B1220', OnHeader: '#FFFFFF', Success: '#10B981', Warning: '#F59E0B', Danger: '#DC2626' }),
  T('Fluent Blue', { Primary: '#0F6CBD', OnPrimary: '#FFFFFF', Accent: '#4F9FE8', Bg: '#F5F5F5', Surface: '#FFFFFF', SurfaceAlt: '#FAFAFA', Border: '#E0E0E0', Text: '#242424', TextMuted: '#616161', InputFill: '#FFFFFF', InputBorder: '#C7C7C7', HeaderBg: '#0F6CBD', OnHeader: '#FFFFFF', Success: '#107C10', Warning: '#C47F00', Danger: '#C50F1F' }, { font: 'Segoe UI', radius: 4, header: 'solid' }),
  T('Midnight', { Primary: '#38BDF8', OnPrimary: '#04121F', Accent: '#A78BFA', Bg: '#0B1120', Surface: '#131C2E', SurfaceAlt: '#1B2740', Border: '#26344F', Text: '#E5EDF7', TextMuted: '#94A3B8', InputFill: '#0F1729', InputBorder: '#334463', HeaderBg: '#070C17', OnHeader: '#E5EDF7', Success: '#34D399', Warning: '#FBBF24', Danger: '#F87171' }, { dark: true, radius: 10 }),
  T('Federal Navy', { Primary: '#1C3966', OnPrimary: '#FFFFFF', Accent: '#B22234', Bg: '#F3F5F8', Surface: '#FFFFFF', SurfaceAlt: '#EEF2F8', Border: '#D5DCE6', Text: '#14213D', TextMuted: '#5B6B82', InputFill: '#FFFFFF', InputBorder: '#B9C4D4', HeaderBg: '#1C3966', OnHeader: '#FFFFFF', Success: '#2E7D32', Warning: '#B26A00', Danger: '#B22234' }, { radius: 6 }),
  T('Forest', { Primary: '#2F6B4F', OnPrimary: '#FFFFFF', Accent: '#8DBF6A', Bg: '#F3F5F0', Surface: '#FFFFFF', SurfaceAlt: '#F7F9F4', Border: '#DCE3D5', Text: '#1E2A22', TextMuted: '#5E6E62', InputFill: '#FFFFFF', InputBorder: '#C3CEBB', HeaderBg: '#1F3D2E', OnHeader: '#FFFFFF', Success: '#2F855A', Warning: '#B7791F', Danger: '#C53030' }, { font: 'Lato', radius: 12 }),
  T('Sunset', { Primary: '#E4572E', OnPrimary: '#FFFFFF', Accent: '#F3A712', Bg: '#FFF8F3', Surface: '#FFFFFF', SurfaceAlt: '#FFF1E8', Border: '#F2DED0', Text: '#2B1A12', TextMuted: '#7A5B4C', InputFill: '#FFFFFF', InputBorder: '#E7C9B6', HeaderBg: '#2B1A12', OnHeader: '#FFFFFF', Success: '#3C8D5A', Warning: '#D98A00', Danger: '#C8241B' }, { font: 'Open Sans', radius: 14 }),
  T('Graphite', { Primary: '#111827', OnPrimary: '#FFFFFF', Accent: '#6B7280', Bg: '#FFFFFF', Surface: '#FFFFFF', SurfaceAlt: '#F9FAFB', Border: '#E5E7EB', Text: '#111827', TextMuted: '#6B7280', InputFill: '#FFFFFF', InputBorder: '#D1D5DB', HeaderBg: '#FFFFFF', OnHeader: '#111827', Success: '#15803D', Warning: '#A16207', Danger: '#B91C1C' }, { radius: 6, header: 'solid' }),
];
export const defaultTheme = () => clone(THEME_PRESETS[0]);

export const hexToRgba = (hex) => {
  let h = String(hex || '').trim().replace('#', '');
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  const n = (i) => parseInt(h.slice(i, i + 2), 16);
  if (!/^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(h)) return [0, 0, 0, 1];
  return [n(0), n(2), n(4), h.length === 8 ? Math.round((n(6) / 255) * 100) / 100 : 1];
};
export const isHex = (v) => /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(String(v || ''));

/** Resolve a color value ('@Token' | '#hex' | 'transparent') to CSS for the designer. */
export const cssColor = (v, theme) => {
  if (!v || v === 'transparent') return 'transparent';
  if (v[0] === '@') { const c = theme.colors[v.slice(1)]; return isHex(c) ? c : '#ff00ff'; }
  if (isHex(v)) { const [r, g, b, a] = hexToRgba(v); return `rgba(${r},${g},${b},${a})`; }
  const rgba = String(v).replace(/\s+/g, '');
  if (/^rgba\(\d{1,3},\d{1,3},\d{1,3},(?:0|1|0?\.\d+)\)$/.test(rgba)) return rgba; // RGBA()/ColorValue results from Power Fx overrides
  return 'transparent'; // anything else could break out of a style attribute
};
/** A value safe to store as a color: theme token, hex or 'transparent'. */
export const isColorValue = (v) => v === '' || v === 'transparent' || isHex(v) || /^@[A-Za-z]+$/.test(String(v || ''));

// ---------- control catalog ----------
// pa: Power Apps control type written to YAML. Designer-only types expand to several controls on emit.
export const CATALOG = {
  label: { label: 'Label', cat: 'Basic', pa: 'Label', prefix: 'lbl', w: 240, h: 40,
    props: { text: 'Text', size: 13, weight: 'Normal', color: '@Text', fill: 'transparent', align: 'Left', valign: 'Middle', italic: false } },
  button: { label: 'Button', cat: 'Basic', pa: 'Classic/Button', prefix: 'btn', w: 160, h: 40,
    props: { text: 'Button', size: 13, weight: 'Semibold', color: '@OnPrimary', fill: '@Primary', borderColor: '@Primary', borderThickness: 0, radius: 'theme' } },
  badge: { label: 'Status pill', cat: 'Basic', pa: 'Classic/Button', prefix: 'pill', w: 110, h: 28,
    props: { text: 'Pending', size: 10, weight: 'Semibold', color: '#FFFFFF', fill: '@Warning', colorMap: 'Approved=@Success\nRejected=@Danger\nPending=@Warning', radius: 14 } },
  icon: { label: 'Icon', cat: 'Basic', pa: 'Image', prefix: 'ico', w: 32, h: 32,
    props: { icon: 'home', color: '@Primary', stroke: 2 } },
  card: { label: 'Card / shape', cat: 'Basic', pa: 'HtmlViewer', prefix: 'card', w: 320, h: 180,
    props: { fill: '@Surface', fill2: '', gradAngle: 135, borderColor: '@Border', borderThickness: 1, radius: 'theme', shadow: true } },
  rect: { label: 'Rectangle', cat: 'Basic', pa: 'Rectangle', prefix: 'rect', w: 200, h: 120,
    props: { fill: '@Surface', borderColor: '@Border', borderThickness: 0 } },
  image: { label: 'Image', cat: 'Media', pa: 'Image', prefix: 'img', w: 160, h: 120,
    props: { src: '', fit: 'Fit' } },
  html: { label: 'HTML text', cat: 'Media', pa: 'HtmlViewer', prefix: 'html', w: 320, h: 120,
    props: { html: "<b>Bold</b> and <i>italic</i> rich text" } },
  textinput: { label: 'Text input', cat: 'Input', pa: 'Classic/TextInput', prefix: 'txt', w: 300, h: 40,
    props: { hint: 'Type here', defaultText: '', multiline: false, numeric: false, size: 12 } },
  dropdown: { label: 'Dropdown', cat: 'Input', pa: 'Classic/DropDown', prefix: 'dd', w: 300, h: 40,
    props: { items: 'Option A\nOption B\nOption C', size: 12 } },
  datepicker: { label: 'Date picker', cat: 'Input', pa: 'Classic/DatePicker', prefix: 'dtp', w: 300, h: 40,
    props: { size: 12 } },
  checkbox: { label: 'Checkbox', cat: 'Input', pa: 'Classic/CheckBox', prefix: 'chk', w: 260, h: 36,
    props: { text: 'Checkbox', size: 12, defaultOn: false, color: '@Text' } },
  toggle: { label: 'Toggle', cat: 'Input', pa: 'Classic/Toggle', prefix: 'tgl', w: 80, h: 32,
    props: { defaultOn: false } },
  person: { label: 'People picker', cat: 'Input', pa: 'Classic/TextInput', prefix: 'pp', w: 300, h: 40,
    props: { hint: 'Search people', size: 12 } },
  gallery: { label: 'Gallery', cat: 'Data', pa: 'Gallery@2.15.0', prefix: 'gal', w: 600, h: 420,
    props: { templateSize: 84, templateFill: '@Surface', selectedFill: '@SurfaceAlt', divider: true } },
  timer: { label: 'Timer', cat: 'Advanced', pa: 'Timer', prefix: 'tmr', w: 60, h: 30,
    props: { duration: 1000, repeat: false, autoStart: false } },
  raw: { label: 'Imported control', cat: 'Hidden', pa: '', prefix: 'ctl', w: 160, h: 60, props: {} },
};
export const CATS = ['Basic', 'Input', 'Data', 'Media', 'Advanced'];

/** Controls that accept field bindings (and map to SharePoint column types). */
export const INPUT_TYPES = new Set(['textinput', 'dropdown', 'datepicker', 'checkbox', 'toggle', 'person']);

export function newControl(project, type, over = {}) {
  const c = CATALOG[type];
  over = Object.fromEntries(Object.entries(over).filter(([, v]) => v !== undefined));
  const ctl = {
    id: uid(), type, name: over.name || uniqueName(project, c.prefix),
    x: 40, y: 40, w: c.w, h: c.h, props: clone(c.props), fx: {}, actions: {}, ...over,
  };
  if (over.props) ctl.props = { ...clone(c.props), ...over.props };
  if (type === 'gallery') {
    ctl.children = ctl.children || [];
    ctl.data = { list: '', searchCtl: '', searchField: 'Title', filterField: '', filterValue: '', sortField: 'Modified', sortDir: 'Descending', ...(over.data || {}) };
  }
  return ctl;
}

export function newScreen(project, name, over = {}) {
  return { id: uid(), name: uniqueScreenName(project, name), fill: '@Bg', onVisible: '', controls: [], ...over };
}

export function newProject() {
  const p = {
    version: 1,
    app: { name: 'My App', formFactor: 'tablet', startScreen: '', onStart: '', formulas: '', features: { loading: false, usage: false, confirm: true, stepper: true, scope: 'all', adminGroup: '' } },
    theme: defaultTheme(),
    lists: [],
    screens: [],
    flows: [],
  };
  const scr = newScreen(p, 'scrHome');
  p.screens.push(scr);
  p.app.startScreen = scr.id;
  return p;
}

// ---------- traversal ----------
export function* walk(controls, parent = null) {
  for (const c of controls || []) {
    yield [c, parent];
    if (c.children) yield* walk(c.children, c);
  }
}
export function findCtl(project, id) {
  for (const scr of project.screens) for (const [c, parent] of walk(scr.controls)) if (c.id === id) return { ctl: c, parent, screen: scr };
  return null;
}
/** Extra Power Apps control names a designer control expands into (must also be unique). */
export function companions(ctl) {
  const n = ctl.name, out = [];
  if (ctl.type === 'person') out.push(n + 'Results', n + 'ResultName', n + 'ResultTap');
  if (['icon', 'card', 'image', 'label', 'rect', 'html'].includes(ctl.type) && hasActions(ctl)) out.push(n + 'Tap');
  if (ctl.type === 'gallery' && hasActions(ctl)) out.push(n + 'RowTap');
  return out;
}
export const hasActions = (ctl) => (ctl.actions?.OnSelect || []).length > 0;

export function allNames(project, exceptId = null) {
  const names = new Set();
  for (const scr of project.screens) {
    names.add(scr.name);
    for (const [c] of walk(scr.controls)) {
      if (c.id === exceptId) continue;
      names.add(c.name);
      companions(c).forEach((x) => names.add(x));
    }
  }
  return names;
}
export function uniqueName(project, base, used = null) {
  const names = used || allNames(project);
  let i = 1;
  while (names.has(base + i)) i++;
  if (used) used.add(base + i);
  return base + i;
}
/**
 * Give a copied control (and its gallery children) fresh ids and names that are unique across the app
 * AND across the other copies being made in the same operation. Pass one shared `used` Set per operation.
 */
export function renameCopy(project, ctl, used) {
  // keep the descriptive part of the name: cardHeaderHome -> cardHeaderHome2, lbl3 -> lbl4
  const base = String(ctl.name).replace(/\d+$/, '') || CATALOG[ctl.type]?.prefix || 'ctl';
  let i = 2;
  while (used.has(base + i) || companions({ ...ctl, name: base + i }).some((n) => used.has(n))) i++;
  ctl.id = uid();
  ctl.name = base + i;
  used.add(ctl.name);
  companions(ctl).forEach((n) => used.add(n));
  (ctl.children || []).forEach((ch) => renameCopy(project, ch, used));
  return ctl;
}

// ---------------------------------------------------------------- loading untrusted project files
const SAFE_ID = /^[A-Za-z0-9_-]{1,40}$/;
/**
 * Make a project loaded from storage or a shared file safe and well-formed: ids that are safe in HTML
 * attributes (references remapped), colors that cannot break out of style attributes, numbers that are numbers.
 * Returns the same object, mutated.
 */
export function normalizeProject(p) {
  p.app = { name: 'My App', formFactor: 'tablet', startScreen: '', onStart: '', formulas: '', ...(p.app || {}) };
  if (!FORM_FACTORS[p.app.formFactor]) p.app.formFactor = 'tablet';
  for (const k of ['name', 'startScreen', 'onStart', 'formulas', 'siteUrl', 'appLink']) p.app[k] = String(p.app[k] ?? '');
  const ft = p.app.features && typeof p.app.features === 'object' ? p.app.features : {};
  p.app.features = { loading: !!ft.loading, usage: !!ft.usage, confirm: ft.confirm !== false, stepper: ft.stepper !== false, scope: ft.scope === 'mine' ? 'mine' : 'all',
    adminGroup: /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(ft.adminGroup || '').trim()) ? String(ft.adminGroup).trim() : '' };
  for (const sc of p.screens || []) if (sc && sc.role !== undefined && typeof sc.role !== 'string') delete sc.role;
  if (p.app.deepLink && typeof p.app.deepLink === 'object') p.app.deepLink = { screen: String(p.app.deepLink.screen || ''), list: String(p.app.deepLink.list || '') }; else delete p.app.deepLink;
  if (p.app.flowSample && typeof p.app.flowSample !== 'object') delete p.app.flowSample;
  const su = p.app.setup && typeof p.app.setup === 'object' ? p.app.setup : {};
  p.app.setup = { usageConfirmed: !!su.usageConfirmed, ...(typeof su.adminDraft === 'string' && su.adminDraft ? { adminDraft: su.adminDraft.slice(0, 60) } : {}) };
  const dep = p.deployed && typeof p.deployed === 'object' && !Array.isArray(p.deployed) ? p.deployed : {};
  p.deployed = {};
  for (const [k, v] of Object.entries(dep)) {
    if (!/^(lists|app|datasources|formulas|onstart|startscreen|test|scr:[A-Za-z0-9_-]{1,40}|flow:[A-Za-z0-9_-]{1,40})$/.test(k) || !v || typeof v !== 'object') continue;
    p.deployed[k] = { hash: /^[0-9a-f]{0,8}$/.test(String(v.hash ?? '')) ? String(v.hash ?? '') : '', at: String(v.at || '').slice(0, 40), name: String(v.name || '').slice(0, 120) };
  }
  p.flows = Array.isArray(p.flows) ? p.flows : [];
  const KINDS = ['created', 'statusChanged', 'reminder', 'button'];
  for (const f of p.flows) {
    if (!SAFE_ID.test(f.id || '')) f.id = uid();
    if (!KINDS.includes(f.kind)) f.kind = 'created';
    for (const k of ['name', 'list', 'subject', 'body', 'statusField', 'reminderValue', 'timeZone']) f[k] = String(f[k] ?? '');
    f.recipients = Array.isArray(f.recipients) ? f.recipients.filter((r) => r && typeof r === 'object').map((r) => ({ kind: String(r.kind || 'address'), value: String(r.value ?? '') })) : [];
    f.statusValues = Array.isArray(f.statusValues) ? f.statusValues.map(String) : [];
    f.days = clamp(Number(f.days) || 0, 0, 365); f.hour = clamp(Number(f.hour) || 0, 0, 23); f.enabled = f.enabled !== false;
  }
  const base = defaultTheme();
  p.theme = { ...base, ...(p.theme || {}) };
  p.theme.colors = { ...base.colors, ...(p.theme.colors || {}) };
  for (const k of Object.keys(p.theme.colors)) if (!isHex(p.theme.colors[k])) p.theme.colors[k] = base.colors[k] || '#000000';
  if (!FONTS.includes(p.theme.font)) p.theme.font = 'Segoe UI';
  p.theme.radius = clamp(Number(p.theme.radius) || 0, 0, 48);
  p.lists = Array.isArray(p.lists) ? p.lists : [];
  for (const l of p.lists) {
    if (!SAFE_ID.test(l.id || '')) l.id = uid();
    l.name = String(l.name || 'List'); l.columns = Array.isArray(l.columns) ? l.columns : [];
    for (const c of l.columns) { if (!SAFE_ID.test(c.id || '')) c.id = uid(); c.name = String(c.name || 'Column'); if (!COLTYPES[c.type]) c.type = 'Text'; c.choices = Array.isArray(c.choices) ? c.choices.map(String) : []; c.lookupList = String(c.lookupList || ''); c.required = !!c.required; }
  }
  p.screens = Array.isArray(p.screens) ? p.screens : [];
  const screenMap = {};
  for (const s of p.screens) {
    if (!SAFE_ID.test(s.id || '')) { const n = uid(); screenMap[s.id] = n; s.id = n; }
    s.name = String(s.name || 'scrScreen'); s.controls = Array.isArray(s.controls) ? s.controls : [];
    if (s.fill && !isColorValue(s.fill)) s.fill = '@Bg';
    s.onVisible = String(s.onVisible || '');
  }
  if (screenMap[p.app.startScreen]) p.app.startScreen = screenMap[p.app.startScreen];
  if (p.app.deepLink && screenMap[p.app.deepLink.screen]) p.app.deepLink.screen = screenMap[p.app.deepLink.screen];
  const fixCtl = (c) => {
    if (!SAFE_ID.test(c.id || '')) c.id = uid();
    if (!CATALOG[c.type]) c.type = 'raw';
    c.name = String(c.name || 'ctl');
    for (const k of ['x', 'y', 'w', 'h']) c[k] = Math.round(Number(c[k]) || 0);
    c.props = c.props && typeof c.props === 'object' ? c.props : {};
    for (const [k, v] of Object.entries(c.props)) if (/^(fill|fill2|color|borderColor|templateFill|selectedFill)$/.test(k) && !isColorValue(v)) c.props[k] = 'transparent';
    for (const k of ['size', 'borderThickness', 'gradAngle', 'templateSize', 'duration', 'stroke']) if (k in c.props) c.props[k] = Number(c.props[k]) || 0;
    c.fx = c.fx && typeof c.fx === 'object' ? c.fx : {};
    c.actions = c.actions && typeof c.actions === 'object' ? c.actions : {};
    for (const acts of Object.values(c.actions)) for (const a of acts || []) if (screenMap[a.screen]) a.screen = screenMap[a.screen];
    if (c.type === 'gallery') { c.children = Array.isArray(c.children) ? c.children : []; c.data = c.data || {}; }
    (c.children || []).forEach(fixCtl);
  };
  p.screens.forEach((s) => s.controls.forEach(fixCtl));
  return p;
}
export function uniqueScreenName(project, base) {
  const names = allNames(project);
  if (!names.has(base)) return base;
  let i = 2;
  while (names.has(base + i)) i++;
  return base + i;
}

// ---------- sample data for play mode / designer previews ----------
const FIRST = ['Alex', 'Jordan', 'Morgan', 'Taylor', 'Casey', 'Riley', 'Jamie', 'Avery', 'Quinn', 'Drew'];
const LAST = ['Rivera', 'Chen', 'Patel', 'Brooks', 'Nguyen', 'Carter', 'Diaz', 'Foster', 'Hayes', 'Kim'];
const WORDS = ['Laptop refresh', 'Badge access', 'Conference room', 'Training seat', 'Software license', 'Travel request',
  'Monitor upgrade', 'Shared drive', 'Parking pass', 'Desk move', 'Phone line', 'Printer repair'];
export const PEOPLE = FIRST.map((f, i) => ({ DisplayName: `${f} ${LAST[i]}`, Mail: `${f}.${LAST[i]}@contoso.com`.toLowerCase() }));

export function sampleValue(col, i, lists) {
  switch (col.type) {
    case 'Text': return col.name === 'Title' ? WORDS[i % WORDS.length] : `${col.name} ${i + 1}`;
    case 'Note': return 'Details for item ' + (i + 1) + '. Lorem notes go here.';
    case 'Number': return (i * 7 + 3) % 40;
    case 'Currency': return Math.round(((i * 137) % 900 + 50) * 100) / 100;
    case 'DateTime': { const d = new Date(); d.setDate(d.getDate() + ((i * 5) % 21) - 7); return d.toISOString().slice(0, 10); }
    case 'Choice': return col.choices.length ? col.choices[i % col.choices.length] : '';
    case 'YesNo': return i % 3 === 0;
    case 'Person': return PEOPLE[i % PEOPLE.length];
    case 'Lookup': {
      const tl = lists.find((l) => l.name === col.lookupList);
      return tl ? `${tl.name} ${(i % 4) + 1}` : `Item ${(i % 4) + 1}`;
    }
    default: return '';
  }
}
export function sampleRows(list, lists, n = 8) {
  const rows = [];
  for (let i = 0; i < n; i++) {
    // every third row is "mine" (Alex Rivera, the signed-in user in Play) so My items views have content
    const r = { ID: i + 1, Created: new Date(Date.now() - i * 864e5).toISOString().slice(0, 10), Modified: new Date(Date.now() - i * 36e5).toISOString().slice(0, 10),
      Author: PEOPLE[i % 3 === 0 ? 0 : (i % (PEOPLE.length - 1)) + 1], Editor: PEOPLE[(i + 1) % PEOPLE.length] };
    r['Created By'] = r.Author; r['Modified By'] = r.Editor;
    for (const c of list.columns) r[c.name] = sampleValue(c, i, lists);
    rows.push(r);
  }
  return rows;
}

/** Display text of a column value the way the generated app shows it. */
export function displayValue(col, v) {
  if (v == null || v === '') return '';
  switch (col?.type) {
    case 'Person': return v.DisplayName || '';
    case 'YesNo': return v ? 'Yes' : 'No';
    case 'Currency': return '$' + Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'DateTime': { const d = new Date(v + 'T00:00:00'); return isNaN(d) ? String(v) : `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`; }
    default: return String(v);
  }
}
export const colOf = (project, listName, colName) => {
  if (colName === 'ID') return { name: 'ID', type: 'Number' };
  if (colName === 'Created' || colName === 'Modified') return { name: colName, type: 'DateTime' };
  return project.lists.find((l) => l.name === listName)?.columns.find((c) => c.name === colName);
};
