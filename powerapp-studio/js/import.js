// import.js — paste Power Apps YAML (copied from Studio, or a Src/*.pa.yaml file) back into the designer.
// Known classic controls become editable designer controls; every property we cannot map is
// kept verbatim as a Power Fx override, and unknown controls are kept whole, so the
// round trip loses nothing.
import { newControl, newScreen, allNames } from './core.js';

const TYPE_MAP = {
  Label: 'label', Button: 'button', TextInput: 'textinput', DropDown: 'dropdown', DatePicker: 'datepicker',
  CheckBox: 'checkbox', Toggle: 'toggle', Rectangle: 'rect', HtmlViewer: 'html', Image: 'image', Gallery: 'gallery', Timer: 'timer',
};
const baseType = (t) => String(t || '').split('@')[0].replace(/^Classic\//, '');

const NAMED = { White: '#FFFFFF', Black: '#000000', Transparent: 'transparent', Red: '#FF0000', Blue: '#0000FF', Green: '#008000', Gray: '#808080', LightGray: '#D3D3D3' };
export function parseColor(f) {
  const v = String(f || '').replace(/^=/, '').trim();
  let m = v.match(/^RGBA\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)$/i);
  if (m) {
    const hx = (n) => Number(n).toString(16).padStart(2, '0').toUpperCase();
    const a = Number(m[4]);
    if (a === 0) return 'transparent';
    return '#' + hx(m[1]) + hx(m[2]) + hx(m[3]) + (a < 1 ? hx(Math.round(a * 255)) : '');
  }
  m = v.match(/^ColorValue\(\s*"(#[0-9a-f]{3,8})"\s*\)$/i); if (m) return m[1].toUpperCase();
  m = v.match(/^Color\.(\w+)$/); if (m && NAMED[m[1]]) return NAMED[m[1]];
  m = v.match(/^appTheme\.(\w+)$/); if (m && !/Hex$/.test(m[1])) return '@' + m[1];
  return null;
}
const parseNum = (f) => { const m = String(f || '').match(/^=\s*(-?\d+(?:\.\d+)?)\s*$/); return m ? Number(m[1]) : null; };
const parseStr = (f) => { const m = String(f || '').match(/^=\s*"((?:[^"]|"")*)"\s*$/s); return m ? m[1].replace(/""/g, '"') : null; };
const parseEnum = (f, e) => { const m = String(f || '').match(new RegExp('^=\\s*' + e + '\\.(\\w+)\\s*$')); return m ? m[1] : null; };
const parseBool = (f) => { const m = String(f || '').match(/^=\s*(true|false)\s*$/i); return m ? m[1].toLowerCase() === 'true' : null; };

/** Rename identifiers in a formula (skips strings, comments, member access, record keys). */
export function renameFx(f, map) {
  if (!f || !Object.keys(map).length) return f;
  let out = '', i = 0;
  const L = f.length;
  while (i < L) {
    const c = f[i];
    if (c === '"') { let j = i + 1; while (j < L) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; } out += f.slice(i, j + 1); i = j + 1; continue; }
    if (f.startsWith('//', i)) { let j = f.indexOf('\n', i); j = j < 0 ? L : j; out += f.slice(i, j); i = j; continue; }
    const m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(f.slice(i));
    if (m) {
      const w = m[0], prev = out.trimEnd().slice(-1), nxt = f.slice(i + w.length).trimStart()[0];
      const isKey = nxt === ':' && (prev === '{' || prev === ',');
      out += map[w] && prev !== '.' && !isKey ? map[w] : w; i += w.length; continue;
    }
    out += c; i++;
  }
  return out;
}

/** Keep a control exactly as Studio wrote it (type, variant, every property, children). */
export function rawControl(project, name, def) {
  const props = def.Properties || {};
  const geo = { x: parseNum(props.X) ?? 40, y: parseNum(props.Y) ?? 40, w: parseNum(props.Width) ?? 160, h: parseNum(props.Height) ?? 40 };
  const toNode = (n, d) => ({ name: n, control: d.Control, variant: d.Variant, props: Object.entries(d.Properties || {}).map(([k, v]) => [k, String(v).replace(/^=/, '')]), children: d.Children ? d.Children.map((ch) => { const [[cn, cd]] = Object.entries(ch); return toNode(cn, cd); }) : undefined });
  const raw = newControl(project, 'raw', { name, ...geo });
  raw.paControl = def.Control; raw.variant = def.Variant; raw.rawProps = { ...props };
  if (def.Children) raw.rawChildren = toNode(name, def).children;
  // position formulas (Parent.Width - 120 …) are kept as formulas
  for (const k of ['X', 'Y', 'Width', 'Height']) if (k in props && parseNum(props[k]) === null) raw.fx = { ...(raw.fx || {}), [k]: String(props[k]).replace(/^=/, '') };
  return raw;
}
function mapControl(project, name, def) {
  const props = def.Properties || {};
  const t = TYPE_MAP[baseType(def.Control)];
  const geo = { x: parseNum(props.X) ?? 40, y: parseNum(props.Y) ?? 40, w: parseNum(props.Width) ?? 160, h: parseNum(props.Height) ?? 40 };
  if (!t) return rawControl(project, name, def);
  const ctl = newControl(project, t, { name, ...geo });
  ctl.props = { ...ctl.props };
  const P = ctl.props, fx = {};
  const take = (k, parser, set) => { if (!(k in props)) return; const v = parser(props[k]); if (v === null || v === undefined) fx[k] = String(props[k]).replace(/^=/, ''); else set(v); };
  const skip = new Set(['X', 'Y', 'Width', 'Height']);
  const handled = new Set([...skip]);
  const h = (k, parser, set) => { handled.add(k); take(k, parser, set); };
  if (t === 'label' || t === 'button') {
    h('Text', parseStr, (v) => (P.text = v)); h('Size', (f) => parseNum(f), (v) => (P.size = v));
    h('FontWeight', (f) => parseEnum(f, 'FontWeight'), (v) => (P.weight = v)); h('Color', parseColor, (v) => (P.color = v));
    h('Fill', parseColor, (v) => (P.fill = v));
    if (t === 'label') { h('Align', (f) => parseEnum(f, 'Align'), (v) => (P.align = v)); h('VerticalAlign', (f) => parseEnum(f, 'VerticalAlign'), (v) => (P.valign = v)); h('Italic', parseBool, (v) => (P.italic = v)); }
    if (t === 'button') {
      h('BorderColor', parseColor, (v) => (P.borderColor = v)); h('BorderThickness', parseNum, (v) => (P.borderThickness = v));
      const rs = ['RadiusTopLeft', 'RadiusTopRight', 'RadiusBottomLeft', 'RadiusBottomRight'];
      const vals = rs.map((k) => (k in props ? parseNum(props[k]) : 5));
      if (vals.every((v) => v !== null && v === vals[0])) { rs.forEach((k) => handled.add(k)); P.radius = vals[0]; }
      // classic hover/pressed defaults we regenerate anyway
      ['HoverFill', 'PressedFill', 'HoverColor', 'PressedColor'].forEach((k) => { if (/ColorFade\(Self\.Fill|Self\.Color/.test(props[k] || '')) handled.add(k); });
    }
  }
  if (t === 'rect') { h('Fill', parseColor, (v) => (P.fill = v)); h('BorderColor', parseColor, (v) => (P.borderColor = v)); h('BorderThickness', parseNum, (v) => (P.borderThickness = v)); }
  if (t === 'html') h('HtmlText', parseStr, (v) => (P.html = v));
  if (t === 'image') h('Image', parseStr, (v) => (P.src = v));
  if (t === 'textinput') { h('Default', parseStr, (v) => (P.defaultText = v)); h('HintText', parseStr, (v) => (P.hint = v)); h('Size', parseNum, (v) => (P.size = v)); if (/MultiLine/.test(props.Mode || '')) { handled.add('Mode'); P.multiline = true; } }
  if (t === 'dropdown') {
    handled.add('Items');
    const m = String(props.Items || '').match(/^=\s*\[(.*)\]\s*$/s);
    if (m) P.items = [...m[1].matchAll(/"((?:[^"]|"")*)"/g)].map((x) => x[1].replace(/""/g, '"')).join('\n'); else if (props.Items) fx.Items = String(props.Items).replace(/^=/, '');
    h('Size', parseNum, (v) => (P.size = v));
  }
  if (t === 'checkbox') { h('Text', parseStr, (v) => (P.text = v)); h('Default', parseBool, (v) => (P.defaultOn = v)); h('Color', parseColor, (v) => (P.color = v)); h('Size', parseNum, (v) => (P.size = v)); }
  if (t === 'toggle') h('Default', parseBool, (v) => (P.defaultOn = v));
  if (t === 'timer') { h('Duration', parseNum, (v) => (P.duration = v)); h('Repeat', parseBool, (v) => (P.repeat = v)); h('AutoStart', parseBool, (v) => (P.autoStart = v)); if (props.OnTimerEnd) { handled.add('OnTimerEnd'); ctl.actions.OnTimerEnd = [{ type: 'custom', fx: String(props.OnTimerEnd).replace(/^=/, '') }]; } handled.add('Visible'); }
  if (t === 'gallery') {
    handled.add('Items'); fx.Items = String(props.Items || 'Sequence(5)').replace(/^=/, '');
    const hit = project.lists.find((l) => new RegExp(`(^|[(,\\s])${l.name.replace(/[^A-Za-z0-9_]/g, '')}\\b`).test(fx.Items));
    if (hit) ctl.data.previewList = hit.name;
    h('TemplateSize', parseNum, (v) => (P.templateSize = v)); handled.add('TemplatePadding');
    if (props.TemplateFill) { handled.add('TemplateFill'); fx.TemplateFill = String(props.TemplateFill).replace(/^=/, ''); }
    ctl.children = (def.Children || []).map((ch) => { const [[cn, cd]] = Object.entries(ch); return mapControl(project, cn, cd); });
  }
  const evName = t === 'timer' ? null : 'OnSelect';
  if (evName && props.OnSelect) { handled.add('OnSelect'); ctl.actions.OnSelect = [{ type: 'custom', fx: String(props.OnSelect).replace(/^=/, '') }]; }
  if (props.OnChange && (t === 'textinput' || t === 'dropdown')) { handled.add('OnChange'); ctl.actions.OnChange = [{ type: 'custom', fx: String(props.OnChange).replace(/^=/, '') }]; }
  for (const [k, v] of Object.entries(props)) if (!handled.has(k) && !(k in fx)) fx[k] = String(v).replace(/^=/, '');
  ctl.fx = fx;
  return ctl;
}

/** Parse YAML text. Returns {screens:[...], controls:[...], renamed:{}} or throws. */
export function importYaml(project, text, { asScreen = true } = {}) {
  const yaml = globalThis.jsyaml;
  if (!yaml) throw new Error('YAML parser did not load.');
  const doc = yaml.load(String(text).replace(/\r\n/g, '\n').replace(/\t/g, '  '), { schema: yaml.FAILSAFE_SCHEMA });
  let screensDef = null, children = null;
  if (doc && doc.Screens) screensDef = doc.Screens;
  else if (Array.isArray(doc)) children = doc;
  else if (doc && (doc.Controls || doc.Children)) children = doc.Controls || doc.Children;
  else if (doc && typeof doc === 'object') children = Object.entries(doc).map(([k, v]) => ({ [k]: v }));
  else throw new Error('This does not look like Power Apps YAML (expected "Screens:" or a list of controls).');

  // collect names and resolve clashes with the current app
  const existing = allNames(project), incoming = [];
  const collect = (list) => (list || []).forEach((it) => { const [[n, d]] = Object.entries(it); incoming.push(n); collect(d?.Children); });
  if (screensDef) for (const [sn, sd] of Object.entries(screensDef)) { incoming.push(sn); collect(sd.Children); } else collect(children);
  const map = {};
  for (const n of incoming) if (existing.has(n)) { let i = 2; while (existing.has(n + '_' + i) || incoming.includes(n + '_' + i)) i++; map[n] = n + '_' + i; existing.add(map[n]); }
  const fixDef = (d) => { if (!d || typeof d !== 'object') return d; const out = { ...d }; if (d.Properties) out.Properties = Object.fromEntries(Object.entries(d.Properties).map(([k, v]) => [k, renameFx(String(v), map)])); if (d.Children) out.Children = d.Children.map((ch) => { const [[n, cd]] = Object.entries(ch); return { [map[n] || n]: fixDef(cd) }; }); return out; };
  const result = { screens: [], controls: [], renamed: map };
  const build = (list) => (list || []).map((it) => { const [[n, d]] = Object.entries(it); return mapControl(project, n, d); });
  if (screensDef) {
    for (const [sn, sdRaw] of Object.entries(screensDef)) {
      const sd = fixDef(sdRaw);
      const scr = newScreen(project, map[sn] || sn);
      const fill = parseColor(sd.Properties?.Fill);
      if (fill) scr.fill = fill; else if (sd.Properties?.Fill) scr.fillFx = String(sd.Properties.Fill).replace(/^=/, '');
      if (sd.Properties?.OnVisible) scr.onVisible = String(sd.Properties.OnVisible).replace(/^=/, '');
      scr.controls = build(sd.Children);
      result.screens.push(scr);
    }
  } else {
    const fixed = children.map((ch) => { const [[n, d]] = Object.entries(ch); return { [map[n] || n]: fixDef(d) }; });
    result.controls = build(fixed);
    if (asScreen) { const scr = newScreen(project, 'scrImported'); scr.controls = result.controls; result.screens.push(scr); }
  }
  return result;
}
