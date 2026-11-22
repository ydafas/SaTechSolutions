// emit.js — turns the project into paste-ready Power Apps YAML (schema v3 clipboard dialect),
// App.Formulas / OnStart / StartScreen text, a deploy checklist, a PnP provisioning script,
// and lint findings. Paste rules encoded here (field-tested in Power Apps Studio):
//  * top key `Screens:`; children as a dash list; no version pins except Gallery@2.15.0 + Variant: Vertical
//  * no GalleryTemplate child; template controls nested under the gallery's Children
//  * block scalars for any formula with a newline, `: `, ` #`, `{` or `}`
//  * Rectangle gets no Radius properties; rounded shapes are HtmlViewer cards or Buttons
//  * later in the list = in front; transparent tap buttons go after what they cover
//  * names unique across the whole app; Patch always rebases with LookUp(List, ID = ...)
//  * 3-argument IfError(Patch, Notify(error), success...) shape
import { CATALOG, FORM_FACTORS, q, s, slug, varSel, hexToRgba, isHex, walk, colOf, hasActions, IDENT, RESERVED, allNames, SP_RESERVED_COLS, companions } from './core.js';
import { iconSvg } from './icons.js';

// ---------------------------------------------------------------- color helpers
export function colorFx(v) {
  if (!v || v === 'transparent') return 'RGBA(0, 0, 0, 0)';
  if (v[0] === '@') return 'appTheme.' + v.slice(1);
  if (isHex(v)) { const [r, g, b, a] = hexToRgba(v); return `RGBA(${r}, ${g}, ${b}, ${a})`; }
  return 'RGBA(0, 0, 0, 1)';
}
/** A text-valued Power Fx piece holding a CSS color (for HTML/SVG strings). Returns {fx} or {lit}. */
function cssPiece(v) {
  if (!v || v === 'transparent') return { lit: 'transparent' };
  if (v[0] === '@') return { fx: 'appTheme.' + v.slice(1) + 'Hex' };
  if (isHex(v)) { const [r, g, b, a] = hexToRgba(v); return { lit: a < 1 ? `rgba(${r},${g},${b},${a})` : v }; }
  return { lit: v };
}
/** Join literal and formula pieces into one Power Fx text expression. */
function concat(pieces) {
  const merged = [];
  for (const p of pieces) {
    if (typeof p === 'string') {
      if (merged.length && typeof merged[merged.length - 1] === 'string') merged[merged.length - 1] += p; else merged.push(p);
    } else if (p.lit != null) {
      if (merged.length && typeof merged[merged.length - 1] === 'string') merged[merged.length - 1] += p.lit; else merged.push(p.lit);
    } else merged.push(p);
  }
  return merged.map((p) => (typeof p === 'string' ? s(p) : p.fx)).join(' & ') || '""';
}

// ---------------------------------------------------------------- field formulas
/** Display text formula for a column read from `src` (ThisItem or a varSel record). */
export function readFx(project, list, field, src) {
  const col = colOf(project, list, field);
  const ref = `${src}.${q(field)}`;
  switch (col?.type) {
    case 'Number': return `Text(${ref})`;
    case 'Currency': return `Text(${ref}, "$#,##0.00")`;
    case 'DateTime': return `Text(${ref}, "mm/dd/yyyy")`;
    case 'Choice': case 'Lookup': return `${ref}.Value`;
    case 'Person': return `${ref}.DisplayName`;
    case 'YesNo': return `If(${ref}, "Yes", "No")`;
    default: return ref;
  }
}
/** Power Fx literal for a value written to a column (setField action). */
function valueFx(col, v) {
  switch (col?.type) {
    case 'Choice': return `{Value: ${s(v)}}`;
    case 'YesNo': return /^(true|yes|1)$/i.test(String(v)) ? 'true' : 'false';
    case 'Number': case 'Currency': return String(Number(v) || 0);
    case 'DateTime': return /^today$/i.test(v) || !v ? 'Today()' : `DateValue(${s(v)})`;
    case 'Person': return meClaims();
    default: return s(v);
  }
}
const meClaims = () => `{'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser", Claims: "i:0#.f|membership|" & Lower(User().Email), DisplayName: User().FullName, Email: User().Email, Department: "", JobTitle: "", Picture: ""}`;

function bindTextFx(project, ctl) {
  const b = ctl.bind;
  if (!b) return null;
  if (b.kind === 'count') {
    const L = q(b.list);
    if (!b.field) return `Text(CountRows(${L}))`;
    const col = colOf(project, b.list, b.field);
    const cond = col?.type === 'Choice' ? `${q(b.field)}.Value = ${s(b.value)}` : col?.type === 'YesNo' ? `${q(b.field)} = ${valueFx(col, b.value)}` : `${q(b.field)} = ${s(b.value)}`;
    return `Text(CountRows(Filter(${L}, ${cond})))`;
  }
  const r = b.kind === 'user' ? (b.field === 'Email' ? 'User().Email' : 'User().FullName')
    : readFx(project, b.list, b.field, b.src === 'item' ? 'ThisItem' : varSel(b.list));
  return b.prefix ? `${s(b.prefix)} & ${r}` : r;
}

// ---------------------------------------------------------------- inputs bound to columns
function inputInfo(project, ctl) {
  if (!ctl.field?.list || !ctl.field?.column) return null;
  const col = colOf(project, ctl.field.list, ctl.field.column);
  if (!col) return null;
  return { col, list: ctl.field.list, sel: varSel(ctl.field.list), ref: `${varSel(ctl.field.list)}.${q(col.name)}` };
}
const pickVar = (ctl) => 'varPick' + ctl.name;

function patchValueFx(project, ctl) {
  const f = inputInfo(project, ctl);
  if (!f) return null;
  const n = ctl.name, t = f.col.type;
  switch (ctl.type) {
    case 'textinput': return (t === 'Number' || t === 'Currency') ? `Value(${n}.Text)` : `Trim(${n}.Text)`;
    case 'dropdown': return t === 'Lookup' ? `${n}.Selected` : t === 'Choice' ? `{Value: ${n}.Selected.Value}` : `${n}.Selected.Value`;
    case 'datepicker': return `${n}.SelectedDate`;
    case 'checkbox': case 'toggle': return `${n}.Value`;
    case 'person': {
      const p = pickVar(ctl);
      return `If(IsBlank(${p}), ${f.ref}, {'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser", Claims: "i:0#.f|membership|" & Lower(Coalesce(${p}.Mail, ${p}.UserPrincipalName)), DisplayName: ${p}.DisplayName, Email: Coalesce(${p}.Mail, ${p}.UserPrincipalName), Department: "", JobTitle: "", Picture: ""})`;
    }
  }
  return null;
}
function blankCheckFx(project, ctl) {
  const f = inputInfo(project, ctl);
  if (!f || !f.col.required) return null;
  const n = ctl.name;
  switch (ctl.type) {
    case 'textinput': return `IsBlank(Trim(${n}.Text))`;
    case 'dropdown': return `IsBlank(${n}.Selected.Value)`;
    case 'datepicker': return `IsBlank(${n}.SelectedDate)`;
    case 'person': return `(IsBlank(${pickVar(ctl)}) && IsBlank(${f.ref}.DisplayName))`;
  }
  return null;
}
function screenInputs(screen, list) {
  const out = [];
  for (const [c, parent] of walk(screen.controls)) if (!parent && c.field?.list === list && ['textinput', 'dropdown', 'datepicker', 'checkbox', 'toggle', 'person'].includes(c.type)) out.push(c);
  return out;
}

// ---------------------------------------------------------------- actions
export const ACTION_TYPES = {
  navigate: 'Go to screen', back: 'Go back', newItem: 'Start a new item', openItem: 'Open this row (gallery)',
  save: 'Save form to list', delete: 'Delete selected item', setField: 'Set a column on selected item',
  setVar: 'Set variable', toggleVar: 'Toggle variable', notify: 'Show message', launch: 'Open a link',
  refresh: 'Refresh list', reset: 'Clear inputs on screen', custom: 'Custom Power Fx',
};
const scrName = (project, id) => project.screens.find((x) => x.id === id || x.name === id)?.name;

export function actionFx(project, screen, a, ctx = {}) {
  const nav = (id, tr) => { const n = scrName(project, id); return n ? `Navigate(${n}, ScreenTransition.${tr || 'Fade'})` : ''; };
  const then = (a2) => a2.then === 'navigate' ? nav(a2.screen) : a2.then === 'back' ? 'Back()' : '';
  const L = a.list ? q(a.list) : '';
  const V = a.list ? varSel(a.list) : '';
  switch (a.type) {
    case 'navigate': return nav(a.screen, a.transition);
    case 'back': return 'Back()';
    case 'newItem': return [`Set(${V}, Defaults(${L}))`, a.screen ? nav(a.screen) : ''].filter(Boolean).join(';\n');
    case 'openItem': return [`Set(${V}, ThisItem)`, a.screen ? nav(a.screen) : ''].filter(Boolean).join(';\n');
    case 'save': {
      const inputs = screenInputs(screen, a.list);
      const fields = inputs.map((c) => [c.field.column, patchValueFx(project, c)]).filter((x) => x[1]);
      const rec = fields.length ? '{\n' + fields.map(([k, v]) => `        ${q(k)}: ${v}`).join(',\n') + '\n    }' : '{}';
      const done = [`Notify(${s(a.message || 'Saved')}, NotificationType.Success)`];
      if (a.then === 'stay') {
        done.unshift(`If(IsBlank(${V}.ID), Set(${V}, Defaults(${L})), Set(${V}, LookUp(${L}, ID = ${V}.ID)))`);
        inputs.forEach((c) => done.push(`Reset(${c.name})`));
        inputs.filter((c) => c.type === 'person').forEach((c) => done.push(`Set(${pickVar(c)}, Blank())`));
      } else if (then(a)) done.push(then(a));
      return `IfError(\n    Patch(${L}, Coalesce(LookUp(${L}, ID = ${V}.ID), Defaults(${L})), ${rec}),\n    Notify("Save failed. " & FirstError.Message, NotificationType.Error),\n    ${done.join(';\n    ')}\n)`;
    }
    case 'delete': {
      const done = [`Notify(${s(a.message || 'Deleted')}, NotificationType.Success)`, `Set(${V}, Defaults(${L}))`];
      if (then(a)) done.push(then(a));
      return `IfError(\n    Remove(${L}, LookUp(${L}, ID = ${V}.ID)),\n    Notify("Delete failed. " & FirstError.Message, NotificationType.Error),\n    ${done.join(';\n    ')}\n)`;
    }
    case 'setField': {
      const col = colOf(project, a.list, a.field);
      if (!col) return '';
      const done = [`Set(${V}, LookUp(${L}, ID = ${V}.ID))`, `Notify(${s(a.message || 'Updated')}, NotificationType.Success)`];
      if (then(a)) done.push(then(a));
      return `IfError(\n    Patch(${L}, LookUp(${L}, ID = ${V}.ID), {${q(a.field)}: ${valueFx(col, a.value)}}),\n    Notify("Update failed. " & FirstError.Message, NotificationType.Error),\n    ${done.join(';\n    ')}\n)`;
    }
    case 'setVar': return a.name ? `Set(${a.name}, ${a.value || 'true'})` : '';
    case 'toggleVar': return a.name ? `Set(${a.name}, !${a.name})` : '';
    case 'notify': return `Notify(${s(a.text || 'Done')}, NotificationType.${a.kind || 'Information'})`;
    case 'launch': return a.url ? `Launch(${s(a.url)})` : '';
    case 'refresh': return a.list ? `Refresh(${L})` : '';
    case 'reset': {
      const ins = []; for (const [c, parent] of walk(screen.controls)) if (!parent && ['textinput', 'dropdown', 'datepicker', 'checkbox', 'toggle', 'person'].includes(c.type)) ins.push(c);
      return ins.map((c) => `Reset(${c.name})`).concat(ins.filter((c) => c.type === 'person').map((c) => `Set(${pickVar(c)}, Blank())`)).join(';\n');
    }
    case 'custom': return (a.fx || '').replace(/^=/, '');
  }
  return '';
}
export function actionsFx(project, screen, list) {
  return (list || []).map((a) => actionFx(project, screen, a)).filter(Boolean).join(';\n');
}

// ---------------------------------------------------------------- visibility
export function visibleFx(v) {
  if (!v?.var) return null;
  const lit = /^-?\d+(\.\d+)?$/.test(v.value ?? '') ? v.value : /^(true|false)$/.test(v.value) ? v.value : s(v.value ?? '');
  switch (v.op) {
    case 'false': return `!${v.var}`;
    case 'eq': return `${v.var} = ${lit}`;
    case 'neq': return `${v.var} <> ${lit}`;
    default: return v.var;
  }
}

// ---------------------------------------------------------------- gallery items
export function itemsFx(project, ctl) {
  const d = ctl.data || {};
  if (!d.list) return 'Sequence(5)';
  const L = q(d.list), conds = [];
  if (d.searchCtl && d.searchField) conds.push(`StartsWith(${q(d.searchField)}, ${d.searchCtl}.Text)`);
  if (d.filterField) {
    if (d.filterField === 'CreatedByMe') conds.push(`'Created By'.Email = User().Email`);
    else {
      const col = colOf(project, d.list, d.filterField);
      if (col?.type === 'Choice') conds.push(`${q(d.filterField)}.Value = ${s(d.filterValue)}`);
      else if (col?.type === 'Person') conds.push(`${q(d.filterField)}.Email = User().Email`);
      else if (col?.type === 'YesNo') conds.push(`${q(d.filterField)} = ${valueFx(col, d.filterValue)}`);
      else if (col?.type === 'Number' || col?.type === 'Currency') conds.push(`${q(d.filterField)} = ${Number(d.filterValue) || 0}`);
      else if (col) conds.push(`${q(d.filterField)} = ${s(d.filterValue)}`);
    }
  }
  if (d.extraFilter) conds.push(`(${d.extraFilter})`);
  const src = conds.length ? `Filter(${L}, ${conds.join(' && ')})` : L;
  return d.sortField ? `SortByColumns(${src}, ${s(d.sortField)}, SortOrder.${d.sortDir || 'Descending'})` : src;
}

// ---------------------------------------------------------------- per-control properties
const N = (v) => String(Math.round(Number(v) || 0));
const radiusFx = (r) => (r === 'theme' || r == null ? 'appTheme.Radius' : N(r));
const align = (a) => `Align.${a || 'Left'}`;
const weight = (w) => `FontWeight.${w || 'Normal'}`;

function cardHtmlFx(ctl) {
  const p = ctl.props, shadow = !!p.shadow;
  const h = shadow ? ctl.h - 12 : ctl.h + 10;
  const bg = p.fill2 ? ['linear-gradient(', String(p.gradAngle ?? 135), 'deg,', cssPiece(p.fill), ',', cssPiece(p.fill2), ')'] : [cssPiece(p.fill)];
  const rad = p.radius === 'theme' || p.radius == null ? { fx: 'Text(appTheme.Radius)' } : { lit: N(p.radius) };
  const pieces = [`<div style='${shadow ? 'margin:1px' : 'margin:-5px'};box-sizing:border-box;height:${h}px;background:`, ...bg,
    `;border-radius:`, rad, 'px;'];
  if (Number(p.borderThickness) > 0) pieces.push(`border:${N(p.borderThickness)}px solid `, cssPiece(p.borderColor), ';');
  if (shadow) pieces.push('box-shadow:0 2px 6px rgba(15,23,42,0.14);');
  pieces.push(`'></div>`);
  return concat(pieces);
}
function iconFx(ctl) {
  const MARK = '\u0001';
  const [a, b] = iconSvg(ctl.props.icon, MARK, ctl.props.stroke || 2).split(MARK);
  return `"data:image/svg+xml;utf8," & EncodeUrl(${concat([a, cssPiece(ctl.props.color), b])})`;
}

/** Ordered [prop, formula] pairs for one designer control (not companions). */
function controlProps(project, screen, ctl, ctx) {
  const p = ctl.props || {}, P = [];
  const add = (k, v) => { if (v != null && v !== '') P.push([k, String(v)]); };
  const geo = () => { add('X', N(ctl.x)); add('Y', N(ctl.y)); add('Width', N(ctl.w)); add('Height', N(ctl.h)); };
  const vis = visibleFx(ctl.visibleWhen);
  const f = inputInfo(project, ctl);
  const font = 'appTheme.Font';
  switch (ctl.type) {
    case 'label':
      add('Text', bindTextFx(project, ctl) ?? s(p.text)); geo();
      add('Size', N(p.size)); add('FontWeight', weight(p.weight)); add('Font', font);
      add('Color', colorFx(p.color)); add('Fill', colorFx(p.fill)); add('Align', align(p.align));
      add('VerticalAlign', `VerticalAlign.${p.valign || 'Middle'}`); if (p.italic) add('Italic', 'true');
      break;
    case 'button': {
      add('Text', bindTextFx(project, ctl) ?? s(p.text)); geo();
      add('Size', N(p.size)); add('FontWeight', weight(p.weight)); add('Font', font);
      add('Fill', colorFx(p.fill)); add('Color', colorFx(p.color));
      add('BorderColor', colorFx(p.borderColor)); add('BorderThickness', N(p.borderThickness));
      for (const k of ['RadiusTopLeft', 'RadiusTopRight', 'RadiusBottomLeft', 'RadiusBottomRight']) add(k, radiusFx(p.radius));
      const transparent = !p.fill || p.fill === 'transparent';
      add('HoverFill', transparent ? 'appTheme.HoverTint' : 'ColorFade(Self.Fill, appTheme.HoverFade)');
      add('PressedFill', transparent ? 'appTheme.PressTint' : 'ColorFade(Self.Fill, appTheme.PressFade)');
      add('HoverColor', 'Self.Color'); add('PressedColor', 'Self.Color');
      add('FocusedBorderColor', 'appTheme.Primary'); add('FocusedBorderThickness', '2');
      const save = (ctl.actions?.OnSelect || []).find((a) => a.type === 'save');
      if (save) {
        const checks = screenInputs(screen, save.list).map((c) => blankCheckFx(project, c)).filter(Boolean);
        if (checks.length) add('DisplayMode', `If(${checks.join(' || ')}, DisplayMode.Disabled, DisplayMode.Edit)`);
      }
      break;
    }
    case 'badge': {
      add('Text', bindTextFx(project, ctl) ?? s(p.text)); geo();
      add('Size', N(p.size)); add('FontWeight', weight(p.weight)); add('Font', font);
      const map = String(p.colorMap || '').split('\n').map((l) => l.split('=')).filter((x) => x.length === 2 && x[0].trim());
      const t = bindTextFx(project, ctl);
      add('Fill', t && map.length ? `Switch(${t}, ${map.map(([k, v]) => `${s(k.trim())}, ${colorFx(v.trim())}`).join(', ')}, ${colorFx(p.fill)})` : colorFx(p.fill));
      add('Color', colorFx(p.color)); add('BorderThickness', '0');
      for (const k of ['RadiusTopLeft', 'RadiusTopRight', 'RadiusBottomLeft', 'RadiusBottomRight']) add(k, radiusFx(p.radius));
      add('HoverFill', 'Self.Fill'); add('PressedFill', 'Self.Fill'); add('HoverColor', 'Self.Color'); add('PressedColor', 'Self.Color');
      break;
    }
    case 'icon': add('Image', iconFx(ctl)); geo(); break;
    case 'image': {
      const src = p.src || '';
      add('Image', src ? s(src) : '""'); geo();
      if (p.fit && p.fit !== 'Fit') add('ImagePosition', `ImagePosition.${p.fit}`);
      break;
    }
    case 'card': add('HtmlText', cardHtmlFx(ctl)); geo(); break;
    case 'html': add('HtmlText', s(p.html)); geo(); break;
    case 'rect': geo(); add('Fill', colorFx(p.fill)); add('BorderColor', colorFx(p.borderColor)); add('BorderThickness', N(p.borderThickness)); break;
    case 'textinput': {
      const num = f ? ['Number', 'Currency'].includes(f.col.type) : p.numeric;
      add('Default', f ? (num ? `Text(${f.ref})` : f.ref) : s(p.defaultText));
      add('HintText', s(p.hint));
      if (f ? f.col.type === 'Note' : p.multiline) add('Mode', 'TextMode.MultiLine');
      if (num) add('Format', 'TextFormat.Number');
      geo(); add('Size', N(p.size)); add('Font', font); add('Color', 'appTheme.Text'); add('Fill', 'appTheme.InputFill');
      add('BorderColor', 'appTheme.InputBorder'); add('BorderThickness', '1');
      add('FocusedBorderColor', 'appTheme.Primary'); add('FocusedBorderThickness', '2');
      break;
    }
    case 'dropdown': {
      if (f?.col.type === 'Lookup') { add('Items', `Choices([@${q(f.list)}].${q(f.col.name)})`); add('Default', `${f.ref}.Value`); }
      else {
        const items = f?.col.type === 'Choice' ? f.col.choices : String(p.items || '').split('\n').map((x) => x.trim()).filter(Boolean);
        add('Items', '[' + items.map(s).join(', ') + ']');
        if (f) add('Default', `Text(${f.ref}.Value)`);
      }
      geo(); add('Size', N(p.size)); add('Font', font); add('Color', 'appTheme.Text'); add('Fill', 'appTheme.InputFill');
      add('ChevronBackground', 'appTheme.Primary'); add('ChevronFill', 'appTheme.OnPrimary');
      add('BorderColor', 'appTheme.InputBorder'); add('BorderThickness', '1'); add('HoverFill', 'appTheme.SurfaceAlt');
      add('FocusedBorderColor', 'appTheme.Primary'); add('FocusedBorderThickness', '2');
      break;
    }
    case 'datepicker':
      add('DefaultDate', f ? (f.col.required ? `Coalesce(${f.ref}, Today())` : f.ref) : 'Today()');
      geo(); add('Size', N(p.size)); add('Font', font); add('Color', 'appTheme.Text'); add('Fill', 'appTheme.InputFill');
      add('BorderColor', 'appTheme.InputBorder'); add('IconBackground', 'appTheme.Primary'); add('IconFill', 'appTheme.OnPrimary');
      break;
    case 'checkbox':
      add('Text', s(p.text)); add('Default', f ? `Coalesce(${f.ref}, false)` : (p.defaultOn ? 'true' : 'false'));
      geo(); add('Size', N(p.size)); add('Font', font); add('Color', colorFx(p.color));
      add('FocusedBorderColor', 'appTheme.Primary'); add('FocusedBorderThickness', '2');
      break;
    case 'toggle':
      add('Default', f ? `Coalesce(${f.ref}, false)` : (p.defaultOn ? 'true' : 'false'));
      geo(); add('TrueFill', 'appTheme.Primary');
      break;
    case 'person': {
      const pv = pickVar(ctl);
      add('Default', f ? `If(IsBlank(${pv}), Coalesce(${f.ref}.DisplayName, ""), ${pv}.DisplayName)` : `If(IsBlank(${pv}), "", ${pv}.DisplayName)`);
      add('HintText', s(p.hint)); add('OnChange', `Set(${pv}, Blank())`);
      geo(); add('Size', N(p.size)); add('Font', font); add('Color', 'appTheme.Text'); add('Fill', 'appTheme.InputFill');
      add('BorderColor', 'appTheme.InputBorder'); add('BorderThickness', '1');
      add('FocusedBorderColor', 'appTheme.Primary'); add('FocusedBorderThickness', '2');
      break;
    }
    case 'gallery':
      add('Items', ctl.fx?.Items ? null : itemsFx(project, ctl)); geo();
      add('TemplateSize', N(p.templateSize)); add('TemplatePadding', '0');
      add('TemplateFill', `If(ThisItem.IsSelected, ${colorFx(p.selectedFill)}, ${colorFx(p.templateFill)})`);
      break;
    case 'timer':
      add('Duration', N(p.duration)); add('Repeat', p.repeat ? 'true' : 'false'); add('AutoStart', p.autoStart ? 'true' : 'false');
      add('OnTimerEnd', actionsFx(project, screen, ctl.actions?.OnTimerEnd)); add('Visible', 'false'); geo();
      break;
    case 'raw':
      for (const [k, v] of Object.entries(ctl.rawProps || {})) if (!['X', 'Y', 'Width', 'Height'].includes(k)) add(k, v.replace(/^=/, ''));
      add('X', N(ctl.x)); add('Y', N(ctl.y)); add('Width', N(ctl.w)); add('Height', N(ctl.h));
      break;
  }
  if (vis && ctl.type !== 'timer') add('Visible', vis);
  if (ctl.type === 'button' || ctl.type === 'badge') add('OnSelect', actionsFx(project, screen, ctl.actions?.OnSelect));
  if (ctl.type === 'dropdown' || ctl.type === 'textinput') add('OnChange', actionsFx(project, screen, ctl.actions?.OnChange));
  if (ctl.type === 'toggle' || ctl.type === 'checkbox') { add('OnCheck', actionsFx(project, screen, ctl.actions?.OnCheck)); add('OnUncheck', actionsFx(project, screen, ctl.actions?.OnUncheck)); }
  // advanced overrides win
  for (const [k, v] of Object.entries(ctl.fx || {})) {
    if (!v && v !== 0) continue;
    const i = P.findIndex(([kk]) => kk === k);
    const val = String(v).replace(/^=/, '');
    if (i >= 0) P[i][1] = val; else P.push([k, val]);
  }
  return P;
}

const tapButton = (name, ctl, onSelect, hover = 'appTheme.HoverTint') => ({
  name, control: 'Classic/Button', props: [
    ['Text', '""'], ['X', N(ctl.x)], ['Y', N(ctl.y)], ['Width', N(ctl.w)], ['Height', N(ctl.h)],
    ['Fill', 'RGBA(0, 0, 0, 0)'], ['Color', 'RGBA(0, 0, 0, 0)'], ['BorderThickness', '0'],
    ['RadiusTopLeft', '0'], ['RadiusTopRight', '0'], ['RadiusBottomLeft', '0'], ['RadiusBottomRight', '0'],
    ['HoverFill', hover], ['PressedFill', 'appTheme.PressTint'], ['HoverColor', 'Self.Color'], ['PressedColor', 'Self.Color'],
    ['FocusedBorderThickness', '0'],
    ...(visibleFx(ctl.visibleWhen) ? [['Visible', visibleFx(ctl.visibleWhen)]] : []),
    ['OnSelect', onSelect],
  ],
});

/** Build emitted nodes for a list of designer controls. Returns {nodes, tail}. */
function buildNodes(project, screen, controls, inGallery, gal) {
  const nodes = [], tail = [];
  for (const ctl of controls) {
    const def = CATALOG[ctl.type];
    if (!def) continue;
    const node = { name: ctl.name, control: ctl.type === 'raw' ? ctl.paControl : def.pa, variant: ctl.type === 'gallery' ? 'Vertical' : ctl.type === 'raw' ? ctl.variant : undefined, props: controlProps(project, screen, ctl, { inGallery }) };
    nodes.push(node);
    if (ctl.type === 'gallery') {
      const inner = buildNodes(project, screen, ctl.children || [], true, ctl);
      node.children = inner.nodes;
      if (hasActions(ctl)) node.children.push(tapButton(ctl.name + 'RowTap', { x: 0, y: 0, w: 0, h: 0 }, actionsFx(project, screen, ctl.actions.OnSelect)));
      const rt = node.children[node.children.length - 1];
      if (rt?.name === ctl.name + 'RowTap') { rt.props.find((x) => x[0] === 'Width')[1] = 'Parent.TemplateWidth'; rt.props.find((x) => x[0] === 'Height')[1] = 'Parent.TemplateHeight'; }
      tail.push(...inner.tail);
    } else if (ctl.type === 'raw' && ctl.rawChildren) {
      node.children = ctl.rawChildren;
    } else if (['icon', 'card', 'image', 'label', 'rect', 'html'].includes(ctl.type) && hasActions(ctl)) {
      nodes.push(tapButton(ctl.name + 'Tap', ctl, actionsFx(project, screen, ctl.actions.OnSelect)));
    }
    if (ctl.type === 'person' && !inGallery) {
      const pv = pickVar(ctl), f = inputInfo(project, ctl);
      const hideWhenSame = f ? ` && ${ctl.name}.Text <> Coalesce(${f.ref}.DisplayName, "")` : '';
      tail.push({
        name: ctl.name + 'Results', control: 'Gallery@2.15.0', variant: 'Vertical', props: [
          ['Items', `FirstN(Office365Users.SearchUser({searchTerm: ${ctl.name}.Text}), 5)`],
          ['X', N(ctl.x)], ['Y', N(ctl.y + ctl.h)], ['Width', N(ctl.w)], ['Height', '180'],
          ['TemplateSize', '36'], ['TemplatePadding', '0'], ['TemplateFill', 'appTheme.Surface'],
          ['BorderColor', 'appTheme.Border'], ['BorderThickness', '1'],
          ['Visible', `Len(${ctl.name}.Text) > 1 && IsBlank(${pv})${hideWhenSame}` + (visibleFx(ctl.visibleWhen) ? ` && ${visibleFx(ctl.visibleWhen)}` : '')],
        ],
        children: [
          { name: ctl.name + 'ResultName', control: 'Label', props: [
            ['Text', 'ThisItem.DisplayName & "  -  " & Coalesce(ThisItem.Mail, ThisItem.UserPrincipalName)'],
            ['X', '10'], ['Y', '0'], ['Width', 'Parent.TemplateWidth - 20'], ['Height', '36'], ['Size', '11'],
            ['Font', 'appTheme.Font'], ['Color', 'appTheme.Text']] },
          tapButton(ctl.name + 'ResultTap', { x: 0, y: 0, w: 0, h: 0 }, `Set(${pv}, ThisItem);\nReset(${ctl.name})`),
        ],
      });
      const t = tail[tail.length - 1].children[1];
      t.props.find((x) => x[0] === 'Width')[1] = 'Parent.TemplateWidth'; t.props.find((x) => x[0] === 'Height')[1] = 'Parent.TemplateHeight';
    }
  }
  return { nodes, tail };
}

// ---------------------------------------------------------------- YAML writer
const needsBlock = (v) => /\n|: | #|[{}\t]|:$|\s$/.test(v);
const ALWAYS_BLOCK = new Set(['HtmlText', 'Image', 'OnSelect', 'Items']);
function yProp(pad, key, fx) {
  const v = '=' + fx;
  if (!needsBlock(v) && !(ALWAYS_BLOCK.has(key) && v.length > 60)) return `${pad}${key}: ${v}`;
  return `${pad}${key}: |-\n` + v.split('\n').map((l) => (l ? pad + '  ' + l : '')).join('\n');
}
function yNodes(nodes, pad) {
  const out = [];
  for (const n of nodes) {
    out.push(`${pad}- ${n.name}:`);
    const p2 = pad + '    ';
    out.push(`${p2}Control: ${n.control}`);
    if (n.variant) out.push(`${p2}Variant: ${n.variant}`);
    if (n.props.length) {
      out.push(`${p2}Properties:`);
      for (const [k, v] of n.props) out.push(yProp(p2 + '  ', k, v));
    }
    if (n.children) {
      out.push(`${p2}Children:`);
      out.push(...yNodes(n.children, p2));
    }
  }
  return out;
}
function screenOnVisible(project, screen) {
  const parts = [];
  if (screen.onVisible) parts.push(screen.onVisible.replace(/^=/, ''));
  const inputs = []; for (const [c, parent] of walk(screen.controls)) if (!parent && c.field?.list && ['textinput', 'dropdown', 'datepicker', 'checkbox', 'toggle', 'person'].includes(c.type)) inputs.push(c);
  inputs.filter((c) => c.type === 'person').forEach((c) => parts.push(`Set(${pickVar(c)}, Blank())`));
  inputs.forEach((c) => parts.push(`Reset(${c.name})`));
  return parts.join(';\n');
}
export function screenYaml(project, screen) {
  const { nodes, tail } = buildNodes(project, screen, screen.controls, false);
  const lines = ['Screens:', `  ${screen.name}:`, '    Properties:', yProp('      ', 'Fill', screen.fillFx || colorFx(screen.fill || '@Bg'))];
  const ov = screenOnVisible(project, screen);
  if (ov) lines.push(yProp('      ', 'OnVisible', ov));
  const all = [...nodes, ...tail];
  if (all.length) { lines.push('    Children:'); lines.push(...yNodes(all, '    ')); }
  return lines.join('\n') + '\n';
}
/** YAML for one designer control (plus its companions), as it appears inside the screen. */
export function controlYaml(project, screen, ctl) {
  const inGallery = !screen.controls.includes(ctl);
  const { nodes, tail } = buildNodes(project, screen, [ctl], inGallery);
  return yNodes([...nodes, ...tail], '').join('\n');
}
/** Flat list of every Power Apps control name emitted for a screen (for checks). */
export function emittedNames(project, screen) {
  const { nodes, tail } = buildNodes(project, screen, screen.controls, false);
  const out = [];
  const rec = (ns) => ns.forEach((n) => { out.push(n.name); if (n.children) rec(n.children.filter((c) => c.name)); });
  rec([...nodes, ...tail]);
  return out;
}

// ---------------------------------------------------------------- App-level text
export function themeFormulas(project) {
  const t = project.theme, c = t.colors, lines = [];
  for (const [k, v] of Object.entries(c)) { lines.push(`    ${k}: ${colorFx(v)},`); lines.push(`    ${k}Hex: ${s(v)},`); }
  const [r, g, b] = hexToRgba(c.Text);
  lines.push(`    HoverTint: RGBA(${r}, ${g}, ${b}, ${t.dark ? 0.12 : 0.06}),`);
  lines.push(`    PressTint: RGBA(${r}, ${g}, ${b}, ${t.dark ? 0.2 : 0.12}),`);
  lines.push(`    HoverFade: ${t.dark ? 0.18 : -0.12},`);
  lines.push(`    PressFade: ${t.dark ? 0.3 : -0.25},`);
  lines.push(`    Radius: ${Number(t.radius) || 0},`);
  lines.push(`    Font: Font.${q(t.font || 'Segoe UI')}`);
  let out = `// Theme generated by Power App Studio (${t.name || 'custom'}). Edit a value here to restyle the whole app.\nappTheme = {\n${lines.join('\n')}\n};`;
  if (project.app.formulas?.trim()) out += '\n\n' + project.app.formulas.trim();
  return out;
}
/** Variables that need a typed starting value. */
export function seededVars(project) {
  const vars = new Map();
  const visit = (acts) => (acts || []).forEach((a) => {
    if (a.type === 'toggleVar' && a.name) vars.set(a.name, 'false');
    if (a.type === 'setVar' && a.name && !vars.has(a.name)) {
      const v = String(a.value ?? 'true').trim();
      vars.set(a.name, /^(true|false)$/i.test(v) ? 'false' : /^-?\d/.test(v) ? '0' : /^"/.test(v) ? '""' : null);
    }
  });
  for (const scr of project.screens) for (const [c] of walk(scr.controls)) {
    Object.values(c.actions || {}).forEach(visit);
    if (c.visibleWhen?.var && !vars.has(c.visibleWhen.var)) vars.set(c.visibleWhen.var, c.visibleWhen.op === 'eq' || c.visibleWhen.op === 'neq' ? (/^-?\d/.test(c.visibleWhen.value) ? '0' : '""') : 'false');
  }
  return [...vars].filter(([, v]) => v != null);
}
export function onStartFx(project) {
  const parts = seededVars(project).map(([k, v]) => `Set(${k}, ${v})`);
  const lists = usedLists(project);
  for (const l of lists) if (formListUsed(project, l)) parts.push(`Set(${varSel(l)}, Defaults(${q(l)}))`);
  if (project.app.onStart?.trim()) parts.push(project.app.onStart.trim().replace(/^=/, ''));
  return parts.join(';\n');
}
function formListUsed(project, list) {
  for (const scr of project.screens) for (const [c] of walk(scr.controls)) {
    if (c.field?.list === list) return true;
    if (c.bind?.list === list && c.bind.src === 'sel') return true;
  }
  return false;
}
export function startScreenName(project) {
  return (project.screens.find((x) => x.id === project.app.startScreen) || project.screens[0])?.name;
}
export function usedLists(project) {
  const set = new Set();
  for (const scr of project.screens) for (const [c] of walk(scr.controls)) {
    if (c.data?.list) set.add(c.data.list);
    if (c.field?.list) set.add(c.field.list);
    if (c.bind?.list) set.add(c.bind.list);
    Object.values(c.actions || {}).flat().forEach((a) => a.list && set.add(a.list));
  }
  return [...set].filter((l) => project.lists.some((x) => x.name === l));
}
export function usesPeople(project) {
  for (const scr of project.screens) for (const [c] of walk(scr.controls)) if (c.type === 'person') return true;
  return false;
}
export function checklist(project) {
  const lists = usedLists(project), steps = [];
  const ff = FORM_FACTORS[project.app.formFactor || 'tablet'];
  steps.push(`Create a blank canvas app (${project.app.formFactor === 'phone' ? 'Phone' : 'Tablet'} format, ${ff.w} × ${ff.h}).`);
  if (lists.length || usesPeople(project)) steps.push('Data → Add data: ' + [...lists.map((l) => `SharePoint list "${l}"`), ...(usesPeople(project) ? ['Office 365 Users'] : [])].join(', ') + '. Names must match exactly.');
  steps.push('App → Formulas: paste the App.Formulas text (it holds the theme).');
  if (onStartFx(project)) steps.push('App → OnStart: expand the formula bar first, paste the OnStart text, then App → Run OnStart.');
  steps.push(`For each screen: delete the default Screen1 (or the old copy), click the empty canvas area, press Ctrl+V with that screen's YAML copied.`);
  steps.push(`App → StartScreen: =${startScreenName(project)}`);
  if (usesPeople(project)) steps.push('People pickers: type a name in Play mode. If results are blank, check the field casing with intellisense on First(Office365Users.SearchUser({searchTerm:"a"})). and adjust DisplayName / Mail.');
  steps.push('Press Play (F5) and test from the start screen: OnVisible and timers do not run on the editing canvas.');
  steps.push('Only expected warnings left should be delegation warnings (CountRows, Choice filters). Raise the data row limit to 2000 in Settings if your lists are small.');
  return steps;
}

// ---------------------------------------------------------------- PnP provisioning script
const xmlEsc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&apos;').replace(/"/g, '&quot;');
const psEsc = (t) => String(t).replace(/'/g, "''");
export const internalName = (n) => String(n).replace(/[^A-Za-z0-9_]/g, '') || 'Field';
export function provisioningScript(project) {
  const lists = [...project.lists];
  // lookup targets first
  const order = [], seen = new Set();
  const visit = (l, stack = new Set()) => {
    if (seen.has(l.name) || stack.has(l.name)) return; stack.add(l.name);
    l.columns.filter((c) => c.type === 'Lookup' && c.lookupList).forEach((c) => { const t = lists.find((x) => x.name === c.lookupList); if (t) visit(t, stack); });
    seen.add(l.name); order.push(l);
  };
  lists.forEach((l) => visit(l));
  const out = [];
  out.push('# =====================================================================');
  out.push(`# Provision SharePoint lists for "${project.app.name}" (generated by Power App Studio)`);
  out.push('# Requires PnP.PowerShell 2.x and an Entra app registration (client ID) allowed to use it.');
  out.push('# Safe to re-run: existing lists and columns are reported and left alone.');
  out.push('# =====================================================================');
  out.push('');
  out.push('# ===== EDIT ME =====');
  out.push("$SiteUrl     = 'https://contoso.sharepoint.com/sites/YourSite'");
  out.push("$ClientId    = '00000000-0000-0000-0000-000000000000'   # your Entra app (client) ID");
  out.push("$Environment = 'Production'   # Production | USGovernment (GCC) | USGovernmentHigh (GCC High) | USGovernmentDoD");
  out.push('# ===================');
  out.push('');
  out.push('Import-Module PnP.PowerShell');
  out.push('Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive -AzureEnvironment $Environment');
  out.push('');
  out.push('function Ensure-List($Title) {');
  out.push('    $l = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue');
  out.push('    if ($l) { Write-Host "[exists] list $Title" -ForegroundColor DarkGray }');
  out.push('    else { $l = New-PnPList -Title $Title -Template GenericList -OnQuickLaunch; Write-Host "[added]  list $Title" -ForegroundColor Green }');
  out.push('    return $l');
  out.push('}');
  out.push('function Ensure-Field($List, $Name, $Xml) {');
  out.push('    if (Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue) { Write-Host "[exists] $List.$Name" -ForegroundColor DarkGray; return }');
  out.push('    Add-PnPFieldFromXml -List $List -FieldXml $Xml | Out-Null');
  out.push('    try { $v = Get-PnPView -List $List | Where-Object DefaultView; Set-PnPView -List $List -Identity $v.Id -Fields (@($v.ViewFields) + $Name) | Out-Null } catch {}');
  out.push('    Write-Host "[added]  $List.$Name" -ForegroundColor Green');
  out.push('}');
  out.push('');
  for (const l of order) {
    out.push(`# ---- ${l.name}`);
    out.push(`$null = Ensure-List '${psEsc(l.name)}'`);
    for (const c of l.columns) {
      if (c.builtin) { if (c.required === false) out.push(`Set-PnPField -List '${psEsc(l.name)}' -Identity 'Title' -Values @{Required=$false} | Out-Null`); continue; }
      const iname = internalName(c.name), dn = xmlEsc(c.name), req = c.required ? "Required='TRUE'" : "Required='FALSE'";
      const base = `Name='${iname}' StaticName='${iname}' DisplayName='${dn}' ${req}`;
      let xml;
      switch (c.type) {
        case 'Note': xml = `<Field Type='Note' ${base} NumLines='6' RichText='FALSE' />`; break;
        case 'Number': xml = `<Field Type='Number' ${base} />`; break;
        case 'Currency': xml = `<Field Type='Currency' ${base} LCID='1033' />`; break;
        case 'DateTime': xml = `<Field Type='DateTime' ${base} Format='DateOnly' />`; break;
        case 'Choice': xml = `<Field Type='Choice' ${base} Format='Dropdown'><CHOICES>${c.choices.map((x) => `<CHOICE>${xmlEsc(x)}</CHOICE>`).join('')}</CHOICES>${c.choices[0] ? `<Default>${xmlEsc(c.choices[0])}</Default>` : ''}</Field>`; break;
        case 'YesNo': xml = `<Field Type='Boolean' ${base.replace(/Required='TRUE'/, "Required='FALSE'")}><Default>0</Default></Field>`; break;
        case 'Person': xml = `<Field Type='User' ${base} UserSelectionMode='PeopleOnly' />`; break;
        case 'Lookup': xml = null; break;
        default: xml = `<Field Type='Text' ${base} MaxLength='255' />`;
      }
      if (c.type === 'Lookup') {
        out.push(`$target = Get-PnPList -Identity '${psEsc(c.lookupList || 'MISSING')}'`);
        out.push(`Ensure-Field '${psEsc(l.name)}' '${iname}' ("<Field Type='Lookup' ${base.replace(/'/g, "''")} List='{" + $target.Id + "}' ShowField='Title' />")`);
      } else out.push(`Ensure-Field '${psEsc(l.name)}' '${iname}' '${xml.replace(/'/g, "''")}'`);
    }
    out.push('');
  }
  out.push('Write-Host "Done." -ForegroundColor Cyan');
  return out.join('\n');
}

// ---------------------------------------------------------------- lint
export function lint(project) {
  const F = [];
  const add = (level, where, msg) => F.push({ level, where, msg });
  const ff = FORM_FACTORS[project.app.formFactor || 'tablet'];
  // names
  const seen = new Map();
  for (const scr of project.screens) {
    const reg = (n, where) => { if (seen.has(n)) add('error', where, `Name "${n}" is used twice (also in ${seen.get(n)}). Names must be unique across the whole app (PA2110).`); else seen.set(n, where); };
    if (!IDENT.test(scr.name) || RESERVED.has(scr.name)) add('error', scr.name, `Screen name "${scr.name}" is not a valid Power Apps name.`);
    reg(scr.name, `screen ${scr.name}`);
    for (const n of emittedNames(project, scr)) reg(n, scr.name);
    for (const [c, parent] of walk(scr.controls)) {
      const where = `${scr.name} › ${c.name}`;
      if (!IDENT.test(c.name) || RESERVED.has(c.name)) add('error', where, `"${c.name}" is not a valid control name (letters, digits, underscore; not a reserved word).`);
      const fxAll = Object.values(c.fx || {}).join('\n') + '\n' + Object.values(c.actions || {}).flat().filter((a) => a.type === 'custom').map((a) => a.fx).join('\n');
      if (!parent && c.type !== 'gallery' && /\bThisItem\b/.test(fxAll)) add('error', where, 'Uses ThisItem outside a gallery. Move it into a gallery template or reference a variable.');
      if (/IfError\(\s*Patch\([\s\S]*\),\s*true\s*\)/.test(fxAll)) add('error', where, 'IfError(Patch(...), true): the fallback must be a record. Use Blank() or the 3-argument shape.');
      if (/Set\(\s*col\w*\s*,\s*Blank\(\)\s*\)/.test(fxAll)) add('error', where, 'Set(collection, Blank()) strips the type. Use Clear(col).');
      if (/\bin\s+\w+\.Selected\.\w+/.test(fxAll)) add('warn', where, '"x in Combo.Selected.Field" is a substring test on text. Use = for an exact match.');
      if (c.field?.list && !colOf(project, c.field.list, c.field.column)) add('error', where, `Bound to ${c.field.list}.${c.field.column}, which is not in the data model.`);
      if (c.bind?.list && c.bind.kind !== 'user' && c.bind.field && !colOf(project, c.bind.list, c.bind.field)) add('error', where, `Shows ${c.bind.list}.${c.bind.field}, which is not in the data model.`);
      if (c.bind?.src === 'item' && !parent) add('error', where, 'Shows a gallery row value but is not inside a gallery.');
      if (c.type === 'gallery') {
        if (!c.data?.list && !c.fx?.Items) add('warn', where, 'Gallery has no data source. Pick a list on the Data tab.');
        else if (c.data?.list && !project.lists.some((l) => l.name === c.data.list)) add('error', where, `Gallery list "${c.data.list}" is not in the data model.`);
        if (!(c.children || []).length) add('warn', where, 'Gallery template is empty. Drop controls into the first row.');
        if (c.data?.searchCtl && !seen.has(c.data.searchCtl) && !allNames(project).has(c.data.searchCtl)) add('error', where, `Search box "${c.data.searchCtl}" does not exist.`);
      }
      if (c.type === 'person' && parent) add('error', where, 'People pickers cannot sit inside a gallery.');
      if (c.type === 'image' && (c.props.src || '').length > 300000) add('warn', where, 'Embedded image is large (>300 KB); the paste may be slow. Consider adding it as app media instead.');
      if (!parent && (c.x < 0 || c.y < 0 || c.x + c.w > ff.w + 1 || c.y + c.h > ff.h + 1)) add('warn', where, 'Sits partly outside the screen.');
      if (parent && (c.y + c.h > (parent.props.templateSize || 0) + 1 || c.x + c.w > parent.w + 1)) add('warn', where, 'Extends past the gallery row; it will be clipped.');
      for (const [ev, acts] of Object.entries(c.actions || {})) for (const a of acts || []) {
        if (['navigate'].includes(a.type) && !scrName(project, a.screen)) add('error', where, `${ev}: "Go to screen" has no valid target.`);
        if (['newItem', 'openItem'].includes(a.type) && a.screen && !scrName(project, a.screen)) add('error', where, `${ev}: target screen is missing.`);
        if (['save', 'delete', 'setField', 'newItem', 'openItem', 'refresh'].includes(a.type) && (!a.list || !project.lists.some((l) => l.name === a.list))) add('error', where, `${ev}: "${ACTION_TYPES[a.type]}" needs a list from the data model.`);
        if (a.type === 'openItem' && !parent && c.type !== 'gallery') add('error', where, `${ev}: "Open this row" only works on a gallery or inside one.`);
        if (a.type === 'save' && !screenInputs(scr, a.list).length) add('warn', where, `${ev}: Save has no inputs bound to ${a.list} on this screen.`);
        if (a.type === 'save' && a.list) {
          const bound = new Set(screenInputs(scr, a.list).map((x) => x.field.column));
          const l = project.lists.find((x) => x.name === a.list);
          l?.columns.filter((col) => col.required && !bound.has(col.name)).forEach((col) => add('warn', where, `Required column ${a.list}.${col.name} has no input on this screen; SharePoint will reject the save.`));
        }
        if ((a.type === 'setVar' || a.type === 'toggleVar') && a.name && !IDENT.test(a.name)) add('error', where, `Variable name "${a.name}" is not valid.`);
        if (a.type === 'custom' && /\bNavigate\(/.test(a.fx || '') && ev === 'OnVisible') add('error', where, 'Navigate() is not allowed in OnVisible.');
      }
    }
    if (/\bNavigate\(/.test(scr.onVisible || '')) add('error', scr.name, 'Navigate() is not allowed in a screen\'s OnVisible (even inside If). Hide the button that leads here instead.');
    if (/\bThisItem\b/.test(scr.onVisible || '')) add('error', scr.name, 'ThisItem in OnVisible.');
  }
  if (/\bNavigate\(/.test(project.app.onStart || '')) add('error', 'App.OnStart', 'Navigate() is not allowed in App.OnStart. Set the start screen instead.');
  // data model
  const lnames = new Set();
  for (const l of project.lists) {
    if (lnames.has(l.name)) add('error', `list ${l.name}`, 'Two lists share this name.');
    lnames.add(l.name);
    const cn = new Set();
    for (const c of l.columns) {
      if (cn.has(c.name)) add('error', `${l.name}.${c.name}`, 'Duplicate column name.');
      cn.add(c.name);
      if (!c.builtin && SP_RESERVED_COLS.has(c.name)) add('warn', `${l.name}.${c.name}`, `SharePoint already uses "${c.name}". Prefix it (e.g. Req${c.name}) to avoid a clash.`);
      if (!c.builtin && !IDENT.test(c.name)) add('warn', `${l.name}.${c.name}`, 'Spaces or symbols in column names make internal names differ from display names. Prefer one word.');
      if (c.type === 'Choice' && !c.choices.length) add('error', `${l.name}.${c.name}`, 'Choice column has no choices.');
      if (c.type === 'Lookup' && !project.lists.some((x) => x.name === c.lookupList)) add('error', `${l.name}.${c.name}`, 'Lookup column needs a target list from the data model.');
    }
  }
  for (const scr of project.screens) for (const [c] of walk(scr.controls)) if (c.bind?.kind === 'count') { add('info', `${scr.name} › ${c.name}`, 'CountRows over a SharePoint list is not delegable: counts stop at the data row limit (500 by default, 2000 max).'); break; }
  return F;
}
