// Power App Studio test harness. No build step and no dependencies beyond Node 18+ and the vendored js-yaml.
// Run from the repo root or this folder:  node powerapp-studio/tests/run.mjs
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import path from 'path';

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
globalThis.jsyaml = require(path.join(here, '../vendor/js-yaml.min.js'));
const yaml = globalThis.jsyaml;

const core = await import('../js/core.js');
const { SCENARIOS, generateApp, genEdit } = await import('../js/templates.js');
const E = await import('../js/emit.js');
const FX = await import('../js/fx.js');
const { importYaml } = await import('../js/import.js');
const { createRuntime } = await import('../js/render.js');
const FL = await import('../js/flows.js');
const AI = await import('../js/ai.js');
const DP = await import('../js/deploy.js');
const SU = await import('../js/setup.js');
const GD = await import('../js/guide.js');
const { buildDemo } = await import('../js/templates.js');
const fs = await import('node:fs');
const FIXTURE = JSON.parse(fs.readFileSync(new URL('./fixture-plan.json', import.meta.url), 'utf8'));
const CHANGE_FIXTURE = JSON.parse(fs.readFileSync(new URL('./fixture-change.json', import.meta.url), 'utf8'));
const CH = await import('../js/change.js');
const ED = await import('../js/edit.js');
const SH = await import('../js/share.js');
const TB = await import('../js/tabular.js');
const A11 = await import('../js/a11y.js');
const FB = await import('../js/feedback.js');
const BU = await import('../js/bundle.js');
const HI = await import('../js/history.js');
const DG = await import('../js/delegation.js');
const JSZip = require(path.join(here, '../vendor/jszip.min.js'));

// ---------------------------------------------------------------- tiny runner
let pass = 0, fail = 0;
const failures = [];
const pending = [];
function test(name, fn) {
  const ok = () => { pass++; process.stdout.write('.'); };
  const bad = (e) => { fail++; failures.push(`✗ ${name}\n    ${e.message}`); process.stdout.write('F'); };
  try { const r = fn(); if (r && typeof r.then === 'function') pending.push(r.then(ok, bad)); else ok(); }
  catch (e) { bad(e); }
}
function assert(cond, msg) { if (!cond) throw new Error(msg); }
const eq = (a, b, msg) => assert(a === b, `${msg}: expected ${JSON.stringify(b)}, got ${JSON.stringify(a)}`);

// ---------------------------------------------------------------- helpers
function build(scenario, formFactor = 'tablet', screens = ['home', 'browse', 'detail', 'edit']) {
  const sc = SCENARIOS[scenario];
  const p = core.newProject(); p.app.name = sc.app; p.app.formFactor = formFactor; p.lists = sc.lists();
  generateApp(p, { screens, approvals: sc.approvals || null });
  return p;
}
const errors = (p) => E.lint(p).filter((f) => f.level === 'error');
const codes = (p) => new Set(E.lint(p).map((f) => f.code));

/** Parse one screen's YAML and check every paste rule we can check statically. Returns the parsed screen. */
function checkScreenYaml(text, names) {
  // plain scalars must not contain YAML-significant sequences (those need a |- block)
  for (const line of text.split('\n')) {
    const m = line.match(/^\s+[A-Za-z]+: (=.*)$/);
    if (!m) continue;
    const v = m[1];
    assert(!/: | #|[{}\t]|:$|\s$/.test(v), `plain scalar needs a block: ${line.trim()}`);
  }
  const doc = yaml.load(text, { schema: yaml.FAILSAFE_SCHEMA });
  assert(doc && doc.Screens, 'top-level key must be Screens:');
  const entries = Object.entries(doc.Screens);
  eq(entries.length, 1, 'one screen per paste file');
  const [sn, scr] = entries[0];
  assert(!names.has(sn), `duplicate name ${sn}`); names.add(sn);
  const walk = (children, inGallery) => {
    for (const item of children || []) {
      const [[n, c]] = Object.entries(item);
      assert(!names.has(n), `duplicate control name ${n}`); names.add(n);
      const props = c.Properties || {};
      if (String(c.Control).startsWith('Gallery')) {
        eq(c.Control, 'Gallery@2.15.0', `${n} gallery pin`); eq(c.Variant, 'Vertical', `${n} gallery variant`);
        assert((c.Children || []).length > 0, `${n}: gallery template is empty (wrong indent?)`);
        walk(c.Children, true);
      } else {
        assert(!String(c.Control).includes('@'), `${n}: version pin on ${c.Control}`);
        assert(!c.Children, `${n}: children on a non-gallery`);
      }
      if (c.Control === 'Rectangle') assert(!Object.keys(props).some((k) => k.startsWith('Radius')), `${n}: Rectangle has Radius (PA2108)`);
      for (const [k, v] of Object.entries(props)) {
        assert(String(v).startsWith('='), `${n}.${k} must start with =`);
        const own = String(c.Control).startsWith('Gallery') && (k === 'TemplateFill' || k === 'Items');
        if (!inGallery && !own) assert(!/\bThisItem\b/.test(v), `${n}.${k}: ThisItem outside a gallery`);
        assert(!(k === 'OnStart' || k === 'OnVisible') || !/\bNavigate\(/.test(v), `${n}.${k}: Navigate`);
      }
    }
  };
  walk(scr.Children, false);
  if (scr.Properties?.OnVisible) assert(!/\bNavigate\(/.test(scr.Properties.OnVisible), `${sn}: Navigate in OnVisible`);
  return { name: sn, scr };
}
function checkApp(p) {
  const names = new Set();
  const out = {};
  for (const s of p.screens) out[s.name] = checkScreenYaml(E.screenYaml(p, s), names).scr;
  return out;
}

// ---------------------------------------------------------------- tests
for (const k of Object.keys(SCENARIOS)) for (const ff of ['tablet', 'phone']) {
  test(`scenario ${k} (${ff}) generates lint-clean, paste-safe YAML`, () => {
    const p = build(k, ff);
    const errs = errors(p);
    assert(!errs.length, errs.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
    assert(!E.lint(p).some((f) => f.level === 'warn'), 'unexpected warnings: ' + E.lint(p).filter((f) => f.level === 'warn').map((f) => f.code + ' ' + f.where).join(', '));
    checkApp(p);
    assert(E.themeFormulas(p).startsWith('// Theme'), 'theme formulas');
    assert(!/\bNavigate\(/.test(E.onStartFx(p)), 'Navigate in OnStart');
  });
}

test('tricky literals survive YAML exactly (colon-space, hash, quotes, braces)', () => {
  const p = build('tasks');
  const s = p.screens[0];
  const add = (t, o) => { const c = core.newControl(p, t, o); s.controls.push(c); return c; };
  add('dropdown', { name: 'ddTricky', x: 10, y: 10, props: { items: 'A: one\nB #two\n"C"', size: 12 } });
  add('label', { name: 'lblTricky', x: 10, y: 60, props: { text: 'Next: step #2 {x}' } });
  add('button', { name: 'btnTricky', x: 10, y: 120, props: { text: 'Go' }, actions: { OnSelect: [{ type: 'custom', fx: 'Set(varX, {a: 1, b: "c: d"})' }] } });
  const names = new Set();
  const { scr } = checkScreenYaml(E.screenYaml(p, s), names);
  const get = (n) => scr.Children.map((x) => Object.entries(x)[0]).find(([k]) => k === n)[1].Properties;
  eq(get('ddTricky').Items, '=["A: one", "B #two", """C"""]', 'dropdown Items');
  eq(get('lblTricky').Text, '="Next: step #2 {x}"', 'label Text');
  eq(get('btnTricky').OnSelect, '=Set(varX, {a: 1, b: "c: d"})', 'custom OnSelect');
});

test('Windows line endings in formulas are normalized', () => {
  const p = build('tasks'); const s = p.screens[0];
  s.controls.push(core.newControl(p, 'button', { name: 'btnCrlf', actions: { OnSelect: [{ type: 'custom', fx: 'Set(a, 1);\r\nSet(b, 2)' }] } }));
  assert(!E.screenYaml(p, s).includes('\r'), 'CR leaked into YAML');
});

test('gallery row tap sits behind clickable row controls', () => {
  const p = build('tasks');
  const scr = p.screens.find((x) => x.name.startsWith('scrBrowse'));
  const gal = scr.controls.find((c) => c.type === 'gallery');
  gal.children.push(core.newControl(p, 'icon', { name: 'icoDel', x: 10, y: 10, w: 24, h: 24, props: { icon: 'trash', color: '@Danger', stroke: 2 }, actions: { OnSelect: [{ type: 'delete', list: 'Tasks', then: '' }] } }));
  const doc = yaml.load(E.screenYaml(p, scr), { schema: yaml.FAILSAFE_SCHEMA });
  const g = Object.values(doc.Screens)[0].Children.map((x) => Object.entries(x)[0]).find(([, c]) => String(c.Control).startsWith('Gallery'))[1];
  const order = g.Children.map((x) => Object.keys(x)[0]);
  assert(order.indexOf(gal.name + 'RowTap') < order.indexOf('icoDelTap'), 'row tap must come before the icon tap: ' + order.join(','));
  const passive = gal.children.filter((c) => !c.actions?.OnSelect?.length).map((c) => c.name);
  assert(passive.every((n) => order.indexOf(n) < order.indexOf(gal.name + 'RowTap')), 'passive row content stays behind the row tap');
});

test('tap buttons follow a Power Fx Visible override', () => {
  const p = build('tasks'); const s = p.screens[0];
  s.controls.push(core.newControl(p, 'card', { name: 'cardClick', x: 10, y: 600, fx: { Visible: 'varOpen' }, actions: { OnSelect: [{ type: 'back' }] } }));
  const y = E.screenYaml(p, s);
  const tap = y.slice(y.indexOf('- cardClickTap:'));
  assert(/Visible: =varOpen/.test(tap.slice(0, tap.indexOf('OnSelect'))), 'cardClickTap must be hidden with its card');
});

test('fx builders produce the field-tested shapes', () => {
  eq(FX.ifErrorFx('Patch(L, b, {})', 'Save', ['Notify("ok")', 'Back()']),
    'IfError(\n    Patch(L, b, {}),\n    Notify("Save failed. " & FirstError.Message, NotificationType.Error),\n    Notify("ok");\n    Back()\n)', 'IfError');
  eq(FX.upsertBaseFx('Tasks', 'varSelTasks'), 'Coalesce(LookUp(Tasks, ID = varSelTasks.ID), Defaults(Tasks))', 'upsert base');
  assert(FX.meUserFx().includes('Claims: "i:0#.f|membership|" & Lower(User().Email)'), 'claims');
  eq(FX.equalsFx('Status', 'Choice', 'Done'), 'Status.Value = "Done"', 'choice equals');
  eq(FX.equalsFx('Flag', 'YesNo', 'yes'), 'Flag = true', 'yes/no equals');
  eq(FX.recordFx([]), '{}', 'empty record');
});

test('lint rules fire on known-bad projects', () => {
  const p = build('requests');
  const home = p.screens.find((x) => x.name === 'scrHome');
  const edit = p.screens.find((x) => x.name.startsWith('scrEdit'));
  home.onVisible = 'If(true, Navigate(scrHome))';
  home.controls.push(core.newControl(p, 'label', { name: 'lblBadItem', fx: { Text: 'ThisItem.Title' } }));
  home.controls.push(core.newControl(p, 'html', { name: 'htmlStyle', props: { html: '<style>p{color:red}</style><p>x</p>' } }));
  home.controls.push(core.newControl(p, 'label', { name: 'lblHidden', visibleWhen: { var: 'varNeverSet', op: 'true' } }));
  home.controls.push(core.newControl(p, 'button', { name: 'btnCovered', x: 900, y: 650, w: 100, h: 40, actions: { OnSelect: [{ type: 'back' }] } }));
  home.controls.push(core.newControl(p, 'label', { name: 'lblCover', x: 900, y: 650, w: 100, h: 40 }));
  edit.controls.push(core.newControl(p, 'textinput', { name: 'txtLoose', x: 40, y: 700 }));
  home.controls.push(core.newControl(p, 'label', { name: 'lblDup' }));
  edit.controls.push(core.newControl(p, 'label', { name: 'lblDup' }));
  const c = codes(p);
  for (const code of ['PS001', 'PS004', 'PS026', 'PS036', 'PS037', 'PS040', 'PS041']) assert(c.has(code), `${code} did not fire`);
});

test('lint-ok comment silences one rule on one control', () => {
  const p = build('tasks'); const s = p.screens[0];
  s.controls.push(core.newControl(p, 'button', { name: 'btnOk', x: 900, y: 650, w: 100, h: 40, actions: { OnSelect: [{ type: 'custom', fx: 'Back() // lint-ok PS040' }] } }));
  s.controls.push(core.newControl(p, 'label', { name: 'lblOver', x: 900, y: 650, w: 100, h: 40 }));
  assert(!codes(p).has('PS040'), 'PS040 should be silenced');
});

test('copying a screen gives every copied control a unique name', () => {
  const p = build('requests');
  const src = p.screens.find((x) => x.name.startsWith('scrBrowse'));
  const copy = core.clone(src); copy.id = core.uid(); copy.name = core.uniqueScreenName(p, src.name + 'Copy');
  const used = core.allNames(p); used.add(copy.name);
  copy.controls.forEach((c) => core.renameCopy(p, c, used));
  p.screens.push(copy);
  assert(!codes(p).has('PS001'), 'duplicate names after copy: ' + E.lint(p).filter((f) => f.code === 'PS001').map((f) => f.msg).join('; '));
});

test('normalizeProject neutralizes hostile project files', () => {
  const p = build('tasks');
  const s0 = p.screens[0], s1 = p.screens[1];
  const oldId = s1.id; s1.id = '"><img src=x onerror=alert(1)>';
  s0.controls[0].actions.OnSelect = [{ type: 'navigate', screen: s1.id }];
  s0.controls[0].props.fill = 'red;background:url(javascript:alert(1))';
  p.theme.colors.Primary = 'red" onmouseover="x';
  p.app.startScreen = s1.id;
  core.normalizeProject(p);
  assert(/^[A-Za-z0-9_-]+$/.test(s1.id) && s1.id !== oldId, 'screen id sanitized');
  eq(s0.controls[0].actions.OnSelect[0].screen, s1.id, 'action re-pointed to new id');
  eq(p.app.startScreen, s1.id, 'start screen re-pointed');
  eq(s0.controls[0].props.fill, 'transparent', 'bad color dropped');
  assert(core.isHex(p.theme.colors.Primary), 'theme color reset');
  eq(core.cssColor('red;x', p.theme), 'transparent', 'cssColor refuses non-colors');
});

test('YAML import round-trips a generated screen', () => {
  const p = build('requests');
  const src = p.screens.find((x) => x.name.startsWith('scrBrowse'));
  const r = importYaml(p, E.screenYaml(p, src));
  eq(r.screens.length, 1, 'one screen imported');
  assert(Object.keys(r.renamed).length > 0, 'clashing names renamed');
  p.screens.push(...r.screens);
  assert(!codes(p).has('PS001'), 'no duplicate names after import');
  checkApp(p);
  const gal = r.screens[0].controls.find((c) => c.type === 'gallery');
  eq(gal.data.previewList, 'Requests', 'gallery preview list detected');
});

test('tall tablet forms switch to three columns and stay on screen', () => {
  const p = core.newProject();
  p.lists = [core.newList('Wide', Array.from({ length: 16 }, (_, i) => core.newColumn('Field' + (i + 1), 'Text')))];
  const scr = genEdit(p, p.lists[0]);
  const inputs = scr.controls.filter((c) => c.type === 'textinput');
  eq(new Set(inputs.map((c) => c.x)).size, 3, 'three columns');
  assert(!codes(p).has('PS017'), 'form must fit the screen');
});

test('Play runtime: new item saves, required fields block saving', () => {
  const p = build('requests');
  const toasts = [];
  const rt = createRuntime(p, { toast: (t, k) => toasts.push([t, k]), rerender: () => {} });
  const edit = p.screens.find((x) => x.name.startsWith('scrEdit'));
  const save = edit.controls.find((c) => c.name.startsWith('btnSave'));
  const before = rt.db.Requests.length;
  rt.run(edit, [{ type: 'newItem', list: 'Requests', screen: edit.id }]);
  rt.run(edit, save.actions.OnSelect);
  eq(rt.db.Requests.length, before, 'blank Title must not save');
  assert(toasts.some(([t, k]) => (k === 'error' || k === 'warning') && /Title/.test(t)), 'toast names the missing field');
  assert(rt.vars.varTriedSaveRequests === true, 'validation state set (red borders + still needed)');
  const title = edit.controls.find((c) => c.field?.column === 'Title');
  rt.inputs[title.name] = 'Harness item';
  rt.run(edit, save.actions.OnSelect);
  eq(rt.db.Requests.length, before + 1, 'item saved');
  eq(rt.db.Requests[0].Title, 'Harness item', 'saved title');
});

test('provisioning script creates lookup targets first', () => {
  const p = build('events');
  const ps = E.provisioningScript(p);
  assert(ps.indexOf("Ensure-List 'Events'") < ps.indexOf("Ensure-List 'Registrations'"), 'Events before Registrations');
  assert(ps.includes('Add-PnPFieldFromXml'), 'uses field XML');
});

// ---------------------------------------------------------------- flows
function withFlows() {
  const p = build('requests');
  p.app.siteUrl = 'https://contoso.sharepoint.com/sites/Team/';
  p.app.appLink = 'https://apps.powerapps.com/play/e/env/a/app?tenantId=t';
  p.app.deepLink = { screen: p.screens.find((s) => s.name === 'scrDetailRequests').id, list: 'Requests' };
  p.flows = ['created', 'statusChanged', 'reminder', 'button'].map((k) => FL.newFlow(p, k, 'Requests'));
  return p;
}
/** Every @{...} interpolation and @expression must have balanced parentheses and quotes. */
function checkExpressions(node, path = '') {
  if (typeof node === 'string') {
    for (const m of node.matchAll(/@\{([^}]*(?:\{[^}]*\}[^}]*)*)\}/g)) {
      const e = m[1];
      let depth = 0, q = false;
      for (const ch of e) { if (ch === "'") q = !q; else if (!q && ch === '(') depth++; else if (!q && ch === ')') depth--; }
      assert(depth === 0 && !q, `unbalanced expression at ${path}: ${e}`);
    }
    return;
  }
  if (node && typeof node === 'object') for (const [k, v] of Object.entries(node)) checkExpressions(v, path + '/' + k);
}

test('flow packages have the legacy import layout', () => {
  const p = withFlows();
  for (const f of p.flows) {
    const files = FL.flowPackageFiles(p, f);
    const man = JSON.parse(files['manifest.json']);
    const fm = JSON.parse(files['Microsoft.Flow/flows/manifest.json']);
    const key = fm.flowAssets.assetPaths[0];
    assert(man.resources[key]?.type === 'Microsoft.Flow/flows', `${f.kind}: flow resource keyed by asset path`);
    eq(man.resources[key].hierarchy, 'Root', `${f.kind}: root resource`);
    const def = JSON.parse(files[`Microsoft.Flow/flows/${key}/definition.json`]);
    const apis = JSON.parse(files[`Microsoft.Flow/flows/${key}/apisMap.json`]);
    const conns = JSON.parse(files[`Microsoft.Flow/flows/${key}/connectionsMap.json`]);
    for (const c of ['shared_sharepointonline', 'shared_office365']) {
      assert(man.resources[apis[c]]?.type === 'Microsoft.PowerApps/apis', `${f.kind}: api resource for ${c}`);
      assert(man.resources[conns[c]]?.type === 'Microsoft.PowerApps/apis/connections', `${f.kind}: connection resource for ${c}`);
      assert(man.resources[conns[c]].dependsOn.includes(apis[c]), `${f.kind}: connection depends on its api`);
      assert(man.resources[key].dependsOn.includes(conns[c]), `${f.kind}: flow depends on connection`);
      assert(def.properties.connectionReferences[c], `${f.kind}: connection reference ${c}`);
    }
    eq(def.properties.displayName, f.name, 'display name');
    assert(Object.keys(def.properties.definition.triggers).length === 1, 'one trigger');
    checkExpressions(def.properties.definition);
  }
});

test('flow expressions use internal names and the right sources', () => {
  const p = withFlows();
  p.lists[0].columns.push(core.newColumn('Needed Date', 'DateTime'));
  const [created, changed, reminder, button] = p.flows;
  created.body = '{NeededDate} {Approver.Email} {Status} {Link} {Nope}';
  const def = JSON.stringify(FL.flowDefinition(p, created));
  assert(def.includes("body/NeededDate"), 'internal name without spaces');
  assert(def.includes("body/Approver/Email"), 'person email');
  assert(def.includes("body/Status/Value"), 'choice value');
  assert(def.includes('&ID='), 'deep link');
  assert(def.includes('{Nope}'), 'unknown token left as text');
  assert(E.lint(p).some((x) => x.code === 'PS048'), 'unknown token flagged');
  const rd = JSON.stringify(FL.flowDefinition(p, reminder));
  assert(rd.includes("items('Apply_to_each')"), 'reminder reads loop item');
  assert(rd.includes("Status eq 'Submitted'"), 'reminder OData filter');
  const bd = JSON.stringify(FL.flowDefinition(p, button));
  assert(bd.includes("outputs('Get_item')") && bd.includes('PowerAppV2') && bd.includes('Respond_to_a_PowerApp_or_flow'), 'button flow shape');
  const cd = FL.flowDefinition(p, changed);
  assert(JSON.stringify(cd.actions.Condition.expression).includes("ColumnHasChanged']?['Status']"), 'status change check');
});

test('literal @ stays literal, @{ is escaped', () => {
  const p = withFlows(), f = p.flows[0];
  f.recipients = [{ kind: 'address', value: 'team@agency.mil' }];
  f.subject = '@home @{not a token}';
  eq(FL.recipientsExpr(p, f), 'team@agency.mil', 'email address untouched');
  eq(FL.interpolate(p, f, f.subject), '@@home @@{not a token}', 'leading @ and @{ escaped');
});

test('run-a-flow action, deep link start screen and OnStart', () => {
  const p = withFlows(), btn = p.flows[3];
  const det = p.screens.find((s) => s.name === 'scrDetailRequests');
  det.controls.push(core.newControl(p, 'button', { x: 40, y: 600, actions: { OnSelect: [{ type: 'runFlow', flow: btn.id, message: 'Sent' }] } }));
  const y = E.screenYaml(p, det);
  assert(y.includes(`${FL.flowFxName(btn.name)}.Run(varSelRequests.ID)`), 'runs the flow with the selected ID');
  assert(y.includes('Param("ID")'), 'deep-link OnVisible');
  eq(E.startScreenFx(p), 'If(IsBlank(Param("ID")), scrHome, scrDetailRequests)', 'start screen formula');
  assert(E.onStartFx(p).includes('LookUp(Requests, ID = Value(Param("ID")))'), 'OnStart does not clobber the deep-linked item');
  eq(errors(p).length, 0, 'no lint errors: ' + errors(p).map((x) => x.code + ' ' + x.msg).join('; '));
  checkScreenYaml(y, new Set());
  det.controls.at(-1).actions.OnSelect[0].flow = 'missing';
  assert(codes(p).has('PS049'), 'missing flow flagged');
});

test('flow lint: site URL, recipients, watch column', () => {
  const p = withFlows();
  p.app.siteUrl = ''; p.flows[0].recipients = []; p.flows[1].statusField = 'Title';
  const c = codes(p);
  assert(c.has('PS045') && c.has('PS044') && c.has('PS046'), 'flags ' + [...c].join(','));
});

test('Play: saving and approving report the flows that would run', () => {
  const p = withFlows(), msgs = [];
  const rt = createRuntime(p, { toast: (t) => msgs.push(t), rerender: () => {} });
  const edit = p.screens.find((s) => s.name === 'scrEditRequests');
  rt.vars.varSelRequests = {};
  const title = [...core.walk(edit.controls)].map(([c]) => c).find((c) => c.field?.column === 'Title');
  rt.inputs[title.name] = 'Flow item';
  const save = [...core.walk(edit.controls)].map(([c]) => c).find((c) => c.actions?.OnSelect?.some((a) => a.type === 'save'));
  rt.run(edit, save.actions.OnSelect);
  assert(msgs.some((m) => m.includes(p.flows[0].name)), 'created flow reported: ' + msgs.join(' | '));
  const appr = p.screens.find((s) => s.name === 'scrApprovalsRequests');
  rt.vars.varSelRequests = rt.db.Requests[0];
  const approve = [...core.walk(appr.controls)].map(([c]) => c).find((c) => c.name.startsWith('btnApprove'));
  rt.run(appr, approve.actions.OnSelect);
  assert(msgs.some((m) => m.includes(p.flows[1].name)), 'status flow reported: ' + msgs.join(' | '));
});

test('old project files without flows still load', () => {
  const p = build('tasks'); delete p.flows; delete p.app.siteUrl;
  core.normalizeProject(p);
  assert(Array.isArray(p.flows) && p.app.siteUrl === '', 'defaults added');
  eq(errors(p).length, 0, 'still lint-clean');
});

// ---------------------------------------------------------------- describe with Claude
test('plan schema stays inside structured-output limits', () => {
  const bad = ['minimum', 'maximum', 'maxItems', 'minLength', 'maxLength', 'oneOf', 'not', 'pattern'];
  const walkS = (s, path) => {
    for (const k of bad) assert(!(k in s), `${path} uses unsupported ${k}`);
    if (s.type === 'object') {
      eq(s.additionalProperties, false, path + ' additionalProperties');
      eq(JSON.stringify([...s.required].sort()), JSON.stringify(Object.keys(s.properties).sort()), path + ' requires every property');
      for (const [k, v] of Object.entries(s.properties)) walkS(v, path + '.' + k);
    }
    if (s.type === 'array') walkS(s.items, path + '[]');
  };
  walkS(AI.PLAN_SCHEMA, 'plan');
  const req = AI.buildRequest('An app', 'claude-sonnet-5-5');
  eq(req.output_config.format.type, 'json_schema', 'structured output');
  assert(!('temperature' in req), 'no temperature (removed in API 2026-10-21)');
});

test('callClaude sends browser-access headers and reads the plan', async () => {
  let seen = null;
  const fetchFn = async (url, init) => { seen = { url, init }; return { ok: true, status: 200, json: async () => ({ stop_reason: 'end_turn', content: [{ type: 'thinking', thinking: '…' }, { type: 'text', text: JSON.stringify(FIXTURE) }], usage: { input_tokens: 10, output_tokens: 20 } }) }; };
  const r = await AI.callClaude({ key: 'sk-ant-test', model: 'claude-sonnet-5-5', text: 'x', fetchFn });
  eq(seen.url, 'https://api.anthropic.com/v1/messages', 'endpoint');
  eq(seen.init.headers['anthropic-dangerous-direct-browser-access'], 'true', 'CORS opt-in');
  eq(seen.init.headers['anthropic-version'], '2026-10-21', 'version');
  eq(r.plan.appName, 'Training Requests', 'plan parsed past the thinking block');
  let msg = '';
  try { await AI.callClaude({ key: 'sk-ant-x', text: 'x', fetchFn: async () => ({ ok: false, status: 401, json: async () => ({ error: { message: 'invalid x-api-key' } }) }) }); } catch (e) { msg = e.message; }
  assert(/rejected/.test(msg), '401 explained: ' + msg);
  msg = '';
  try { await AI.callClaude({ key: 'nope', text: 'x', fetchFn }); } catch (e) { msg = e.message; }
  assert(/sk-ant-/.test(msg), 'bad key caught before sending');
});

test('a messy plan becomes a clean, lint-free project', () => {
  const { project: p, notes } = AI.specToProject(FIXTURE, { siteUrl: 'https://contoso.sharepoint.com/sites/Team' });
  eq(p.lists[0].name, 'TrainingRequests', 'list name without spaces');
  const names = p.lists[0].columns.map((c) => c.name);
  assert(names.includes('CourseCost') && !names.includes('Course Cost'), 'column names cleaned');
  assert(names.filter((n) => n === 'Title').length === 1, 'Title not duplicated');
  assert(names.includes('TrainingRequestStartDate') && !names.includes('StartDate'), 'reserved StartDate renamed');
  eq(p.lists[0].columns.find((c) => c.name === 'Format').type, 'Text', 'choice without choices → text');
  eq(p.lists[0].columns.find((c) => c.name === 'Vendor').type, 'Text', 'lookup to a missing list → text');
  assert(p.screens.some((s) => s.name === 'scrApprovalsTrainingRequests'), 'approvals screen');
  const [created, changed, reminder, button] = p.flows;
  assert(created.body.includes('{CourseCost}'), 'tokens follow renamed columns: ' + created.body);
  eq(JSON.stringify(changed.statusValues), '["Approved","Rejected"]', 'unknown status values dropped');
  eq(changed.recipients.length, 1, 'non-person column recipient dropped');
  eq(reminder.hour, 23, 'hour clamped');
  eq(reminder.recipients.length, 1, 'bad address dropped');
  eq(button.name, 'SendSummaryToMe', 'button flow name is an identifier');
  const det = p.screens.find((s) => s.name === 'scrDetailTrainingRequests');
  assert(det.controls.some((c) => c.actions?.OnSelect?.some((a) => a.type === 'runFlow' && a.flow === button.id)), 'button flow wired to a detail button');
  eq(p.app.deepLink?.screen, det.id, 'email links open the detail screen');
  assert(notes.length >= 4, 'fixes reported');
  eq(errors(p).length, 0, 'lint errors: ' + errors(p).map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
  const seen = new Set();
  for (const s of p.screens) checkScreenYaml(E.screenYaml(p, s), seen);
  for (const f of p.flows) checkExpressions(FL.flowDefinition(p, f));
});

test('an empty or broken plan still builds something usable', () => {
  const { project: p } = AI.specToProject({ lists: [], screens: {}, flows: [{ kind: 'nonsense' }] });
  assert(p.lists.length === 1 && p.screens.length > 0, 'fallback list and screens');
  eq(errors(p).length, 0, 'lint-clean');
  let msg = '';
  try { AI.parseResponse({ stop_reason: 'max_tokens', content: [] }); } catch (e) { msg = e.message; }
  assert(/cut off/.test(msg), 'truncation explained');
});

// ---------------------------------------------------------------- premium generation
const ALL_FT = { loading: true, usage: true, confirm: true, stepper: true, scope: 'mine', adminGroup: '11111111-2222-3333-4444-555555555555' };
function buildFt(scenario, formFactor = 'tablet', ft = ALL_FT) {
  const sc = SCENARIOS[scenario];
  const p = core.newProject(); p.app.name = sc.app; p.app.formFactor = formFactor; p.lists = sc.lists();
  p.app.features = { ...ft };
  generateApp(p, { screens: ['home', 'browse', 'detail', 'edit'], approvals: sc.approvals || null });
  return p;
}
/** Every HtmlViewer in the emitted YAML: the first fixed-height box plus its margins fits in Height - 10 (5px padding per side). */
function checkHtmlPadding(parsed) {
  let n = 0;
  const walkY = (children) => (children || []).forEach((item) => {
    const [[name, c]] = Object.entries(item);
    if (c.Control === 'HtmlViewer') {
      const h = Number(String(c.Properties?.Height || '').replace(/^=/, ''));
      const html = String(c.Properties?.HtmlText || '');
      assert(Number.isFinite(h), `${name}: HtmlViewer Height should be a number`);
      if (/height:\s*\d/.test(html)) { n++; eq(E.htmlOverflow(html, h), 0, `${name}: HTML box overflows Height ${h} - 10`); }
    }
    walkY(c.Children);
  });
  for (const scr of Object.values(parsed)) walkY(scr.Children);
  return n;
}

for (const k of Object.keys(SCENARIOS)) for (const ff of ['tablet', 'phone']) {
  test(`premium ${k} (${ff}) with every option on: lint-clean, paste-safe, padded HTML`, () => {
    const p = buildFt(k, ff);
    const bad = E.lint(p).filter((f) => f.level !== 'info');
    assert(!bad.length, bad.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
    const parsed = checkApp(p);
    assert(checkHtmlPadding(parsed) >= 3, 'card HTML present and checked');
    eq(p.screens[0].role, 'loading', 'loading screen first');
    eq(E.startScreenName(p), 'scrLoading', 'loading is the start screen');
    const load = parsed.scrLoading;
    const tmr = load.Children.map((x) => Object.entries(x)[0]).find(([n]) => n.startsWith('tmrLoad'));
    assert(tmr && /Navigate\(scrHome/.test(tmr[1].Properties.OnTimerEnd), 'loading timer navigates home in OnTimerEnd');
    assert(!/Navigate/.test(load.Properties?.OnVisible || ''), 'no Navigate in OnVisible');
    assert(/Office365Groups\.ListGroupMembers\("11111111-2222-3333-4444-555555555555"\)/.test(E.onStartFx(p)) && /IfError\(/.test(E.onStartFx(p)), 'admin check wrapped in IfError');
    assert(p.lists.some((l) => l.name === 'AppUsage'), 'AppUsage list provisioned');
    assert(/AppUsage/.test(tmr[1].Properties.OnTimerEnd), 'usage recorded on the loading screen');
    const yall = p.screens.map((s) => E.screenYaml(p, s)).join('\n');
    assert(/varScope\w+ = "mine"/.test(yall) && /'Created By'\.Email = User\(\)\.Email/.test(yall), 'My items filter');
    assert(/varTriedSave\w+ && /.test(yall), 'red borders / still needed');
    assert(/varDone\w+/.test(yall) && /Add another/.test(yall), 'confirmation page');
    assert(/varConfirmDelete\w+/.test(yall), 'two-click delete');
    assert(!/Radius/.test(yall.split('\n').filter((l) => /Rectangle/.test(l)).join('')), 'no Rectangle radius');
  });
}

test('options off: no loading screen, no admin check, no usage, no confirmation', () => {
  const p = buildFt('requests', 'tablet', { loading: false, usage: false, confirm: false, stepper: false, scope: 'all', adminGroup: '' });
  assert(!p.screens.some((s) => s.role === 'loading'), 'no loading screen');
  assert(!/Office365Groups/.test(E.onStartFx(p)), 'no admin check');
  assert(!p.lists.some((l) => l.name === 'AppUsage'), 'no AppUsage');
  const yall = p.screens.map((s) => E.screenYaml(p, s)).join('\n');
  assert(!/varDone/.test(yall), 'no confirmation overlay');
  assert(!E.lint(p).some((f) => f.level !== 'info'), 'lint-clean');
  checkHtmlPadding(checkApp(p));
});

test('theme FontCss has no single quotes (it is inlined into style=\'...\' attributes)', () => {
  const p = build('requests');
  for (const font of ['Segoe UI', 'Open Sans', 'Courier New', 'Georgia']) {
    p.theme.font = font;
    const line = E.themeFormulas(p).split('\n').find((l) => /FontCss:/.test(l));
    assert(line && !/'/.test(line), 'FontCss quote-free: ' + line);
  }
});

test('PS051 flags an HTML box taller than its HtmlViewer', () => {
  eq(E.htmlOverflow("<div style='margin:3px;height:84px'>x</div>", 100), 0, 'fits exactly');
  eq(E.htmlOverflow("<div style='margin:3px;height:90px'>x</div>", 100), 6, 'overflow by 6');
  eq(E.htmlOverflow("<div style='margin:4px 2px;height:82px'>x</div>", 100), 0, 'two-value margin');
  eq(E.htmlOverflow("<div style='height:91px'>x</div>", 100), 1, 'no margin');
  const p = build('requests');
  const scr = p.screens[0];
  scr.controls.push(core.newControl(p, 'html', { name: 'htmTooTall', x: 10, y: 10, w: 200, h: 60, props: { html: "\"<div style='margin:2px;height:60px'>x</div>\"" } }));
  assert(codes(p).has('PS051'), 'PS051 raised');
});

test('Play: validation, confirmation page, two-click delete, Mine filter', async () => {
  const p = buildFt('requests');
  const toasts = [];
  const rt = createRuntime(p, { toast: (t, k) => toasts.push([t, k]), rerender: () => {} });
  const edit = p.screens.find((x) => x.name === 'scrEditRequests');
  const save = edit.controls.find((c) => c.name.startsWith('btnSave'));
  rt.run(edit, [{ type: 'newItem', list: 'Requests', screen: edit.id }]);
  rt.run(edit, save.actions.OnSelect);
  assert(rt.vars.varTriedSaveRequests === true && toasts.at(-1)[1] === 'warning', 'blocked with a warning');
  const need = edit.controls.find((c) => c.name === 'lblNeededEditRequests');
  const R = await import('../js/render.js');
  eq(R.evalFx(need.fx.Text, { rt, project: p, screen: edit }).trim(), 'Still needed: Title', '"Still needed" names the blank field');
  assert(R.evalFx(need.fx.Visible, { rt, project: p, screen: edit }) === true, '"Still needed" visible');
  rt.inputs[edit.controls.find((c) => c.field?.column === 'Title').name] = 'Premium harness';
  const n = rt.db.Requests.length;
  rt.run(edit, save.actions.OnSelect);
  eq(rt.db.Requests.length, n + 1, 'saved');
  assert(rt.vars.varDoneRequests === true, 'confirmation page shown for a new item');
  assert(rt.vars.varTriedSaveRequests === false, 'validation state reset');
  eq(rt.db.Requests[0].Author?.DisplayName, core.PEOPLE[0].DisplayName, 'new item is mine');
  const det = p.screens.find((x) => x.name === 'scrDetailRequests');
  const del = det.controls.find((c) => c.name.startsWith('btnDelete'));
  rt.vars.varSelRequests = rt.db.Requests[0];
  rt.run(det, del.actions.OnSelect);
  eq(rt.db.Requests.length, n + 1, 'first click only arms the delete');
  assert(rt.vars.varConfirmDeleteRequests === true, 'armed');
  rt.run(det, del.actions.OnSelect);
  eq(rt.db.Requests.length, n, 'second click deletes');
});

test('Play: the loading screen timer steps through and opens Home', () => {
  const p = buildFt('requests');
  const rt = createRuntime(p, { toast: () => {}, rerender: () => {} });
  const load = p.screens[0];
  eq(rt.screen, load.id, 'starts on the loading screen');
  const tmr = load.controls.find((c) => c.type === 'timer');
  for (let i = 0; i < 12 && rt.screen === load.id; i++) rt.run(load, tmr.actions.OnTimerEnd);
  eq(p.screens.find((s) => s.id === rt.screen)?.name, 'scrHome', 'navigated home');
});

test('a plan\'s options reach the generated app', () => {
  const { project: p, notes } = AI.specToProject(FIXTURE);
  const ft = p.app.features;
  assert(ft.loading && ft.usage && ft.confirm && ft.stepper && ft.scope === 'mine', 'options applied: ' + JSON.stringify(ft));
  eq(ft.adminGroup, '', 'a bad group ID is dropped');
  assert(notes.some((x) => /not a group ID/.test(x)), 'and reported');
  assert(p.screens[0].role === 'loading', 'loading screen generated');
  const props = AI.PLAN_SCHEMA.properties.options;
  eq(props.additionalProperties, false, 'options closed');
  eq(JSON.stringify(props.required.slice().sort()), JSON.stringify(Object.keys(props.properties).sort()), 'all option properties required');
  const old = { ...FIXTURE }; delete old.options;
  const { project: q } = AI.specToProject(old);
  assert(q.app.features.loading === true && q.app.features.scope === 'all' && !q.app.features.usage, 'plans without options get sensible defaults');
});

test('old project files without features still load', () => {
  const p = core.normalizeProject({ app: { name: 'Old' }, screens: [], lists: [] });
  eq(JSON.stringify(p.app.features), JSON.stringify({ loading: false, usage: false, confirm: true, stepper: true, scope: 'all', adminGroup: '' }), 'defaults');
  const q = core.normalizeProject({ app: { name: 'X', features: { adminGroup: '"); Remove(x' , scope: 'evil' } }, screens: [], lists: [] });
  eq(q.app.features.adminGroup, '', 'non-GUID group ID dropped');
  eq(q.app.features.scope, 'all', 'unknown scope dropped');
});

// ---------------------------------------------------------------- round 2: paste assistant
test('deploy hash is stable and content-sensitive', () => {
  eq(DP.hash('abc'), DP.hash('abc'), 'same input, same hash');
  assert(DP.hash('abc') !== DP.hash('abd'), 'one character changes the hash');
  assert(/^[0-9a-f]{8}$/.test(DP.hash('')), '8 hex chars');
});

test('deploy record survives save/load and only the edited screen turns Changed', () => {
  const p = withFlows();
  const steps = DP.deploySteps(p);
  assert(steps.some((s) => s.kind === 'screen') && steps.some((s) => s.kind === 'flow') && steps.at(-1).key === 'test', 'screens, flows and the final test step');
  assert(steps.every((s) => s.kind === 'static' || s.text || s.flow), 'every artifact step has something to copy or download');
  eq(DP.pendingSteps(p, steps).length, steps.length, 'first deploy: everything pending');
  for (const s of steps) DP.markDeployed(p, s);
  const p2 = core.normalizeProject(JSON.parse(JSON.stringify(p)));
  eq(Object.keys(p2.deployed).length, steps.length, 'record survives a JSON round trip');
  eq(DP.pendingSteps(p2).map((s) => s.key).join(), 'test', 'nothing changed: only the final test is pending');
  const target = p2.screens.find((s) => s.name === 'scrBrowseRequests');
  target.controls.find((c) => c.type === 'label').props.text = 'Something new';
  const pend = DP.pendingSteps(p2).map((s) => s.key);
  eq(JSON.stringify(pend), JSON.stringify(['scr:' + target.id, 'test']), 'exactly that screen changed');
  p2.flows[0].subject = 'Changed subject';
  assert(DP.pendingSteps(p2).some((s) => s.key === 'flow:' + p2.flows[0].id), 'editing a flow flips that flow');
  const gone = p2.screens.pop();
  assert(DP.removedSinceDeploy(p2).some((r) => r.key === 'scr:' + gone.id && r.name === gone.name), 'removed screens are reported by name');
});

test('normalizeProject cleans a hostile or missing deploy record', () => {
  const p = build('tasks');
  p.deployed = { 'scr:<img onerror=1>': { hash: 'zz' }, formulas: { hash: 'nothex!', at: 5 }, onstart: 'x' };
  core.normalizeProject(p);
  eq(JSON.stringify(Object.keys(p.deployed)), '["formulas"]', 'bad keys dropped');
  eq(p.deployed.formulas.hash, '', 'bad hash blanked');
  const q = build('tasks'); delete q.deployed; core.normalizeProject(q);
  eq(JSON.stringify(q.deployed), '{}', 'missing record becomes empty');
});

// ---------------------------------------------------------------- round 2: finish setup
test('Finish setup clears PS045 after Describe without a site URL', () => {
  const { project: p } = AI.specToProject(FIXTURE, {});
  assert(codes(p).has('PS045'), 'PS045 before');
  const item = SU.setupItems(p).find((i) => i.id === 'siteUrl');
  assert(item && !item.done && item.codes.includes('PS045'), 'site URL item maps to PS045');
  SU.applySetup(p, 'siteUrl', 'not a url');
  assert(!SU.setupItems(p).find((i) => i.id === 'siteUrl').done, 'invalid URL not done');
  SU.applySetup(p, 'siteUrl', 'https://contoso.sharepoint.com/sites/Team');
  assert(!codes(p).has('PS045'), 'every PS045 cleared');
  assert(SU.setupItems(p).find((i) => i.id === 'siteUrl').done, 'item ticked');
});

test('Finish setup: app link, email-link screen, admin GUID, usage confirmation', () => {
  const p = withFlows(); p.app.appLink = ''; delete p.app.deepLink;
  const ids = SU.openSetupItems(p).map((i) => i.id);
  assert(ids.includes('appLink') && ids.includes('deepLink'), 'link items open: ' + ids);
  SU.applySetup(p, 'appLink', 'https://apps.powerapps.com/play/e/x/a/y?tenantId=z');
  const det = p.screens.find((s) => s.name === 'scrDetailRequests');
  SU.applySetup(p, 'deepLink', { screen: det.id });
  eq(p.app.deepLink.list, 'Requests', 'list inferred from the screen');
  assert(!SU.openSetupItems(p).length, 'all done');
  assert(!codes(p).has('PS047'), 'PS047 cleared');
  const q = buildFt('requests');
  q.app.features.adminGroup = '';
  assert(SU.openSetupItems(q).some((i) => i.id === 'adminGroup'), 'admin-gated app without a group asks for it');
  eq(SU.applySetup(q, 'adminGroup', 'abc'), false, 'bad GUID refused');
  eq(SU.applySetup(q, 'adminGroup', '11111111-2222-3333-4444-555555555555'), true, 'GUID accepted');
  assert(!SU.openSetupItems(q).some((i) => i.id === 'adminGroup'), 'admin item done');
  assert(SU.openSetupItems(q).some((i) => i.id === 'usageList'), 'usage list needs confirming');
  SU.applySetup(q, 'usageList', true);
  assert(!SU.openSetupItems(q).length, 'nothing left');
  const r = core.normalizeProject(JSON.parse(JSON.stringify(q)));
  assert(r.app.setup.usageConfirmed === true, 'confirmation survives load');
});

// ---------------------------------------------------------------- round 2: customer guide
test('guide test cases cover every flow, decision, save screen and delete', () => {
  const p = buildFt('requests');
  p.flows = ['created', 'statusChanged', 'reminder', 'button'].map((k) => FL.newFlow(p, k, 'Requests'));
  p.app.siteUrl = 'https://contoso.sharepoint.com/sites/T';
  const cases = GD.guideCases(p), area = (a) => cases.filter((c) => c.area === a).length;
  eq(area('Flow'), p.flows.length, 'one case per enabled flow');
  const acts = GD.appActions(p);
  const decisions = new Set(acts.filter((x) => x.a.type === 'setField').map((x) => `${x.a.list}|${x.a.field}|${x.a.value}`));
  eq(area('Decision'), decisions.size, 'one case per distinct decision');
  const saveScreens = new Set(acts.filter((x) => x.a.type === 'save').map((x) => x.screen.id + x.a.list));
  eq(area('Create'), saveScreens.size, 'one create case per save screen');
  eq(area('Delete'), new Set(acts.filter((x) => x.a.type === 'delete').map((x) => x.a.list)).size, 'one delete case per list');
  assert(area('Start') === 1 && area('Access') === 1, 'loading and admin cases');
  cases.forEach((c, i) => eq(c.id, 'TC' + String(i + 1).padStart(2, '0'), 'numbered'));
});

test('guide HTML is self-contained', () => {
  const p = buildDemo(core.newProject);
  const html = GD.guideHtml(p, { thumbs: { [p.screens[0].id]: '<div class="pa-ctl"><img src="https://evil.example/x.png"></div>' }, css: '.pa-ctl{position:absolute}' });
  assert(html.startsWith('<!DOCTYPE html>'), 'a full page');
  assert(!/(src|href)\s*=\s*["'](https?:)?\/\//i.test(html), 'no external src or href');
  assert(!/url\(\s*["']?(https?:)?\/\//i.test(html), 'no external CSS urls');
  assert(!/<link\b/i.test(html), 'no stylesheet links');
  for (const f of p.flows) assert(html.includes(f.name.replace(/&/g, '&amp;')), 'flow listed: ' + f.name);
  assert(html.includes('<svg') && html.includes('Lifecycle'), 'lifecycle diagram');
  eq((html.match(/class="card tc"/g) || []).length, GD.guideCases(p).length, 'every case rendered');
});

// ---------------------------------------------------------------- round 2: demo
test('the demo is lint-clean, fully set up and paste-safe', () => {
  const p = buildDemo(core.newProject);
  eq(E.lint(p).filter((x) => x.level !== 'info').length, 0, 'no lint errors or warnings: ' + E.lint(p).filter((x) => x.level !== 'info').map((x) => x.code + ' ' + x.msg).join('; '));
  eq(SU.openSetupItems(p).length, 0, 'setup complete');
  eq(p.flows.length, 2, 'two flows');
  assert(p.screens[0].name === 'scrLoading', 'starts on the loading screen');
  const seen = new Set();
  for (const s of p.screens) checkScreenYaml(E.screenYaml(p, s), seen);
});


// ---------------------------------------------------------------- round 3: ask for a change
/** Structured-output limits: closed objects, every property required, no size/one-of keywords. */
function schemaLimits(node, path, out) {
  if (!node || typeof node !== 'object') return out;
  for (const bad of ['minItems', 'maxItems', 'minimum', 'maximum', 'minLength', 'maxLength', 'oneOf', 'anyOf', 'allOf', 'pattern', 'format']) if (bad in node) out.push(`${path}: ${bad}`);
  if (node.type === 'object') {
    if (node.additionalProperties !== false) out.push(`${path}: additionalProperties must be false`);
    const keys = Object.keys(node.properties || {});
    if (JSON.stringify([...(node.required || [])].sort()) !== JSON.stringify([...keys].sort())) out.push(`${path}: every property must be required`);
    for (const k of keys) schemaLimits(node.properties[k], `${path}.${k}`, out);
  }
  if (node.type === 'array') schemaLimits(node.items, `${path}[]`, out);
  return out;
}
test('change schema stays inside structured-output limits', () => {
  const bad = schemaLimits(CH.CHANGE_SCHEMA, 'change', []);
  assert(!bad.length, bad.join('; '));
  const opKeys = Object.keys(CH.CHANGE_SCHEMA.properties.ops.items.properties);
  for (const k of Object.keys(CHANGE_FIXTURE.ops[0])) assert(opKeys.includes(k), 'fixture field in schema: ' + k);
  for (const o of ['addList', 'addColumn', 'renameColumn', 'removeColumn', 'setChoices', 'addFlow', 'editFlow', 'removeFlow', 'setOption', 'addScreenFromTemplate', 'addFieldToForm', 'addFieldToDetail', 'addFieldToGallery', 'setTheme', 'renameApp'])
    assert(CH.CHANGE_SCHEMA.properties.ops.items.properties.op.enum.includes(o), 'op in enum: ' + o);
});
test('change request: summary has names and types only, never sample rows or setup values', () => {
  const p = buildDemo(core.newProject);
  p.lists[0].rows = [{ Title: 'secret row' }];
  const req = CH.buildChangeRequest(p, 'Add a Cost Center', 'claude-sonnet-5-5');
  eq(req.output_config.format.type, 'json_schema', 'structured output');
  assert(!('temperature' in req), 'no temperature');
  const sent = JSON.stringify(req.messages);
  assert(!sent.includes('secret row'), 'no row data');
  assert(!sent.includes(p.app.siteUrl), 'no site URL');
  assert(sent.includes('Requests') && sent.includes('EstimatedCost') && sent.includes('scrEditRequests'), 'lists, columns and screens');
});
test('change ops from a mocked Claude response: review, apply, skip, untouched screens byte-identical', async () => {
  const p = buildDemo(core.newProject);
  const before = Object.fromEntries(p.screens.map((s) => [s.id, E.screenYaml(p, s)]));
  const snapshot = JSON.stringify(p);
  let sent = null;
  const fetchFn = async (url, init) => { sent = JSON.parse(init.body); return { ok: true, status: 200, json: async () => ({ stop_reason: 'end_turn', content: [{ type: 'text', text: JSON.stringify(CHANGE_FIXTURE) }], usage: {} }) }; };
  const r = await AI.callClaude({ key: 'sk-ant-test', request: CH.buildChangeRequest(p, 'Add a cost center', 'claude-sonnet-5-5'), fetchFn });
  eq(sent.system, CH.CHANGE_PROMPT, 'change prompt sent');
  const rows = CH.reviewOps(p, r.plan.ops);
  eq(rows.filter((x) => x.error).length, 2, 'two invalid ops flagged: ' + rows.map((x) => x.error).join('|'));
  assert(rows.find((x) => x.op.op === 'removeColumn').destructive && rows.find((x) => x.op.op === 'setChoices').destructive, 'destructive flagged');
  const q = core.clone(p);
  const res = CH.applyOps(q, r.plan.ops);
  eq(JSON.stringify(p), snapshot, 'the original project is not touched (undo restores it exactly)');
  eq(res.applied.length, 6, 'six applied');
  eq(res.skipped.length, 2, 'two skipped with reasons');
  assert(res.skipped.every((x) => x.reason), 'reasons');
  const l = q.lists.find((x) => x.name === 'Requests');
  assert(l.columns.some((c) => c.name === 'CostCenter' && c.required), 'column added with a clean name');
  assert(l.columns.find((c) => c.name === 'Category').choices.includes('Finance'), 'choices updated');
  assert(q.flows.some((f) => f.name === 'Weekly pending reminder' && f.kind === 'reminder' && f.recipients[0].value === 'Approver'), 'flow added');
  const edit = q.screens.find((s) => s.name === 'scrEditRequests');
  assert([...core.walk(edit.controls)].some(([c]) => c.field?.column === 'CostCenter'), 'on the form');
  assert(res.touched.has(edit.id), 'form touched');
  assert(!edit.controls.find((c) => c.name === 'lblNeededByEditRequests').fx?.Text, 'the NeededBy caption is not mistaken for the Still needed strip');
  assert(/Cost Center/.test(edit.controls.find((c) => c.name === 'lblNeededEditRequests').fx.Text), 'Still needed strip knows the new required field');
  let changed = 0;
  for (const s of q.screens) {
    const y = E.screenYaml(q, s);
    if (res.touched.has(s.id)) { if (y !== before[s.id]) changed++; continue; }
    eq(y, before[s.id], `untouched ${s.name} is byte-identical`);
  }
  assert(changed >= 3, 'form, details and browse changed');
  const errs = errors(q); assert(!errs.length, errs.map((x) => x.code + ' ' + x.msg).join('; '));
  assert(!E.lint(q).some((f) => f.level === 'warn'), 'no new warnings: ' + E.lint(q).filter((f) => f.level === 'warn').map((f) => f.code + ' ' + f.where + ' ' + f.msg).join('; '));
  checkApp(q);
  const fm = E.lint(q); void fm;
  // the form still fits above the Save bar
  const bar = edit.controls.find((c) => c.name.startsWith('rectBar'));
  assert(Math.max(...edit.controls.filter((c) => c.y < bar.y && !/^(rectDone|cardDone|icoDone|lblDone|btnDone)/.test(c.name)).map((c) => c.y + c.h)) <= bar.y, 'form fits above the bar');
});
test('change ops: lists, screens, options and flows in one batch', () => {
  const p = buildDemo(core.newProject), q = core.clone(p);
  const B = CHANGE_FIXTURE.ops[0], op = (o) => ({ ...B, reason: '', required: false, list: '', column: '', type: '', ...o });
  const res = CH.applyOps(q, [
    op({ op: 'addList', list: 'Vendors', columns: [{ name: 'Contact', type: 'Person', required: false, choices: [], lookupList: '' }, { name: 'Tier', type: 'Choice', required: true, choices: ['Gold', 'Silver'], lookupList: '' }] }),
    op({ op: 'addScreenFromTemplate', template: 'edit', list: 'Vendors' }),
    op({ op: 'addScreenFromTemplate', template: 'detail', list: 'Vendors' }),
    op({ op: 'addScreenFromTemplate', template: 'browse', list: 'Vendors' }),
    op({ op: 'renameColumn', list: 'Requests', column: 'Priority', newName: 'Urgency' }),
    op({ op: 'setOption', option: 'adminGroupId', value: '' }),
    op({ op: 'setOption', option: 'loadingScreen', value: 'false' }),
    op({ op: 'editFlow', flowName: p.flows[0].name, subject: 'Heads up: {Title}' }),
    op({ op: 'removeFlow', flowName: p.flows[1].name }),
    op({ op: 'renameApp', newName: 'Ops Desk' }),
  ]);
  eq(res.skipped.length, 0, 'none skipped: ' + res.skipped.map((x) => x.reason).join('; '));
  assert(q.screens.some((s) => s.name === 'scrBrowseVendors') && !q.screens.some((s) => s.role === 'loading'), 'screens added/removed');
  assert(q.app.startScreen === q.screens[0].id, 'start screen valid');
  const hay = JSON.stringify(q.screens);
  assert(!/\bPriority\b/.test(hay), 'no stale Priority references');
  assert(!/varIsAdmin/.test(hay), 'admin gating removed');
  eq(q.flows.length, 1, 'flow removed');
  assert(q.flows[0].subject.startsWith('Heads up'), 'flow edited');
  const errs = errors(q); assert(!errs.length, errs.map((x) => x.code + ' ' + x.msg).join('; '));
  checkApp(q);
});
test('field adds on phone keep overlays in front and the form above the bar', () => {
  const p = build('requests', 'phone'), q = core.clone(p);
  const B = CHANGE_FIXTURE.ops[0];
  const res = CH.applyOps(q, [{ ...B }, { ...CHANGE_FIXTURE.ops[1] }, { ...CHANGE_FIXTURE.ops[2] }]);
  eq(res.skipped.length, 0, 'applied: ' + res.skipped.map((x) => x.reason).join('; '));
  const edit = q.screens.find((s) => s.name.startsWith('scrEdit'));
  const iNew = edit.controls.findIndex((c) => c.field?.column === 'CostCenter');
  const iBar = edit.controls.findIndex((c) => /^rectBar/.test(c.name));
  assert(iNew >= 0 && iNew < iBar, 'new input sits behind the bottom bar and overlays');
  checkApp(q);
});

// ---------------------------------------------------------------- round 3: CSV / Excel import
test('CSV parser: quotes, doubled quotes, commas and newlines in quotes, CRLF, BOM, blank lines, other delimiters', () => {
  const rows = TB.parseCSV('﻿a,b,c\r\n"x, y","say ""hi""","l1\nl2"\r\n\r\n1,,3\n');
  eq(JSON.stringify(rows), JSON.stringify([['a', 'b', 'c'], ['x, y', 'say "hi"', 'l1\nl2'], ['1', '', '3']]), 'rows');
  eq(JSON.stringify(TB.parseCSV('a;b\n1;"2;3"')), JSON.stringify([['a', 'b'], ['1', '2;3']]), 'semicolon');
  eq(JSON.stringify(TB.parseCSV('a\tb\n1\t2')), JSON.stringify([['a', 'b'], ['1', '2']]), 'tab');
  eq(JSON.stringify(TB.parseCSV('a,b\n1,2')), JSON.stringify([['a', 'b'], ['1', '2']]), 'no trailing newline');
  const back = TB.parseCSV(TB.writeCSV([['x, y', 'say "hi"', 'l1\nl2', ' pad']]));
  eq(JSON.stringify(back), JSON.stringify([['x, y', 'say "hi"', 'l1\nl2', ' pad']]), 'round trip');
});
test('type inference: Number, Currency, DateTime, YesNo, Choice, Person, Note, Text', () => {
  const head = ['Name', 'Owner', 'Qty', 'Cost', 'Due', 'Done', 'Dept', 'Notes', 'Zip', 'Phone'];
  const rows = [
    ['Laptop', 'ana@contoso.com', '3', '$1,200.00', '2026-01-03', 'Yes', 'IT', 'short', '02134', '555-0101'],
    ['Desk', 'ben@contoso.com', '12', '$85.50', '1/5/2026', 'no', 'HR', 'a\nmulti-line note', '10001', '555-0102'],
    ['Chair', 'cy@contoso.com', '1,000', '$40', '2026-02-01', 'TRUE', 'IT', 'ok', '99999', '555-0103'],
    ['Phone', 'di@contoso.com', '-4', '$7', '2026-03-01', 'false', 'IT', '', '12345', '555-0104'],
  ];
  const t = Object.fromEntries(TB.inferColumns(head, rows).map((c) => [c.header, c.type]));
  eq(JSON.stringify(t), JSON.stringify({ Name: 'Text', Owner: 'Person', Qty: 'Number', Cost: 'Currency', Due: 'DateTime', Done: 'YesNo', Dept: 'Choice', Notes: 'Note', Zip: 'Text', Phone: 'Text' }), 'types');
  eq(TB.inferColumns(['Budget'], [['1200.50'], ['80']])[0].type, 'Currency', 'money header with plain numbers');
  eq(TB.inferColumns(['Mixed'], [['ana@contoso.com'], ['not an email']])[0].type, 'Text', 'Person only when every value is an email');
  eq(TB.inferColumns(['Unique'], [['a'], ['b'], ['c'], ['d']])[0].type, 'Text', 'Choice needs repeats');
  eq(TB.convertCell('Currency', '($1,200.50)'), -1200.5, 'accounting negative');
  eq(TB.convertCell('DateTime', '2/30/2026'), '', 'invalid date');
  eq(TB.convertCell('DateTime', '12/31/25'), '2025-12-31', 'two-digit year');
});
/** A minimal real .xlsx: shared strings, an inline string, a styled date, a boolean, a skipped cell. */
async function makeXlsx() {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  z.file('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Assets" sheetId="1" r:id="rId7"/><sheet name="Other" sheetId="2" r:id="rId8"/></sheets></workbook>');
  z.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId8" Target="worksheets/sheet1.xml"/><Relationship Id="rId7" Target="worksheets/sheet2.xml"/></Relationships>');
  z.file('xl/sharedStrings.xml', '<sst><si><t>Asset</t></si><si><t>Purchased</t></si><si><r><t>Ow</t></r><r><t>ner</t></r><rPh><t>x</t></rPh></si><si><t>Active</t></si><si><t xml:space="preserve">Laptop &amp; dock</t></si><si><t>ana@contoso.com</t></si><si><t>Cost</t></si></sst>');
  z.file('xl/styles.xml', '<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="14"/></cellXfs></styleSheet>');
  z.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>WRONG SHEET</t></is></c></row></sheetData></worksheet>');
  z.file('xl/worksheets/sheet2.xml', '<worksheet><sheetData>'
    + '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="s"><v>6</v></c></row>'
    + '<row r="2"><c r="A2" t="s"><v>4</v></c><c r="B2" s="1"><v>46025</v></c><c r="C2" t="s"><v>5</v></c><c r="D2" t="b"><v>1</v></c><c r="E2"><v>1299.5</v></c></row>'
    + '<row r="3"><c r="A3" t="inlineStr"><is><t>Monitor</t></is></c><c r="B3" s="2"><v>46054</v></c><c r="D3" t="b"><v>0</v></c><c r="E3"><v>199</v></c></row>'
    + '</sheetData></worksheet>');
  return z.generateAsync({ type: 'uint8array' });
}
test('xlsx reader: first sheet in workbook order, shared/rich/inline strings, dates from styles, booleans, gaps', async () => {
  const rows = await TB.readXlsx(await makeXlsx(), JSZip);
  eq(JSON.stringify(rows[0]), JSON.stringify(['Asset', 'Purchased', 'Owner', 'Active', 'Cost']), 'header (rich text joined, phonetic dropped)');
  eq(JSON.stringify(rows[1]), JSON.stringify(['Laptop & dock', '2026-01-03', 'ana@contoso.com', 'TRUE', '1299.5']), 'row 2');
  eq(JSON.stringify(rows[2]), JSON.stringify(['Monitor', '2026-02-01', '', 'FALSE', '199']), 'row 3 keeps the gap');
  const types = TB.inferColumns(rows[0], rows.slice(1)).map((c) => c.type);
  eq(JSON.stringify(types), JSON.stringify(['Text', 'DateTime', 'Person', 'YesNo', 'Currency']), 'types');
});
test('imported rows become a list with clean names, are capped at 200, and Play uses them', () => {
  const p = core.newProject();
  const head = ['Item Name', 'Created', 'Status', 'Owner', 'Amount'];
  const data = Array.from({ length: 230 }, (_, i) => [`Thing ${i + 1}`, '2026-01-0' + ((i % 9) + 1), ['Open', 'Closed'][i % 2], i % 2 ? 'ana@contoso.com' : 'zed.q@fabrikam.com', String(i * 10)]);
  const info = TB.inferColumns(head, data).map((c, i) => ({ ...c, index: i, include: true, name: c.header }));
  const { list, notes, truncated } = TB.listFromTable(p, { name: 'my assets!', rows: data, cols: info, titleIndex: 0 });
  eq(list.name, 'MyAssets', 'list name cleaned');
  assert(truncated && notes.some((n) => /200/.test(n)), 'capped with a note');
  eq(list.rows.length, 200, '200 rows kept');
  const names = list.columns.map((c) => c.name);
  assert(!names.includes('Created') && names.some((n) => /Created$/.test(n)), 'reserved name prefixed: ' + names.join(','));
  eq(list.columns.find((c) => c.name === 'Status').type, 'Choice', 'choice');
  eq(list.rows[0].Title, 'Thing 1', 'Title from the chosen column');
  p.lists.push(list);
  generateApp(p, { screens: ['browse', 'detail', 'edit'] });
  assert(p.lists[0].rows.length === 200, 'generateApp keeps the rows');
  const saved = JSON.parse(JSON.stringify(p)); core.normalizeProject(saved);
  eq(saved.lists[0].rows.length, 200, 'rows survive save/load');
  saved.lists[0].rows.unshift({ Title: 'x', Evil: { a: 1 }, Nope: 'y' }); core.normalizeProject(saved);
  eq(saved.lists[0].rows.length, 200, 'normalize caps rows');
  eq(JSON.stringify(saved.lists[0].rows[0]), JSON.stringify({ Title: 'x' }), 'unknown keys and objects dropped');
  const rt = createRuntime(p, { toast: () => {}, rerender: () => {} });
  eq(rt.db.MyAssets.length, 200, 'Play has every imported row');
  eq(rt.db.MyAssets[0].Title, 'Thing 1', 'Play row 1');
  eq(rt.db.MyAssets[0].Owner.Mail, 'zed.q@fabrikam.com', 'unknown person keeps the address');
  eq(rt.db.MyAssets[0].Owner.DisplayName, 'Zed Q', 'name from the address');
  eq(typeof rt.db.MyAssets[3].Amount, 'number', 'numbers stay numbers');
  const back = TB.parseCSV(TB.listToCsv(p.lists[0]));
  eq(back.length, 201, 'header + rows');
  eq(back[0][0], 'Title', 'internal names as headers');
  eq(back[1][0], 'Thing 1', 'first row');
  checkApp(p);
});

// ---------------------------------------------------------------- round 3: accessibility
test('contrast math matches WCAG reference values and the fix reaches AA with the smallest nudge', () => {
  const near = (a, b, m) => assert(Math.abs(a - b) < 0.01, `${m}: ${a} vs ${b}`);
  near(A11.contrast('#000000', '#FFFFFF'), 21, 'black on white');
  near(A11.contrast('#FFFFFF', '#FFFFFF'), 1, 'same color');
  near(A11.contrast('#767676', '#FFFFFF'), 4.54, '#767676 is the lightest AA gray on white');
  near(A11.contrast('#777777', '#FFFFFF'), 4.48, '#777777 just fails');
  near(A11.contrast('#0F6CBD', '#FFFFFF'), A11.contrast('#FFFFFF', '#0F6CBD'), 'symmetric');
  const colors = { ...core.THEME_PRESETS[0].colors, TextMuted: '#A0AEC0', OnPrimary: '#FFFFFF', Primary: '#F59E0B' };
  const issues = A11.contrastIssues(colors).map((x) => x.fg + '/' + x.bg);
  assert(issues.includes('TextMuted/Surface') && issues.includes('OnPrimary/Primary'), 'issues found: ' + issues);
  const f1 = A11.fixContrast(colors, 'TextMuted', 'Surface');
  eq(f1.key, 'TextMuted', 'text color nudged first');
  assert(A11.contrast(f1.value, colors.Surface) >= 4.5 && A11.contrast(f1.value, colors.Surface) < 5.2, 'just past AA: ' + f1.value);
  const f2 = A11.fixContrast(colors, 'OnPrimary', 'Primary');
  eq(f2.key, 'Primary', 'white text cannot get lighter, so the button color darkens');
  assert(A11.contrast(colors.OnPrimary, f2.value) >= 4.5, 'button passes: ' + f2.value);
  const p = buildDemo(core.newProject);
  p.theme.colors.TextMuted = '#A0AEC0';
  const f = E.lint(p).filter((x) => x.code === 'PS052' && /TextMuted/.test(x.msg));
  eq(f.length, 1, 'one PS052 finding for TextMuted');
  eq(f[0].level, 'info', 'info only');
  p.theme.colors.TextMuted = A11.fixContrast(p.theme.colors, 'TextMuted', 'Surface').value;
  eq(E.lint(p).filter((x) => x.code === 'PS052' && /TextMuted/.test(x.msg)).length, 0, 'fixed');
});
test('accessibility labels are opt-in: default YAML has none, opting in labels inputs and tap buttons', () => {
  const p = buildDemo(core.newProject);
  const before = p.screens.map((s) => E.screenYaml(p, s)).join('\n');
  assert(!before.includes('AccessibleLabel'), 'off by default');
  const saved = JSON.parse(JSON.stringify(p)); core.normalizeProject(saved);
  assert(!('a11y' in saved.app.features), 'flag not added to existing projects');
  p.app.features.a11y = true;
  const seen = new Set();
  const edit = p.screens.find((s) => s.name === 'scrEditRequests');
  const y = E.screenYaml(p, edit);
  checkScreenYaml(y, seen);
  assert(/txtTitleEditRequests:[\s\S]*?AccessibleLabel: ="Title"/.test(y), 'input labeled with its column');
  assert(/ppApproverEditRequests:[\s\S]*?AccessibleLabel: ="Approver"/.test(y), 'people picker labeled');
  assert(/icoBackEditRequestsTap:[\s\S]*?AccessibleLabel: ="Back"/.test(y), 'back icon tap button labeled');
  for (const s of p.screens) if (s !== edit) checkScreenYaml(E.screenYaml(p, s), seen);
  const saved2 = JSON.parse(JSON.stringify(p)); core.normalizeProject(saved2);
  eq(saved2.app.features.a11y, true, 'flag survives save/load');
});

// ---------------------------------------------------------------- round 3: drift check against a Studio export
test('drift check: reads a solution zip with a nested .msapp, finds Studio-only controls and screens, imports them verbatim', async () => {
  const DR = await import('../js/drift.js');
  const p = buildDemo(core.newProject);
  const z = new JSZip();
  for (const s of p.screens) {
    let y = E.screenYaml(p, s);
    if (s.name === 'scrDetailRequests') y += '\n    - btnStudioPrint:\n        Control: Classic/Button\n        Properties:\n          Text: ="Print"\n          X: =Parent.Width - 200\n          Y: =100\n          OnSelect: =Print()\n    - grpStudio:\n        Control: GroupContainer\n        Variant: ManualLayout\n        Children:\n          - lblInGroup:\n              Control: Label\n              Properties:\n                Text: ="hi"\n';
    if (s.name !== 'scrApprovalsRequests') z.file(`Src/${s.name}.pa.yaml`, y);
  }
  z.file('Src/App.pa.yaml', 'App:\n  Properties:\n    OnStart: =1\n');
  z.file('Src/scrStudioOnly.pa.yaml', 'Screens:\n  scrStudioOnly:\n    Children:\n      - lblHelloStudio:\n          Control: Label\n          Properties:\n            Text: ="Hello"\n');
  const sol = new JSZip(); sol.file('CanvasApps/contoso_app_DocumentUri.msapp', await z.generateAsync({ type: 'uint8array' }));
  const st = await DR.readStudioExport([{ name: 'Solution.zip', data: await sol.generateAsync({ type: 'uint8array' }) }], JSZip);
  eq(st.notes.length, 0, 'no notes');
  const rep = DR.driftReport(p, st);
  const det = rep.screens.find((s) => s.name === 'scrDetailRequests');
  eq(JSON.stringify(det.studioOnly.map((c) => c.name)), JSON.stringify(['btnStudioPrint', 'grpStudio']), 'Studio-only top-level controls (children ride along)');
  eq(det.studioOnly[1].nested, 1, 'nested count');
  eq(rep.newScreens[0].name, 'scrStudioOnly', 'Studio-only screen');
  assert(!rep.screens.some((s) => s.name === 'scrApprovalsRequests'), 'screens missing from the export are not reported as drift');
  const before = Object.fromEntries(p.screens.map((s) => [s.id, E.screenYaml(p, s)]));
  const res = DR.importDrift(p, st, [{ screen: 'scrDetailRequests', names: ['btnStudioPrint', 'grpStudio'] }], ['scrStudioOnly']);
  eq(res.added, 3, 'three imported');
  for (const s of p.screens) if (before[s.id] && s.name !== 'scrDetailRequests') eq(E.screenYaml(p, s), before[s.id], `${s.name} untouched`);
  const y = E.screenYaml(p, p.screens.find((s) => s.name === 'scrDetailRequests'));
  assert(/btnStudioPrint:\n\s+Control: Classic\/Button[\s\S]*?OnSelect: =Print\(\)[\s\S]*?X: =Parent\.Width - 200/.test(y), 'kept verbatim, formula position kept');
  assert(/grpStudio:\n\s+Control: GroupContainer\n\s+Variant: ManualLayout[\s\S]*?lblInGroup:/.test(y), 'container with children');
  const after = DR.driftReport(p, st);
  eq(after.screens.filter((s) => s.studioOnly.length).length + after.newScreens.length, 0, 'no drift left');
  const none = await DR.readStudioExport([{ name: 'old.msapp', data: await new JSZip().file('Header.json', '{}').generateAsync({ type: 'uint8array' }) }], JSZip);
  assert(/no YAML source/.test(none.notes[0] || ''), 'explains an .msapp without Src');
});


// ---------------------------------------------------------------- round 4: contrast, approval chains
const CHN = await import('../js/chain.js');
const R4 = await import('../js/render.js');
const tryFxBool = (p, rt, src) => !!R4.tryFx(src, { project: p, rt }, false);
const { createHash } = await import('node:crypto');
test('every theme preset passes WCAG AA for its text pairs', () => {
  for (const t of core.THEME_PRESETS) { const bad = A11.contrastIssues(t.colors); assert(!bad.length, `${t.name}: ${JSON.stringify(bad)}`); }
  eq(core.THEME_PRESETS.find((t) => t.name === 'SaTech Cyan').colors.Primary, '#0E7490', 'SaTech Cyan primary');
  eq(core.THEME_PRESETS.find((t) => t.name === 'Sunset').colors.Primary, '#C4451C', 'Sunset primary');
});
const STAGES = [
  { name: 'Manager', approver: 'column' },
  { name: 'Finance', approver: 'group', groupId: '11111111-2222-3333-4444-555555555555', groupLabel: 'Finance team', notify: 'finance@contoso.com' },
  { name: 'Director', approver: 'column' },
  { name: 'Legal', approver: 'column' },
];
function buildChain(n, ff = 'tablet', stages = STAGES) {
  const p = core.newProject(); p.app.name = 'Requests'; p.app.formFactor = ff; p.lists = SCENARIOS.requests.lists();
  p.app.siteUrl = 'https://contoso.sharepoint.us/sites/ops'; p.app.appLink = 'https://apps.gov.powerapps.us/play/e/x/a/y';
  generateApp(p, { screens: ['home', 'browse', 'detail', 'edit'], approvals: SCENARIOS.requests.approvals, chain: { list: 'Requests', stages: stages.slice(0, n) } });
  AI.addChainFlows(p);
  return p;
}
const ctlIn = (p, scrName, re) => [...core.walk(p.screens.find((s) => s.name === scrName).controls)].map(([c]) => c).find((c) => re.test(c.name));
for (const n of [1, 2, 3, 4]) for (const ff of ['tablet', 'phone']) {
  test(`approval chain: ${n} stage${n > 1 ? 's' : ''} (${ff}) is lint-clean and paste-safe`, () => {
    const p = buildChain(n, ff);
    const bad = E.lint(p).filter((f) => f.level !== 'info');
    assert(!bad.length, bad.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
    const yamls = checkApp(p);
    assert(yamls.scrApprovalsRequests, 'approvals screen');
    const l = p.lists[0];
    for (const st of STAGES.slice(0, n)) {
      const k = CHN.stageCols(st);
      for (const c of [k.decision, k.date, k.by, k.comments]) assert(l.columns.some((x) => x.name === c), 'column ' + c);
      eq(l.columns.find((x) => x.name === k.decision).choices.join('/'), 'Pending/Approved/Rejected', 'decision choices');
    }
    assert(l.columns.some((c) => c.name === 'CurrentStage' && c.type === 'Number'), 'CurrentStage');
    assert(l.columns.some((c) => c.name === 'ApprovalHistory' && c.type === 'Note'), 'ApprovalHistory');
    assert(E.provisioningScript(p).includes('CurrentStage') && /<Default>1<\/Default>/.test(E.provisioningScript(p)), 'provisioning adds CurrentStage with default 1');
    // stepper: one coloured step per stage, each read from its own decision column
    const html = ctlIn(p, 'scrDetailRequests', /^htmlChain/);
    for (const st of STAGES.slice(0, n)) assert(html.fx.HtmlText.includes(`varSelRequests.${CHN.stageCols(st).K}Decision.Value`), 'stepper reads ' + st.name);
    eq((html.fx.HtmlText.match(/appTheme\.Success/g) || []).length >= n, true, 'a green state per stage');
    // home shows what is waiting on me; save sets CurrentStage on new items
    const editY = E.screenYaml(p, p.screens.find((s) => s.name === 'scrEditRequests'));
    assert(/CurrentStage: Coalesce\(LookUp\(Requests, ID = varSelRequests\.ID\)\.CurrentStage, 1\)/.test(editY), 'save starts new items at stage 1');
    eq(p.flows.length, n + 2, 'flows: stage 1 new + resubmit, one per hand-off, outcome');
  });
}
test('approval chain: decision Patch is rebased, guarded and appends history', () => {
  const p = buildChain(2);
  const det = p.screens.find((s) => s.name === 'scrDetailRequests');
  const ap = ctlIn(p, 'scrDetailRequests', /^btnChainApprove/), rj = ctlIn(p, 'scrDetailRequests', /^btnChainReject/);
  const fx = E.actionFx(p, det, ap.actions.OnSelect[0]);
  assert(/^With\(\{cur: LookUp\(Requests, ID = varSelRequests\.ID\)\}/.test(fx), 'rebased on a fresh LookUp');
  assert(/Switch\(cur\.CurrentStage,\s*1,/.test(fx) && /\n\s*2,/.test(fx), 'switch per stage');
  assert(/ManagerDecision: \{Value: "Approved"\}/.test(fx) && /ManagerDecidedOn: Today\(\)/.test(fx), 'decision + date');
  assert(/ManagerDecidedBy: \{\s*'@odata\.type'/.test(fx) || /ManagerDecidedBy: \{/.test(fx), 'decided-by person record');
  assert(/CurrentStage: 2/.test(fx) && /CurrentStage: 0/.test(fx), 'advance, then finish');
  assert(/Status: \{Value: "Approved"\}/.test(fx) && /Status: \{Value: "Submitted"\}/.test(E.actionFx(p, det, { type: 'resubmit', list: 'Requests' })), 'overall status');
  assert(/ApprovalHistory: .*If\(IsBlank\(cur\.ApprovalHistory\), "", Char\(10\) & cur\.ApprovalHistory\)/.test(fx), 'history line prepended, never replaced');
  const ifErr = fx.match(/IfError\(/g).length; eq(ifErr, 2, 'one IfError per stage');
  assert(/Notify\("Approval failed\. " & FirstError\.Message, NotificationType\.Error\)/.test(fx) || /FirstError\.Message/.test(fx), 'failure branch');
  const rfx = E.actionFx(p, det, rj.actions.OnSelect[0]);
  assert(/\{Value: "Rejected"\}/.test(rfx) && /Status: \{Value: "Rejected"\}/.test(rfx), 'reject sets both');
  assert(/IsBlank\(Trim\(txtDecisionComment\w*\.Text\)\)/.test(rj.fx.DisplayMode || ''), 'reject disabled until there is a comment');
  const can = ap.fx.Visible;
  assert(/Lower\(varSelRequests\.'Created By'\.Email\) <> Lower\(User\(\)\.Email\)/.test(can), 'never your own request');
  assert(/varSelRequests\.CurrentStage = 2 && varIsApproverFinance/.test(can), 'group stage uses the group check');
  const on = E.onStartFx(p);
  assert(/Set\(varIsApproverFinance, IfError\(!IsBlank\(LookUp\(Office365Groups\.ListGroupMembers\("11111111-2222-3333-4444-555555555555"\)\.value, Lower\(mail\) = Lower\(User\(\)\.Email\)\)\), false\)\)/.test(on), 'group check wrapped in IfError');
  assert(E.checklist(p).some((s) => /Office 365 Groups/.test(JSON.stringify(s))), 'checklist adds Office 365 Groups');
  const q = ctlIn(p, 'scrApprovalsRequests', /^gal/);
  assert(/Filter\(\s*Filter\(Requests, /.test(q.fx.Items) && /'Created By'\.Email <> User\(\)\.Email/.test(q.fx.Items), 'queue: delegable inner filter, own items removed in the outer one');
});
test('approval chain: Play walks approve → next stage → final, reject needs a comment, history and resubmit', () => {
  const p = buildChain(2);
  const toasts = [];
  const rt = createRuntime(p, { toast: (t, k) => toasts.push(`${k}: ${t}`), rerender() {} });
  const ch = p.app.chain, rows = rt.db.Requests;
  const queue = rows.filter((r) => CHN.playCanDecide(ch, r, rt.vars));
  assert(queue.length >= 2, 'sample data leaves work in the queue');
  assert(!queue.some((r) => r.Author.DisplayName === 'Alex Rivera'), 'own items never queued');
  const det = p.screens.find((s) => s.name === 'scrDetailRequests');
  const ap = ctlIn(p, 'scrDetailRequests', /^btnChainApprove/), rj = ctlIn(p, 'scrDetailRequests', /^btnChainReject/), rs = ctlIn(p, 'scrDetailRequests', /^btnResubmit/);
  const row = rows.find((r) => r.CurrentStage === 1 && r.Author.DisplayName !== 'Alex Rivera');
  rt.vars.varSelRequests = row; rt.screen = det.id;
  rt.run(det, ap.actions.OnSelect);
  eq(row.CurrentStage, 2, 'moved to stage 2'); eq(row.ManagerDecision, 'Approved', 'stage 1 approved'); eq(row.Status, 'In review', 'in review');
  assert(toasts.some((t) => /Notify Finance after Manager/.test(t)), 'hand-off email simulated');
  toasts.length = 0;
  rt.run(det, rj.actions.OnSelect);
  eq(row.CurrentStage, 2, 'reject without comment is refused'); assert(toasts.some((t) => /comment/.test(t)), 'asks for a comment');
  rt.inputs[rj.actions.OnSelect[0].commentCtl] = 'Over budget';
  rt.run(det, rj.actions.OnSelect);
  eq(row.Status, 'Rejected', 'rejected'); eq(row.FinanceDecision, 'Rejected', 'stage 2 rejected'); eq(row.CurrentStage, 0, 'chain ended');
  assert(toasts.some((t) => /Tell requester the outcome/.test(t)), 'outcome email simulated');
  const lines = row.ApprovalHistory.split('\n');
  assert(/Finance · Rejected · Over budget$/.test(lines[0].replace(/\s+/g, ' ')) && /Manager · Approved/.test(lines[1].replace(/\s+/g, ' ')), 'history newest first: ' + lines.join(' | '));
  // the requester resubmits their own rejected item
  const own = rows.find((r) => r.Author.DisplayName === 'Alex Rivera' && r.Status === 'Rejected');
  rt.vars.varSelRequests = own;
  eq(tryFxBool(p, rt, ap.fx.Visible), false, 'cannot approve own request');
  eq(tryFxBool(p, rt, rs.fx.Visible), true, 'can resubmit own rejected request');
  rt.run(det, rs.actions.OnSelect);
  eq(own.CurrentStage, 1, 'back to stage 1'); eq(own.Status, 'Submitted', 'pending again'); assert(/Resubmitted/.test(own.ApprovalHistory.split('\n')[0]), 'resubmit logged');
  assert(toasts.some((t) => /Notify Manager of resubmitted item/.test(t)), 'stage 1 re-notified');
  // the full chain: approve stage 1 and stage 2 of another item
  const r2 = rows.find((r) => r.CurrentStage === 1 && r !== row && r.Author.DisplayName !== 'Alex Rivera') || rows.find((r) => r.CurrentStage === 2);
  rt.vars.varSelRequests = r2; while (r2.CurrentStage) rt.run(det, ap.actions.OnSelect);
  eq(r2.Status, 'Approved', 'final approval');
});
test('approval chain: default output is byte-identical to round 3', () => {
  const out = [];
  const dump = (p) => { out.push(E.themeFormulas(p), E.onStartFx(p), E.provisioningScript(p), JSON.stringify(E.checklist(p))); for (const s of p.screens) out.push(E.screenYaml(p, s)); for (const f of p.flows || []) out.push(JSON.stringify(FL.flowDefinition(p, f))); };
  const navy = () => core.clone(core.THEME_PRESETS.find((t) => t.name === 'Federal Navy'));
  // the scenarios that existed in round 3 (later rounds add more; their output is covered by their own tests)
  for (const sc of ['requests', 'assets', 'events', 'tasks', 'contacts', 'blank'].map((k) => SCENARIOS[k])) for (const ff of ['tablet', 'phone']) {
    const p = core.newProject(); p.app.name = sc.app; p.app.formFactor = ff; p.lists = sc.lists(); p.theme = navy(); p.app.siteUrl = 'https://x.sharepoint.us/sites/a';
    generateApp(p, { screens: ['home', 'browse', 'detail', 'edit'], approvals: sc.approvals || null }); dump(p);
  }
  const r = AI.specToProject(FIXTURE); r.project.theme = navy(); dump(r.project);
  // sha256 of the same output on origin/main 308cbb9 (round 3)
  // round 6: detail section titles now hide with their cards (Visible: has)
  // tenant fix 1: HtmlViewer boxes get 2px slack + clip (grey scrollbars in a real tenant); flows guard blank Person recipients
  eq(createHash('sha256').update(out.join('\n')).digest('hex'), 'ddf433e886bc52c8055ecc22e8b8c789ced58d977f8dabd3c94f7b473655416d', 'output hash');
  assert(!r.project.app.chain, 'the round-3 plan fixture (no stages) stays a single approvals queue');
});
test('approval chain: lint catches bad group IDs, wrong approver columns and missing columns', () => {
  const p = buildChain(2, 'tablet', [STAGES[0], { ...STAGES[1], groupId: 'finance' }]);
  assert(codes(p).has('PS054'), 'PS054 bad group id');
  const q = buildChain(2);
  q.lists[0].columns = q.lists[0].columns.filter((c) => c.name !== 'FinanceDecidedOn');
  assert(codes(q).has('PS056'), 'PS056 missing column');
  q.app.chain.stages[0].column = 'Title';
  assert(codes(q).has('PS055'), 'PS055 approver not a Person column');
  const r = buildChain(1);
  delete r.app.chain;
  assert(codes(r).has('PS057'), 'PS057 decide without a chain');
});
test('approval chain: survives save/load, rename, and Claude plans with stages', () => {
  const p = buildChain(3);
  const back = core.normalizeProject(JSON.parse(JSON.stringify(p)));
  eq(JSON.stringify(back.app.chain), JSON.stringify(p.app.chain), 'chain kept');
  const junk = core.normalizeProject({ ...JSON.parse(JSON.stringify(p)), app: { ...p.app, chain: { list: '', stages: 'x' } } });
  assert(!junk.app.chain, 'unusable chain dropped');
  eq(ED.renameColumnIn(p, 'Requests', 'ManagerApprover', 'LineManager'), null, 'rename approver');
  eq(p.app.chain.stages[0].column, 'LineManager', 'chain follows the rename');
  assert(/LineManager\.Email/.test(ctlIn(p, 'scrApprovalsRequests', /^gal/).fx.Items), 'queue formula renamed');
  assert(/approval chain/.test(ED.renameColumnIn(p, 'Requests', 'ManagerDecision', 'X') || ''), 'chain columns are not renamed by hand');
  assert(ED.removeColumnIn(p, 'Requests', 'CurrentStage').error, 'chain columns are not removed by hand');
  ED.renameListIn(p, 'Requests', 'Purchases'); eq(p.app.chain.list, 'Purchases', 'list rename');
  const plan = JSON.parse(JSON.stringify(FIXTURE));
  const apl = plan.lists.find((l) => l.name === plan.screens.approvals.list) || plan.lists[0];
  plan.screens.approvals.enabled = true; plan.screens.approvals.list = apl.name;
  plan.screens.approvals.stages = [{ name: 'Supervisor', approverKind: 'personColumn', approverColumn: 'Supervisor', groupId: '', notifyEmail: '' }, { name: 'Finance', approverKind: 'group', approverColumn: '', groupId: 'not-a-guid', notifyEmail: 'fin@contoso.com' }];
  const r = AI.specToProject(plan, { siteUrl: 'https://contoso.sharepoint.us/sites/ops' });
  const ch = r.project.app.chain;
  assert(ch && ch.stages.length === 2, 'plan stages become a chain');
  eq(ch.stages[0].column, 'Supervisor', 'existing Person column reused');
  eq(ch.stages[1].groupId, '', 'bad group id dropped'); assert(r.notes.some((n) => /not a group ID/.test(n)), 'note about the bad id');
  assert(r.project.flows.some((f) => f.chain && /Notify Finance/.test(f.name)), 'chain flows added');
  assert(!errors(r.project).filter((f) => f.code !== 'PS054').length, errors(r.project).map((f) => f.code + f.msg).join('; '));
});
test('ask for a change: setApprovalStages rebuilds details/approvals in place', () => {
  const p = build('requests');
  const ids = Object.fromEntries(p.screens.map((s) => [s.name, s.id]));
  const browseY = E.screenYaml(p, p.screens.find((s) => s.name === 'scrBrowseRequests'));
  const ops = [{ op: 'setApprovalStages', list: 'Requests', column: 'Status', stages: [{ name: 'Manager', approverKind: 'personColumn', approverColumn: '', groupId: '', notifyEmail: '' }, { name: 'Director', approverKind: 'personColumn', approverColumn: '', groupId: '', notifyEmail: '' }] }];
  const rv = CH.reviewOps(p, ops);
  assert(!rv[0].error && rv[0].destructive, 'valid, flagged destructive');
  const res = CH.applyOps(p, ops);
  eq(res.applied.length, 1, 'applied: ' + JSON.stringify(res.skipped));
  for (const n of ['scrDetailRequests', 'scrApprovalsRequests']) eq(p.screens.find((s) => s.name === n)?.id, ids[n], n + ' keeps its id');
  eq(E.screenYaml(p, p.screens.find((s) => s.name === 'scrBrowseRequests')), browseY, 'other screens untouched');
  assert(ctlIn(p, 'scrDetailRequests', /^htmlChain/), 'stepper on details');
  assert([...core.walk(p.screens.find((s) => s.name === 'scrEditRequests').controls)].some(([c]) => c.field?.column === 'ManagerApprover'), 'approver on the form');
  p.app.siteUrl = 'https://contoso.sharepoint.us/sites/ops';
  const errs = errors(p); assert(!errs.length, errs.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
  checkApp(p);
  assert(CH.CHANGE_SCHEMA.properties.ops.items.properties.op.enum.includes('setApprovalStages'), 'op in schema');
});


// ---------------------------------------------------------------- round 4: saved blocks
const BL = await import('../js/blocks.js');
const memStore = () => { const m = new Map(); return { getItem: (k) => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), m }; };
test('blocks: save a selection, insert it elsewhere with fresh names and remapped formulas', () => {
  const p = build('requests');
  const browse = p.screens.find((s) => s.name === 'scrBrowseRequests');
  const gal = browse.controls.find((c) => c.type === 'gallery');
  const search = browse.controls.find((c) => c.name === gal.data.searchCtl);
  assert(search, 'browse has a search box');
  gal.fx.Visible = `!IsBlank(${search.name}.Text) || true`;
  const b = BL.makeBlock([search, gal], 'Search + list');
  eq(Math.min(...b.controls.map((c) => c.x)), 0, 'relative x'); eq(Math.min(...b.controls.map((c) => c.y)), 0, 'relative y');
  assert(b.lists.includes('Requests'), 'records the lists it uses');
  // round-trip through storage and an export file
  const store = memStore();
  assert(BL.saveBlocks([b], store), 'saved'); const [loaded] = BL.loadBlocks(store);
  eq(loaded.name, 'Search + list', 'name kept');
  const { blocks } = BL.importBlocksJson(BL.exportBlocksJson([loaded]));
  const home = p.screens.find((s) => s.name === 'scrHome');
  const before = core.allNames(p);
  const r = BL.insertSavedBlock(p, home, blocks[0], { x: 40, y: 400 });
  eq(r.added.length, 2, 'two top-level controls');
  for (const [c] of core.walk(r.added)) assert(!before.has(c.name), 'fresh name ' + c.name);
  const g2 = r.added.find((c) => c.type === 'gallery'), s2 = r.added.find((c) => c.type === 'textinput');
  eq(g2.data.searchCtl, s2.name, 'search link follows the rename');
  assert(g2.fx.Visible.includes(s2.name + '.Text') && !g2.fx.Visible.includes(search.name + '.Text'), 'formula remapped: ' + g2.fx.Visible);
  eq(Math.min(...r.added.map((c) => c.y)), 400, 'placed at the drop point');
  eq(r.warnings.length, 0, 'no warnings in the same app');
  checkApp(p);
  assert(!errors(p).length, errors(p).map((x) => x.code + ' ' + x.msg).join('; '));
  // into an app without that list
  const q = build('tasks');
  const r2 = BL.insertSavedBlock(q, q.screens[0], blocks[0]);
  assert(r2.warnings.some((w) => /Requests/.test(w)), 'warns about the missing list');
});
test('blocks: imports are sanitized and bad files explain themselves', () => {
  const evil = { format: BL.FORMAT, version: 1, blocks: [
    { name: '<img src=x onerror=alert(1)>', controls: [{ id: '"><script>', type: 'label', name: 'lbl: x\n  Bad', x: '5', y: 'NaN', w: 100, h: 30, props: { text: 'hi', fill: 'url(javascript:alert(1))' } }] },
    { name: 'raw', controls: [{ type: 'definitelyNot', name: 'x' }] },
    { name: 'empty', controls: [] },
  ] };
  const { blocks, skipped } = BL.importBlocksJson(JSON.stringify(evil));
  eq(blocks.length, 1, 'one usable'); eq(skipped, 2, 'two skipped');
  const c = blocks[0].controls[0];
  assert(/^[A-Za-z_]\w*$/.test(c.name), 'name made safe: ' + c.name);
  assert(/^[A-Za-z0-9_-]+$/.test(c.id), 'id made safe');
  eq(c.props.fill, 'transparent', 'bad color dropped');
  assert(!/[<>]/.test(blocks[0].name), 'block name escaped');
  for (const bad of ['not json', '{"x":1}', 'x'.repeat(BL.MAX_BYTES + 1)]) { let msg = ''; try { BL.importBlocksJson(bad); } catch (e) { msg = e.message; } assert(msg, 'rejects ' + bad.slice(0, 10)); }
  const broken = { getItem: () => { throw new Error('denied'); }, setItem: () => { throw new Error('quota'); } };
  eq(BL.loadBlocks(broken).length, 0, 'storage errors read as no blocks'); eq(BL.saveBlocks([], broken), false, 'storage errors reported');
});

// ---------------------------------------------------------------- share a Play preview
test('share link: compressed round trip keeps the app and drops secrets', async () => {
  const p = buildDemo(core.newProject);
  p.app.siteUrl = 'https://contoso.sharepoint.com/sites/Team';
  p.app.appLink = 'https://apps.powerapps.com/play/e/x/a/y?tenantId=t';
  p.flows[0].recipients = [{ kind: 'address', value: 'boss@agency.mil' }];
  p.app.onStart = 'Set(varKey, "sk-ant-api03-SECRETSECRETSECRET")';
  const r = await SH.encodeShare(p);
  assert(r.compressed && r.hash.startsWith('#play=z'), 'compressed link');
  assert(!r.tooLong, 'demo fits in a link: ' + r.chars);
  const back = await SH.decodeShare(r.hash);
  eq(back.screens.length, p.screens.length, 'screens survive');
  eq(back.lists.map((l) => l.name).join(), p.lists.map((l) => l.name).join(), 'lists survive');
  const text = JSON.stringify(back);
  assert(!text.includes('sk-ant-api03'), 'API key removed');
  assert(!text.includes('contoso.sharepoint.com') && !text.includes('tenantId=t'), 'site URL and app link left out by default');
  assert(!text.includes('boss@agency.mil'), 'fixed email addresses hidden');
  eq(errors(back).filter((x) => !['PS045', 'PS047'].includes(x.code)).length, 0, 'shared copy is still lint-clean apart from the removed setup values');
  const withLinks = await SH.decodeShare((await SH.encodeShare(p, { includeLinks: true })).hash);
  eq(withLinks.app.siteUrl, p.app.siteUrl, 'links included when ticked');
});

test('share link: plain fallback, caps, and damaged or oversized links explain themselves', async () => {
  const p = build('assets');
  p.lists[0].rows = Array.from({ length: 120 }, (_, i) => ({ ID: i + 1, Title: 'Row ' + i }));
  const r = await SH.encodeShare(p, { compress: false });
  assert(r.hash.startsWith('#play=j'), 'plain JSON fallback');
  const back = await SH.decodeShare(r.hash);
  eq(back.lists[0].rows.length, SH.MAX_ROWS, 'sample rows capped');
  const msg = async (h) => { try { await SH.decodeShare(h); return ''; } catch (e) { return e.message; } };
  assert(/damaged/.test(await msg(r.hash.slice(0, 200))), 'cut-off link');
  assert(/not a preview/.test(await msg('#other')), 'wrong hash');
  assert(/newer version/.test(await msg('#play=x' + 'A'.repeat(10))), 'unknown format');
  const big = build('requests');
  big.screens[0].controls.push(core.newControl(big, 'image', { props: { src: 'data:image/png;base64,' + Array.from({ length: 200000 }, () => 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[Math.floor(Math.random() * 26)]).join('') } }));
  assert((await SH.encodeShare(big)).tooLong, 'huge projects are refused');
});

test('share link: hostile payload is normalized like any untrusted project', async () => {
  const evil = { v: 1, p: { app: { name: '<img src=x onerror=alert(1)>' }, screens: [{ id: '"><script>', name: 'scrHome', controls: [{ id: 'x"y', type: 'label', name: 'lbl1', props: { color: 'url(javascript:alert(1))' } }] }], lists: [] } };
  const h = '#play=j' + Buffer.from(JSON.stringify(evil)).toString('base64url');
  const p = await SH.decodeShare(h);
  assert(/^[A-Za-z0-9_-]+$/.test(p.screens[0].id) && /^[A-Za-z0-9_-]+$/.test(p.screens[0].controls[0].id), 'ids made safe');
  assert(!/javascript/.test(JSON.stringify(p.screens[0].controls[0].props)), 'bad color refused');
});

// ---------------------------------------------------------------- round 5: Report a paste error
const demo = () => { const p = buildDemo(core.newProject); core.normalizeProject(p); return p; };
const firstOf = (p, type) => { for (const s of p.screens) for (const [c] of core.walk(s.controls)) if (c.type === type) return { s, c }; return null; };

test('paste-error parser: PA2108 finds the control and property and switches it off', () => {
  const p = demo(), { s, c } = firstOf(p, 'datepicker');
  const [f] = FB.analyze(p, `Error PA2108: Unknown property 'IconBackground' on control '${c.name}'.`);
  eq(f.family, 'unknownProperty', 'family'); eq(f.code, 'PA2108', 'code');
  assert(f.where.includes(c.name), 'mapped to the control: ' + f.where);
  const sup = f.actions.find((a) => a.kind === 'suppress');
  eq(sup?.control, 'Classic/DatePicker', 'control type'); eq(sup?.prop, 'IconBackground', 'property');
  assert(f.actions.some((a) => a.kind === 'select' && a.ctlId === c.id), 'select action');
  const before = E.screenYaml(p, s);
  FB.applyAction(p, sup); core.normalizeProject(p);
  const after = E.screenYaml(p, s);
  assert(before.includes('IconBackground') && !after.includes('IconBackground'), 'property no longer emitted');
  eq(after.split('\n').filter((l) => !/IconBackground/.test(l)).join('\n'), before.split('\n').filter((l) => !/IconBackground/.test(l)).join('\n'), 'nothing else changed');
  checkScreenYaml(after, new Set());
  c.fx.IconBackground = 'Color.Red';
  assert(codes(p).has('PS058'), 'override on a switched-off property is flagged');
  FB.applyAction(p, { kind: 'unsuppress', control: 'Classic/DatePicker', prop: 'IconBackground' });
  assert(!p.app.suppressProps, 'unsuppress cleans up');
});

test('paste-error parser: every error family', () => {
  const p = demo(), L = p.lists[0].name, scr = p.screens[1].name, ctl = firstOf(p, 'gallery').c.name;
  const fam = (t) => FB.analyze(p, t)[0];
  eq(fam(`PA2110 An entity with name '${ctl}' already exists`).family, 'duplicateName', 'PA2110');
  eq(fam('PA2109 Unknown control type Classic/Fancy@9.9.9').family, 'controlVersion', 'PA2109');
  eq(fam('PA1001 Error while parsing: mapping values are not allowed here').family, 'parse', 'PA1001');
  const ds = fam(`Name isn't valid. '${L}' isn't recognized.`);
  eq(ds.family, 'missingDataSource', 'list'); assert(ds.actions.some((a) => a.kind === 'dataTab'), 'data tab action');
  eq(fam(`Name isn't valid. 'Office365Users' isn't recognized.`).family, 'missingConnector', 'connector');
  eq(fam(`'appTheme' isn't recognized`).family, 'missingFormulas', 'theme');
  eq(fam(`'ThisItem' isn't recognized`).family, 'thisItem', 'ThisItem');
  eq(fam(`'${scr}' isn't recognized`).family, 'missingScreen', 'screen');
  eq(fam(`'varIsAdmin' isn't recognized`).family, 'missingVariable', 'variable');
  const fl = p.flows.find((f) => f.kind === 'button');
  if (fl) eq(fam(`'${FL.flowFxName(fl.name)}' isn't recognized`).family, 'missingFlow', 'flow');
  eq(fam('The type of this argument does not match: expected record, found text. Status').family, 'choiceAsText', 'choice');
  eq(fam('Invalid argument type (Boolean). Expecting a Record value instead.').family, 'choiceAsText', 'record expected');
  eq(fam('Invalid argument type. Expects a number.').family, 'argType', 'arg type');
  eq(fam("The specified column 'Needed By' does not exist. The column with the most similar name is 'NeededBy'.").family, 'columnName', 'column');
  const dg = fam(`Delegation warning. The "Filter" part of this formula might not work correctly on large data sets. ${L}`);
  eq(dg.family, 'delegation', 'delegation'); assert(dg.actions.some((a) => a.kind === 'largeList' && a.list === L), 'mark large action');
  eq(fam('The connection for SharePoint was not found. Select a connection during import.').family, 'flowConnection', 'connection');
  eq(fam("InvalidTemplate. Unable to process template language expressions in action 'Send_an_email_(V2)'").title, 'Flow action "Send_an_email_(V2)" failed validation', 'template');
  eq(fam('Import Package failed: Importing packages with flows is disabled').family, 'flowImport', 'import');
  eq(fam('Something completely new happened').family, 'unknown', 'unknown');
  eq(FB.analyze(p, 'PA2110 a already exists\n\nPA1001 parse failure').length, 2, 'several messages split');
});

test('findings log: logged, capped, exported, survives save/load', () => {
  const p = demo();
  FB.logFindings(p, FB.analyze(p, "PA2108 Unknown property 'Foo'"));
  for (let i = 0; i < core.MAX_FINDINGS + 5; i++) FB.logFindings(p, FB.analyze(p, 'odd ' + i));
  eq(p.findings.length, core.MAX_FINDINGS, 'capped');
  const back = core.normalizeProject(JSON.parse(JSON.stringify(p)));
  eq(back.findings.length, core.MAX_FINDINGS, 'survives load');
  const txt = FB.exportFindings(p);
  assert(txt.includes(p.app.name) && txt.includes('odd '), 'export text');
  const sh = SH.shareProject(p);
  assert(!sh.findings && !sh.snapshots, 'share link strips findings and snapshots');
});

test('suppressProps default is byte-identical and junk is dropped', () => {
  const p = demo(), base = p.screens.map((s) => E.screenYaml(p, s)).join('');
  p.app.suppressProps = {}; core.normalizeProject(p);
  assert(!('suppressProps' in p.app), 'empty map removed');
  eq(p.screens.map((s) => E.screenYaml(p, s)).join(''), base, 'identical');
  p.app.suppressProps = { 'Classic/Button': ['X', 'bad name', 'HoverFill'], '<script>': ['Fill'] }; core.normalizeProject(p);
  eq(JSON.stringify(p.app.suppressProps), '{"Classic/Button":["HoverFill"]}', 'sanitized');
});

// ---------------------------------------------------------------- round 5: bundles
test('bundle: project → zip → project is identical', async () => {
  const p = demo();
  HI.takeSnapshot(p, 'First version');
  p.deployed = { formulas: { hash: 'abc12345', at: '2026-11-24T10:00:00Z', name: 'App.Formulas' } };
  core.normalizeProject(p);
  const { name, files, root } = BU.bundleFiles(p, { date: '2026-11-24' });
  assert(name === `${root}.zip` && /_bundle_2026-11-24$/.test(root), 'dated name: ' + name);
  for (const k of ['canonical/App.pa.yaml', 'schema/schema.json', 'reports/lint.txt', 'reports/drift.txt', 'CHANGELOG.md', 'bundle.json']) assert(files[`${root}/${k}`], 'has ' + k);
  for (const s of p.screens) checkScreenYaml(files[`${root}/canonical/${s.name}.pa.yaml`], new Set());
  assert(Object.keys(files).some((k) => /\/flows\/[^/]+\/definition\.json$/.test(k)), 'flow definitions');
  assert(files[`${root}/CHANGELOG.md`].includes('First version'), 'changelog from snapshots');
  yaml.load(files[`${root}/canonical/App.pa.yaml`]);
  const z = new JSZip(); for (const [k, v] of Object.entries(files)) z.file(k, v);
  const bytes = await z.generateAsync({ type: 'uint8array' });
  const r = await BU.readBundle(bytes, JSZip, core.newProject);
  eq(r.source, 'embedded', 'restored from bundle.json');
  eq(JSON.stringify(r.project), JSON.stringify(p), 'identical project');
});

test('bundle: schema digest shape and import without an embedded project', async () => {
  const p = demo();
  p.lists[0].columns.push(core.newColumn('Needed Date', 'DateTime'));
  const d = BU.schemaDigest(p);
  const col = d.lists[0].columns.find((c) => c.display === 'Needed Date');
  eq(col.name, 'NeededDate', 'internal name'); eq(col.type, 'DateTime', 'type');
  eq(d.lists[0].columns[0].name, 'Title', 'Title first');
  assert(d.lists.flatMap((l) => l.columns).every((c) => ['name', 'display', 'type', 'required'].every((k) => k in c)), 'digest fields');
  const yesno = BU.listsFromDigest({ lists: [{ name: 'T', columns: [{ name: 'Flag', type: 'Boolean' }, { name: 'Who', type: 'User' }, { name: 'Title', type: 'Text', required: false }] }] })[0];
  eq(yesno.columns.map((c) => c.type).join(), 'Text,YesNo,Person', 'SharePoint types mapped back');
  const { files, root } = BU.bundleFiles(p);
  const meta = JSON.parse(files[`${root}/bundle.json`]); delete meta.project; files[`${root}/bundle.json`] = JSON.stringify(meta);
  const z = new JSZip(); for (const [k, v] of Object.entries(files)) z.file(k, v);
  const r = await BU.readBundle(await z.generateAsync({ type: 'uint8array' }), JSZip, core.newProject);
  eq(r.source, 'canonical', 'rebuilt from canonical');
  eq(r.project.screens.length, p.screens.length, 'all screens');
  eq(r.project.lists.map((l) => l.name).join(), p.lists.map((l) => l.name).join(), 'lists from schema');
  eq(p.screens.find((s) => s.id === p.app.startScreen)?.name || p.screens[0].name, r.project.screens.find((s) => s.id === r.project.app.startScreen)?.name, 'start screen');
  let msg = ''; const z2 = new JSZip(); z2.file('x.txt', 'hi');
  try { await BU.readBundle(await z2.generateAsync({ type: 'uint8array' }), JSZip, core.newProject); } catch (e) { msg = e.message; }
  assert(/bundle/.test(msg), 'non-bundle explained');
});

// ---------------------------------------------------------------- round 5: large lists
test('large lists: warnings for non-delegable formulas', () => {
  const L = new Set(['Requests']);
  const has = (fx) => DG.delegationIssues(fx, L).length > 0;
  assert(has('Text(CountRows(Requests))') === true && !has(DG.capCounts('Text(CountRows(Requests))', L)), 'CountRows flagged unless capped');
  assert(has('Search(Requests, txt.Text, "Title")'), 'Search');
  assert(has('Filter(Requests, Status.Value <> "Closed")'), '<>');
  assert(has('Filter(Requests, Title in colPicked)'), 'in');
  assert(has('Filter(Requests, DueDate = Blank())'), '= Blank()');
  assert(!has('Filter(Requests, StartsWith(Title, x.Text) && Status.Value = "Open" && \'Created By\'.Email = User().Email)'), 'delegable shapes pass');
  assert(!has('Filter(Other, Status.Value <> "x")'), 'other lists untouched');
  assert(!has('Filter(Filter(Requests, Owner.Email = User().Email), Status.Value <> "x")'), 'skill shape passes');
});

test('large lists: generated apps stay lint-clean and paste-rule clean; default unchanged', () => {
  for (const sc of Object.keys(SCENARIOS)) for (const ff of ['tablet', 'phone']) {
    const base = buildFt(sc, ff), big = buildFt(sc, ff);
    for (const l of big.lists) l.size = 'large';
    const errs = E.lint(big).filter((x) => x.level !== 'info');
    eq(errs.length, 0, `${sc} ${ff} large: ` + errs.map((x) => `${x.code} ${x.where} ${x.msg}`).join(' | '));
    const seen = new Set();
    for (const s of big.screens) checkScreenYaml(E.screenYaml(big, s), seen);
    for (const l of base.lists) l.size = 'small';
    core.normalizeProject(base);
    const plain = buildFt(sc, ff);
    eq(base.screens.map((s) => E.screenYaml(base, s)).join(''), plain.screens.map((s) => E.screenYaml(plain, s)).join(''), `${sc} ${ff}: small is byte-identical`);
  }
  const d = demo(); for (const l of d.lists) l.size = 'large';
  const home = d.screens.find((s) => s.name === 'scrHome');
  const y = E.screenYaml(d, home);
  assert(y.includes('"2000+"') && !/Text\(CountRows\(Filter\(/.test(y), 'KPI counts capped');
  checkScreenYaml(y, new Set());
  const g = firstOf(d, 'gallery').c; g.data.extraFilter = 'Status.Value <> "Closed"';
  const items = E.itemsFx(d, g);
  assert(/Filter\((?:Filter|SortByColumns)\(/.test(items) && !DG.delegationIssues(items, DG.largeLists(d)).length, 'non-delegable extra filter reshaped: ' + items);
  g.data.extraFilter = 'Status.Value = "Open"';
  assert(!/Filter\(Filter\(/.test(E.itemsFx(d, g)), 'delegable extra filter stays inside');
  d.app.setup = {}; core.normalizeProject(d);
  assert(SU.setupItems(d).some((i) => i.id === 'rowLimit' && !i.done), 'row limit setup item');
  SU.applySetup(d, 'rowLimit', true);
  assert(SU.setupItems(d).find((i) => i.id === 'rowLimit').done, 'confirmed');
});

// ---------------------------------------------------------------- round 5: version history
test('history: snapshot, diff, restore and cap', () => {
  const p = demo();
  HI.takeSnapshot(p, 'Start');
  const l = p.lists[0]; l.columns.push(core.newColumn('CostCenter', 'Text'));
  const g = firstOf(p, 'gallery').c; g.props.templateSize += 4;
  p.flows.push(FL.newFlow(p, 'created', l.name)); p.flows.at(-1).name = 'New flow';
  p.theme.colors.Primary = '#123456';
  const snap = p.snapshots[0];
  const d = HI.diffProjects(HI.snapshotProject(snap), p);
  const lines = HI.describeDiff(d).join('\n');
  assert(lines.includes('CostCenter'), 'column added: ' + lines);
  assert(lines.includes(g.name), 'control changed');
  assert(lines.includes('New flow'), 'flow added');
  assert(lines.includes('Theme changed'), 'theme');
  assert(HI.diffIsEmpty(HI.diffProjects(p, JSON.parse(JSON.stringify(p)))), 'no diff with itself');
  const r = core.normalizeProject(HI.restoreSnapshot(p, snap.id));
  assert(!r.lists[0].columns.some((c) => c.name === 'CostCenter'), 'restored state');
  eq(r.snapshots.length, 2, 'snapshot of the replaced state added');
  assert(r.snapshots[1].name.startsWith('Before restoring'), 'named');
  for (let i = 0; i < core.MAX_SNAPSHOTS + 3; i++) HI.takeSnapshot(r, 'auto ' + i, { auto: true });
  eq(r.snapshots.length, core.MAX_SNAPSHOTS, 'capped');
  assert(r.snapshots.some((x) => x.name === 'Start'), 'manual snapshot kept; automatic ones dropped first');
  const sizeOk = HI.snapshotsBytes(r) > 0; assert(sizeOk, 'size measured');
  const n = core.normalizeProject(JSON.parse(JSON.stringify({ ...r, snapshots: [...r.snapshots, { data: 'not json' }, null] })));
  eq(n.snapshots.length, core.MAX_SNAPSHOTS, 'junk snapshots dropped on load');
  assert(!JSON.parse(n.snapshots[0].data).snapshots, 'snapshots do not nest');
});

// ---------------------------------------------------------------- round 6: scenario gallery
const SC6 = await import('../js/scenarios.js');
const NEW_SCENARIOS = ['workorders', 'checkout', 'incidents', 'onboarding', 'leave', 'policies'];
/** What Finish setup fills in, so lint judges the app rather than the missing setup values. */
function setupDone(p) {
  p.app.siteUrl = 'https://contoso.sharepoint.com/sites/Team';
  p.app.appLink = 'https://apps.powerapps.com/play/e/env/a/app?tenantId=t';
  for (const st of p.app.chain?.stages || []) if (st.approver === 'group' && !st.groupId) st.groupId = '11111111-2222-4333-8444-555555555555';
  if (p.app.features?.usage) p.app.setup = { ...(p.app.setup || {}), usageConfirmed: true };
  return p;
}
test('round 6: every new scenario builds a lint-clean, paste-safe app (tablet + phone, options on and off)', () => {
  const variants = [{}, { loading: false, confirm: false, stepper: false, scope: 'all' }, { loading: true, confirm: true, stepper: true, scope: 'mine', usage: true, adminGroup: '11111111-2222-4333-8444-555555555555' }];
  for (const k of NEW_SCENARIOS) for (const ff of ['tablet', 'phone']) for (const v of variants) {
    const p = setupDone(SC6.buildScenario(k, { formFactor: ff, features: v }));
    const bad = E.lint(p).filter((x) => x.level !== 'info');
    eq(bad.length, 0, `${k} ${ff} ${JSON.stringify(v)}: ` + bad.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
    const seen = new Set();
    for (const s of p.screens) checkScreenYaml(E.screenYaml(p, s), seen);
    for (const f of p.flows || []) checkExpressions(FL.flowDefinition(p, f));
  }
});
test('round 6: scenarios carry their flows, chain and attachments', () => {
  const wo = SC6.buildScenario('workorders');
  eq(wo.flows.length, 3, 'work order flows');
  assert(wo.flows.every((f) => f.recipients.length), 'every flow has a recipient');
  assert(wo.app.deepLink?.screen, 'email links open the details screen');
  const lv = SC6.buildScenario('leave');
  eq(lv.app.chain?.stages.length, 2, 'leave has a 2-stage chain');
  assert(lv.flows.some((f) => f.chain), 'chain flows added');
  const ob = SC6.buildScenario('onboarding');
  eq(ob.lists.find((l) => l.name === 'OnboardingTasks').columns.find((c) => c.name === 'Hire').lookupList, 'NewHires', 'lookup to hires');
  assert(SC6.buildScenario('policies').lists[0].attachments === true, 'policies turn attachments on');
  assert(/flows/.test(SC6.scenarioIncludes('incidents')) && /2-step approvals/.test(SC6.scenarioIncludes('leave')), 'gallery summaries');
  eq(SC6.buildScenario('checkout', { flows: false }).flows.length, 0, 'flows can be left out');
});
test('round 6: flows of a scenario survive the user removing their list', () => {
  const p = core.newProject(); p.lists = SCENARIOS.checkout.lists().filter((l) => l.name === 'Equipment');
  generateApp(p, {});
  const added = SC6.addScenarioFlows(p, SCENARIOS.checkout);
  eq(added.length, 0, 'flows for the removed Checkouts list are skipped');
});

// ---------------------------------------------------------------- round 6: attachments on forms
const AT = await import('../js/attach.js');
test('round 6 attachments: off by default, nothing names a form', () => {
  for (const k of ['requests', 'assets', 'events', 'tasks', 'contacts', 'workorders', 'leave']) {
    const p = SC6.buildScenario(k);
    for (const s of p.screens) assert(!/frmAttach/.test(E.screenYaml(p, s)), `${k} ${s.name} names no attachments form`);
    assert(!/EnableAttachments/.test(E.provisioningScript(p)), `${k}: provisioning untouched`);
    assert(!DP.deploySteps(p).some((x) => x.kind === 'attach'), `${k}: no attachments step`);
  }
});
test('round 6 attachments: edit screen loads, saves and resets the hand-added form; details only shows it', () => {
  for (const ff of ['tablet', 'phone']) {
    const p = setupDone(SC6.buildScenario('policies', { formFactor: ff }));
    const edit = p.screens.find((x) => x.name === 'scrEditPolicies'), det = p.screens.find((x) => x.name === 'scrDetailPolicies');
    const ey = E.screenYaml(p, edit), dy = E.screenYaml(p, det);
    assert(/Patch\(Policies, [^\n]*\n(?:[^\n]*\n)*?\s*\}, frmAttachPolicies\.Updates\)/.test(ey), 'Save merges frmAttachPolicies.Updates as an extra Patch argument');
    assert(ey.includes('If(IsBlank(varSelPolicies.ID), NewForm(frmAttachPolicies), EditForm(frmAttachPolicies))'), 'loader in OnVisible');
    assert(ey.includes('ResetForm(frmAttachPolicies)'), '"Add another" clears the files');
    assert(/htmlAttachEditPolicies/.test(ey) && /frmAttachViewPolicies/.test(dy), 'placeholders on both screens');
    eq(dy.split('\n').filter((l) => /frmAttachView/.test(l) && !/HtmlText|<div/.test(l)).length, 0, 'no formula on details names the view form');
    const lint = E.lint(p);
    eq(lint.filter((x) => x.level !== 'info').length, 0, ff + ' lint clean: ' + lint.filter((x) => x.level !== 'info').map((x) => x.code + ' ' + x.where).join('; '));
    eq(lint.filter((x) => x.code === 'PS060').length, 2, 'PS060 reminder per placeholder');
    const seen = new Set();
    for (const s of p.screens) checkScreenYaml(E.screenYaml(p, s), seen);
    assert(/Set-PnPList -Identity 'Policies' -EnableAttachments \$true/.test(E.provisioningScript(p)), 'provisioning turns attachments on');
  }
});
test('round 6 attachments: Paste assistant step, paste-error report, normalize', () => {
  const p = setupDone(SC6.buildScenario('policies'));
  const steps = DP.deploySteps(p);
  const at = steps.filter((x) => x.kind === 'attach');
  eq(at.length, 2, 'one step per screen');
  const i = steps.findIndex((x) => x.key === 'attach:' + at[1].screen.id);
  eq(steps[i - 1].key, 'scr:' + at[1].screen.id, 'right after its screen');
  const how = at.find((x) => x.screen.name === 'scrEditPolicies').how;
  for (const bit of ['Rename it to frmAttachPolicies', 'Data source: Policies', 'remove every field except Attachments', 'Item: =varSelPolicies', 'DefaultMode: =FormMode.Edit', 'Unlock', 'DisplayMode: =DisplayMode.Edit', 'Notepad']) assert(how.includes(bit), 'instructions: ' + bit);
  assert(steps.find((x) => x.key === 'scr:' + at[0].screen.id).how.includes('Notepad'), 'screen step warns about re-pastes');
  for (const s of steps) DP.markDeployed(p, s);
  const edit = p.screens.find((x) => x.name === 'scrEditPolicies');
  edit.controls.find((c) => c.name.startsWith('btnCancel')).props.text = 'Close';
  const pend = DP.pendingSteps(p).map((x) => x.key);
  assert(pend.includes('scr:' + edit.id) && pend.includes('attach:' + edit.id), 're-pasting the screen brings the attachments step back');
  const f = FB.analyze(p, "Name isn't valid. 'frmAttachPolicies' isn't recognized.")[0];
  eq(f.family, 'attachForm', 'recognized');
  assert(f.where.startsWith('scrEditPolicies') && f.actions.some((a) => a.kind === 'select'), 'points at the placeholder');
  const n = core.normalizeProject(JSON.parse(JSON.stringify({ ...p, lists: p.lists.map((l) => ({ ...l, attachments: 'yes' })) })));
  assert(!n.lists[0].attachments, 'only true turns attachments on');
  const bad = JSON.parse(JSON.stringify(p));
  bad.screens[0].controls.push({ id: 'x1', type: 'html', name: 'htmlX', props: {}, attachSlot: { list: 7, mode: 'oops' } });
  assert(!core.normalizeProject(bad).screens[0].controls.at(-1).attachSlot, 'junk placeholder metadata dropped');
});
test('round 6 attachments: toggling on an existing app adds, rebuilds or removes the placeholders', () => {
  for (const ff of ['tablet', 'phone']) {
    const p = setupDone(SC6.buildScenario('requests', { formFactor: ff }));
    const ids = p.screens.map((s) => s.id + s.name).join();
    p.lists[0].attachments = true;
    const tight = ED.attachNoRoom(p, 'Requests');
    const r = ED.syncAttachSlots(p, 'Requests', { rebuild: true });
    eq(r.rebuilt.length, tight.length, 'screens without room are rebuilt');
    eq(r.added.length + r.rebuilt.length, 2, 'edit and details both get the box');
    eq(p.screens.map((s) => s.id + s.name).join(), ids, 'same screen ids and names');
    eq(E.lint(p).filter((x) => x.level !== 'info').length, 0, ff + ' lint clean after the toggle');
    assert(E.screenYaml(p, p.screens.find((x) => x.name === 'scrEditRequests')).includes('frmAttachRequests.Updates'), 'save uses the form');
    delete p.lists[0].attachments;
    const off = ED.syncAttachSlots(p, 'Requests');
    eq(off.removed.length, 2, 'off removes both');
    assert(p.screens.every((s) => !/frmAttach/.test(E.screenYaml(p, s))), 'no reference left');
  }
  // without rebuild the box still lands behind every clickable control
  const q = setupDone(SC6.buildScenario('requests'));
  q.lists[0].attachments = true;
  ED.syncAttachSlots(q, 'Requests');
  assert(!E.lint(q).some((x) => x.code === 'PS040'), 'never covers a button');
  // a list rename follows the placeholder and its form name
  const p2 = SC6.buildScenario('policies');
  ED.renameListIn(p2, 'Policies', 'Procedures');
  const ey = E.screenYaml(p2, p2.screens.find((x) => /^scrEdit/.test(x.name)));
  assert(ey.includes('frmAttachProcedures.Updates') && ey.includes('<b>frmAttachProcedures</b>'), 'renamed form');
});
test('round 6 attachments: Play keeps files with the item', () => {
  const p = SC6.buildScenario('policies');
  const rt = createRuntime(p, { toast: () => {}, rerender: () => {} });
  assert(rt.db.Policies.some((r) => r.Attachments?.length), 'sample rows carry files');
  const edit = p.screens.find((x) => x.name === 'scrEditPolicies');
  const slot = edit.controls.find((c) => c.attachSlot);
  const save = edit.controls.find((c) => c.name.startsWith('btnSave'));
  rt.vars.varSelPolicies = rt.db.Policies[0];
  rt.inputs[slot.name] = [...(rt.db.Policies[0].Attachments || []), { DisplayName: 'new.pdf' }];
  rt.run(edit, save.actions.OnSelect);
  assert(rt.db.Policies[0].Attachments.some((f) => f.DisplayName === 'new.pdf'), 'saved with the item');
});

// ---------------------------------------------------------------- round 6: approval-chain extras
const ADM = '11111111-2222-4333-8444-555555555555';
/** The leave scenario with every extra on (deadlines 3 and 2 days, reassign, admin override). */
function leaveExtras(ff = 'tablet', adminGroup = ADM) {
  const sc = SCENARIOS.leave, orig = sc.chain;
  sc.chain = { ...orig, reassign: true, override: true, stages: orig.stages.map((st, i) => ({ ...st, slaDays: i ? 2 : 3 })) };
  try { return setupDone(SC6.buildScenario('leave', { formFactor: ff, features: { adminGroup } })); } finally { sc.chain = orig; }
}
test('round 6 chain extras: off by default, every extra opt-in', () => {
  const p = SC6.buildScenario('leave');
  const l = p.lists.find((x) => x.name === 'LeaveRequests');
  assert(!l.columns.some((c) => /Delegate$|^StageSince$/.test(c.name)), 'no extra columns');
  assert(!p.flows.some((f) => f.chainStage), 'no reminder flows');
  assert(!p.screens.some((s) => [...core.walk(s.controls)].some(([c]) => /^(btnChainMore|lblOverdue|btnPanelDetails)/.test(c.name))), 'no extra controls');
  const n = core.normalizeChain({ list: 'X', stages: [{ name: 'A', slaDays: '400' }, { name: 'B', slaDays: -2 }], reassign: 'yes', override: 1 });
  eq(n.stages[0].slaDays, 90, 'deadline capped'); assert(!('slaDays' in n.stages[1]) && !n.reassign && !n.override, 'only real values turn extras on');
});
test('round 6 chain extras: lint-clean, paste-safe screens and formulas (tablet + phone)', () => {
  for (const ff of ['tablet', 'phone']) {
    const p = leaveExtras(ff);
    const bad = E.lint(p).filter((x) => x.level !== 'info');
    eq(bad.length, 0, ff + ': ' + bad.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
    const seen = new Set();
    for (const s of p.screens) checkScreenYaml(E.screenYaml(p, s), seen);
    for (const f of p.flows) checkExpressions(FL.flowDefinition(p, f));
    const l = p.lists.find((x) => x.name === 'LeaveRequests');
    for (const c of ['SupervisorDelegate', 'HRDelegate', 'StageSince']) assert(l.columns.some((x) => x.name === c && x.chain), c + ' column (hidden from forms)');
    const det = E.screenYaml(p, p.screens.find((s) => s.name === 'scrDetailLeaveRequests'));
    assert(det.includes('SupervisorDelegate: {\'@odata.type\'') && det.includes('Reassigned to " & varPickppReassign'), 'reassign writes the stand-in and a history line');
    assert(det.includes('Approved (admin override)') && det.includes('Give a reason for the override.'), 'override needs a reason and is logged');
    assert(det.includes('StageSince: Today()'), 'decisions restart the stage clock');
    assert(/overdue/.test(det), 'details say when a stage is overdue');
    const ap = E.screenYaml(p, p.screens.find((s) => s.name === 'scrApprovalsLeaveRequests'));
    assert(ap.includes('SupervisorDelegate.Email = User().Email') && ap.includes('lblOverdue'), 'queue includes stand-ins and the overdue badge');
    const edit = E.screenYaml(p, p.screens.find((s) => s.name === 'scrEditLeaveRequests'));
    assert(edit.includes('StageSince: Coalesce(LookUp(LeaveRequests, ID = varSelLeaveRequests.ID).StageSince, Today())'), 'new items start the clock');
    const rem = p.flows.filter((f) => f.chainStage);
    eq(rem.length, 2, 'a reminder per stage with a deadline');
    const q = FL.flowDefinition(p, rem[0]).actions.Get_items.inputs.parameters.$filter;
    assert(q.startsWith('CurrentStage eq 1 and StageSince lt'), 'reminder filter: ' + q);
    assert(rem[1].recipients.some((r) => r.value === 'HRDelegate'), 'group stage reminder goes to its stand-in');
  }
  const noAdmin = leaveExtras('tablet', '');
  assert(E.lint(noAdmin).some((x) => x.code === 'PS061'), 'override without an admin group is flagged');
});
test('round 6 chain extras: Play reassigns, overrides and shows overdue items', () => {
  const p = leaveExtras();
  const ch = p.app.chain;
  const rt = createRuntime(p, { toast: () => {}, rerender: () => {} });
  const rows = rt.db.LeaveRequests;
  assert(rows.some((r) => CHN.playOverdue(ch, r) > 0), 'sample data has overdue items');
  const det = p.screens.find((s) => s.name === 'scrDetailLeaveRequests');
  const r = rows.find((x) => x.CurrentStage === 1 && x.Author === core.PEOPLE[0]);
  assert(r && !CHN.playCanDecide(ch, r, rt.vars), 'own request cannot be decided');
  rt.vars.varSelLeaveRequests = r;
  const ppName = [...core.walk(det.controls)].map(([c]) => c).find((c) => c.name.startsWith('ppReassign')).name;
  rt.inputs[ppName] = core.PEOPLE[4];
  rt.run(det, [{ type: 'reassign', list: 'LeaveRequests', picker: ppName }]);
  eq(r.SupervisorDelegate?.DisplayName, core.PEOPLE[4].DisplayName, 'stand-in set');
  assert(r.ApprovalHistory.split('\n')[0].includes('Reassigned to ' + core.PEOPLE[4].DisplayName), 'history line');
  assert(CHN.playCanDecide(ch, r, rt.vars, core.PEOPLE[4].Mail), 'the stand-in may decide');
  rt.vars.varIsAdmin = true;
  const reason = [...core.walk(det.controls)].map(([c]) => c).find((c) => c.name.startsWith('txtOverrideReason')).name;
  rt.run(det, [{ type: 'override', list: 'LeaveRequests', decision: 'Approved', commentCtl: reason }]);
  eq(r.CurrentStage, 1, 'no reason, no override');
  rt.inputs[reason] = 'Supervisor on leave';
  rt.run(det, [{ type: 'override', list: 'LeaveRequests', decision: 'Approved', commentCtl: reason }]);
  eq(r.CurrentStage, 2, 'override moved it on');
  assert(r.ApprovalHistory.split('\n')[0].includes('Approved (admin override)  ·  Supervisor on leave'), 'override logged with its reason');
  eq(r.StageSince, new Date().toISOString().slice(0, 10), 'stage clock restarted');
});
test('round 6 chain extras: Ask for a change and plans carry them', () => {
  const p = build('requests');
  const ops = [{ op: 'setApprovalStages', list: 'Requests', column: 'Status', reassign: true, adminOverride: false, stages: [{ name: 'Manager', approverKind: 'personColumn', approverColumn: '', groupId: '', notifyEmail: '', slaDays: 2 }, { name: 'Director', approverKind: 'personColumn', approverColumn: '', groupId: '', notifyEmail: '', slaDays: 0 }] }];
  eq(CH.applyOps(p, ops).applied.length, 1, 'applied');
  assert(p.app.chain.reassign && !p.app.chain.override && p.app.chain.stages[0].slaDays === 2 && !p.app.chain.stages[1].slaDays, 'chain settings');
  assert(p.flows.some((f) => f.chainStage === 1), 'reminder flow added');
  p.app.siteUrl = 'https://contoso.sharepoint.us/sites/ops';
  const errs = errors(p); assert(!errs.length, errs.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
  assert(CH.CHANGE_SCHEMA.properties.ops.items.properties.reassign && AI.STAGE_SPEC.properties.slaDays, 'schemas');
  const plan = JSON.parse(JSON.stringify(FIXTURE));
  plan.screens.approvals.stages = [{ name: 'Manager', approverKind: 'personColumn', approverColumn: '', groupId: '', notifyEmail: '', slaDays: 4 }, { name: 'Finance', approverKind: 'personColumn', approverColumn: '', groupId: '', notifyEmail: '', slaDays: 0 }];
  plan.screens.approvals.reassign = true; plan.screens.approvals.adminOverride = true;
  const { project: q, notes } = AI.specToProject(plan, { siteUrl: 'https://contoso.sharepoint.com/sites/Team' });
  assert(q.app.chain?.reassign && q.app.chain.override && q.app.chain.stages[0].slaDays === 4, 'plan → chain extras');
  assert(notes.some((n) => /admin group/.test(n)) || q.app.features.adminGroup, 'override without a group is explained');
});

// ---------------------------------------------------------------- round 6: polish (lint size actions, versions, bundle merge)
const VST = await import('../js/versionstore.js');
test('round 6 polish: lint findings name the list behind PS035 and PS059', () => {
  const p = build('requests');
  const f35 = E.lint(p).find((x) => x.code === 'PS035');
  if (f35) eq(f35.list, 'Requests', 'PS035 carries its list');
  p.lists[0].size = 'large';
  p.screens.find((s) => s.name === 'scrBrowseRequests').controls.find((c) => c.type === 'gallery').fx.Items = 'Search(Requests, "x", "Title")';
  const f59 = E.lint(p).find((x) => x.code === 'PS059');
  assert(f59 && f59.list === 'Requests', 'PS059 carries its list');
});
test('round 6 polish: YAML line diff between two versions', () => {
  const rows = HI.lineDiff('a\nb\nc\nd\ne\nf\ng', 'a\nb\nX\nd\ne\nf\ng\nh', 1);
  eq(rows.map((r) => r.t).join(''), '… -+ … +', 'folded rows, one line of context');
  const a = build('requests'), b = JSON.parse(JSON.stringify(a));
  b.screens.find((s) => s.name === 'scrEditRequests').controls.find((c) => c.name.startsWith('btnSave')).props.text = 'Send';
  const d = HI.screenYamlDiffs(a, b, E.screenYaml);
  eq(d.length, 1, 'one screen differs'); eq(d[0].name, 'scrEditRequests', 'the edit screen');
  assert(d[0].rows.some((r) => r.t === '+' && r.s.includes('"Send"')) && d[0].rows.some((r) => r.t === '-' && r.s.includes('"Save"')), 'the changed line');
});
test('round 6 polish: versions stored apart from the project (localStorage fallback without IndexedDB)', async () => {
  const mem = new Map(), ls = { getItem: (k) => mem.get(k) ?? null, setItem: (k, v) => mem.set(k, String(v)), removeItem: (k) => mem.delete(k) };
  const p = build('requests');
  HI.takeSnapshot(p, 'one'); HI.takeSnapshot(p, 'two');
  eq(await VST.writeVersions(p.snapshots, { idb: null, ls }), 'ls', 'falls back to localStorage');
  const back = await VST.readVersions({ idb: null, ls });
  eq(back.map((x) => x.name).join(), 'one,two', 'read back');
  eq(VST.versionsSig(back), VST.versionsSig(p.snapshots), 'same signature');
  await VST.writeVersions([], { idb: null, ls });
  eq((await VST.readVersions({ idb: null, ls })).length, 0, 'emptied');
  eq(await VST.writeVersions(p.snapshots, { idb: null, ls: null }), '', 'reports when nothing can store them');
});
test('round 6 polish: a bundle can be added to the open app', () => {
  const t = setupDone(SC6.buildScenario('requests')), src = SC6.buildScenario('requests');
  src.lists[0].columns.push(core.newColumn('Extra', 'Text'));
  const before = t.screens.length;
  const r = BU.mergeProjects(t, src);
  eq(r.lists.join(), 'Requests2', 'a different list with the same name is renamed');
  eq(t.screens.length, before + src.screens.length, 'screens added');
  const names = [...core.allNames(t)]; eq(new Set(names).size, names.length, 'names unique');
  const bad = E.lint(t).filter((x) => x.level === 'error');
  eq(bad.length, 0, bad.map((x) => `${x.code} ${x.where} ${x.msg}`).join('; '));
  const seen = new Set(); for (const s of t.screens) checkScreenYaml(E.screenYaml(t, s), seen);
  const t2 = SC6.buildScenario('requests'), s2 = SC6.buildScenario('checkout');
  const r2 = BU.mergeProjects(t2, s2);
  eq(r2.lists.join(), 'Equipment,Checkouts', 'new lists come in as they are');
  assert(t2.flows.length >= 2, 'flows added');
});

// ---------------------------------------------------------------- tenant test 1 fixes
test('every HtmlViewer has slack: its first box fits in Width-10 x Height-10 with 2px to spare (no scrollbars)', () => {
  let n = 0;
  const all = Object.keys(SCENARIOS);
  for (const k of all) for (const ff of ['tablet', 'phone']) {
    const p = core.newProject(); const sc = SCENARIOS[k]; p.app.name = sc.app; p.app.formFactor = ff; p.lists = sc.lists();
    p.app.features = { loading: true, usage: true, confirm: true, stepper: true, scope: 'all', adminGroup: '11111111-2222-4333-8444-555555555555' };
    generateApp(p, { screens: ['home', 'browse', 'detail', 'edit'], approvals: sc.approvals || null });
    for (const s of p.screens) {
      const doc = yaml.load(E.screenYaml(p, s), { schema: yaml.FAILSAFE_SCHEMA });
      const walkY = (ch) => (ch || []).forEach((it) => {
        const [[name, c]] = Object.entries(it);
        if (c.Control === 'HtmlViewer') {
          const P = c.Properties, W = Number(String(P.Width).replace('=', '')), H = Number(String(P.Height).replace('=', ''));
          if (!Number.isNaN(W) && !Number.isNaN(H)) {
            const html = String(P.HtmlText);
            const first = (/style='([^']*)'/.exec(html) || [, ''])[1];
            const w = /width:(\d+)px/.exec(first), h = /height:(\d+)px/.exec(first), m = /margin:(\d+)px/.exec(first);
            const mg = m ? 2 * Number(m[1]) : 0;
            assert(/overflow:hidden/.test(first), `${k}/${ff} ${s.name}.${name}: first box clips`);
            assert(h && Number(h[1]) + mg <= H - 12, `${k}/${ff} ${s.name}.${name}: height ${h && h[1]}+${mg} vs control ${H}`);
            assert(w && Number(w[1]) + mg <= W - 12, `${k}/${ff} ${s.name}.${name}: width ${w && w[1]}+${mg} vs control ${W}`);
            n++;
          }
        }
        walkY(c.Children);
      });
      walkY(Object.values(doc.Screens)[0].Children);
    }
  }
  assert(n > 200, 'checked many HtmlViewers: ' + n);
});
test('flows guard a blank Person recipient instead of failing with "To field cannot be empty"', () => {
  const p = build('requests'); p.app.siteUrl = 'https://x.sharepoint.com/sites/a';
  const f = FL.newFlow(p, 'created', 'Requests');
  f.recipients = [{ kind: 'column', value: 'Approver' }];
  p.lists[0].columns.find((c) => c.name === 'Approver').required = false;
  p.flows = [f];
  let d = FL.flowDefinition(p, f);
  assert(d.actions.Compose_recipients && d.actions.Check_recipients, 'compose + condition');
  eq(d.actions.Check_recipients.actions['Send_an_email_(V2)'].inputs.parameters['emailMessage/To'], "@{outputs('Compose_recipients')}", 'sends to the composed line');
  eq(Object.keys(d.actions.Check_recipients.else.actions).length, 0, 'no fallback: skip');
  assert(E.lint(p).some((x) => x.code === 'PS062'), 'optional recipient noted');
  f.fallback = 'author'; d = FL.flowDefinition(p, f);
  assert(JSON.stringify(d.actions.Check_recipients.else.actions.Send_to_fallback).includes("body/Author/Email"), 'fallback to Created By');
  f.recipients = [{ kind: 'author', value: '' }]; d = FL.flowDefinition(p, f);
  assert(d.actions['Send_an_email_(V2)'] && !d.actions.Compose_recipients, 'always-present recipients stay a plain send');
  const b = FL.newFlow(p, 'button', 'Requests'); b.recipients = [{ kind: 'column', value: 'Approver' }];
  d = FL.flowDefinition(p, b);
  eq(JSON.stringify(d.actions.Respond_to_a_PowerApp_or_flow.runAfter), '{"Check_recipients":["Succeeded"]}', 'button flow responds after the check');
  checkExpressions(d);
});

await Promise.all(pending);
console.log(`\n\n${pass} passed, ${fail} failed`);
if (failures.length) { console.log(failures.join('\n')); process.exit(1); }
