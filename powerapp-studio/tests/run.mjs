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
const TB = await import('../js/tabular.js');
const A11 = await import('../js/a11y.js');
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

await Promise.all(pending);
console.log(`\n\n${pass} passed, ${fail} failed`);
if (failures.length) { console.log(failures.join('\n')); process.exit(1); }
