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
const fs = await import('node:fs');
const FIXTURE = JSON.parse(fs.readFileSync(new URL('./fixture-plan.json', import.meta.url), 'utf8'));

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
  assert(toasts.some(([t, k]) => k === 'error' && /Title/.test(t)), 'error toast names the field');
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

await Promise.all(pending);
console.log(`\n\n${pass} passed, ${fail} failed`);
if (failures.length) { console.log(failures.join('\n')); process.exit(1); }
