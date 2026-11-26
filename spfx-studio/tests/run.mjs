// SPFx Studio tests: node spfx-studio/tests/run.mjs
// Package shape against a toolchain-built reference (SPFx 1.21.1 gulp package-solution --ship), bundle contract,
// CSS scoping, spec cleaning, source project. The browser side (preview host, REST emulator, UI) is covered by the
// Playwright walkthrough described in the README; tests/harness.html mounts any example or bundle in the preview.
import vm from 'vm';
import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import { EXAMPLES } from '../js/examples.js';
import { buildWebPart, bundleText, manifest, scopeCss, names, DEPS, hash20, pascal, kebab } from '../js/bundle.js';
import { packageFiles } from '../js/sppkg.js';
import { sourceFiles } from '../js/source.js';
import { cleanSpec, checkCode, SPEC_SCHEMA, SYSTEM_PROMPT, buildRequest, changeHistory, ICON_NAMES } from '../js/ai.js';
import { RUNTIME_SOURCE } from '../js/runtime.js';

const here = path.dirname(fileURLToPath(import.meta.url));
let passed = 0; const failed = [];
const test = (name, fn) => { try { fn(); passed++; process.stdout.write('.'); } catch (e) { failed.push([name, e.message]); process.stdout.write('F'); } };
const assert = (c, m) => { if (!c) throw new Error(m || 'assertion failed'); };
const eq = (a, b, m) => { if (a !== b) throw new Error(`${m || 'not equal'}: expected ${JSON.stringify(b)}, got ${JSON.stringify(a)}`); };
const IDS = { solution: 'a1b2c3d4-0000-4000-8000-000000000001', feature: 'a1b2c3d4-0000-4000-8000-000000000002', component: 'a1b2c3d4-0000-4000-8000-000000000003', assets: 'a1b2c3d4-0000-4000-8000-000000000004' };
const proj = (key, build = 3) => ({ spec: JSON.parse(JSON.stringify(EXAMPLES.find((e) => e.key === key).spec)), ids: IDS, build, props: {} });

/** Run a bundle with stand-in SPFx modules; returns the module and the deps it asked for. */
function load(text) {
  let mod, asked, name;
  class Base { constructor() { this.properties = {}; } }
  const pp = new Proxy({}, { get: (_, k) => (k === 'PropertyPaneButtonType' ? { Normal: 0 } : (key, props) => ({ type: k, targetProperty: key, properties: props })) });
  const fakes = { '@microsoft/sp-core-library': { Version: { parse: (v) => v }, DisplayMode: { Read: 1, Edit: 2 } }, '@microsoft/sp-property-pane': pp, '@microsoft/sp-webpart-base': { BaseClientSideWebPart: Base }, '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } }, '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: 'x' } } };
  vm.runInNewContext(text, { define: (n, deps, f) => { name = n; asked = deps; mod = f(...deps.map((d) => fakes[d])); } });
  return { mod, asked, name, Base };
}

// ------------------------------------------------------------------ bundle
test('bundle: AMD module named <componentId>_<version>, same deps as the toolchain, default export extends BaseClientSideWebPart', () => {
  for (const ex of EXAMPLES) {
    const p = proj(ex.key);
    const { mod, asked, name, Base } = load(bundleText(p));
    eq(name, `${IDS.component}_1.0.3`, ex.key + ' define name');
    eq(JSON.stringify(asked), JSON.stringify(DEPS.map((d) => d[0])), ex.key + ' deps');
    assert(mod.__esModule && typeof mod.default === 'function', 'default export');
    assert(new mod.default() instanceof Base, 'extends the base class');
  }
});
test('bundle: matches the toolchain bundle for the same project (define name and dependency list)', () => {
  const tc = fs.readFileSync(path.join(here, 'fixtures/toolchain-announcements.js'), 'utf8');
  const head = /^define\("([^"]+)",(\[[^\]]*\])/.exec(tc);
  const ours = /^define\("([^"]+)",(\[[^\]]*\])/.exec(bundleText(proj('announcements')));
  eq(ours[1], head[1], 'define name'); eq(ours[2], head[2], 'deps');
});
test('bundle: property pane comes from the spec, with a Create button per list and typed defaults', () => {
  const p = proj('announcements');
  const { mod } = load(bundleText(p));
  const wp = new mod.default(); wp.properties = {};
  const cfg = wp.getPropertyPaneConfiguration();
  const groups = cfg.pages[0].groups.map((g) => g.groupName);
  eq(groups.join(','), 'Display,Data,Lists,About', 'groups');
  const fields = cfg.pages[0].groups[0].groupFields.map((f) => f.type);
  eq(fields.join(','), 'PropertyPaneTextField,PropertyPaneSlider,PropertyPaneToggle,PropertyPaneToggle,PropertyPaneTextField', 'field kinds');
  assert(cfg.pages[0].groups[2].groupFields.some((f) => f.type === 'PropertyPaneButton' && /Create "Announcements" list/.test(f.properties.text)), 'create-list button');
  const props = wp._props();
  eq(props.maxItems, 6, 'slider default is a number'); eq(props.showImages, true, 'toggle default is boolean'); eq(props.listName, 'Announcements', 'list default');
});
test('bundle: user code runs in its own block, so redeclaring h or props is allowed', () => {
  const p = proj('countdown'); p.spec.code = 'const h = ctx.h; let props = 1; return props;';
  assert(!checkCode(p.spec.code), 'checkCode accepts it');
  load(bundleText(p));
});
test('bundle: file name carries a 20-hex content hash; changes when the code changes', () => {
  const a = buildWebPart(proj('quicklinks')), p = proj('quicklinks'); p.spec.code += '\n// x';
  assert(/^quick-links-tiles-web-part_[0-9a-f]{20}\.js$/.test(a.file), a.file);
  assert(buildWebPart(p).file !== a.file, 'hash changes');
  eq(hash20('abc').length, 20); assert(/^[0-9a-f]+$/.test(hash20('x'.repeat(5000))), 'hex only');
});
test('runtime: no references outside its arguments (it is serialised into the bundle)', () => {
  assert(/^function spfxRuntime\(m, SPEC, CSS, APP\)/.test(RUNTIME_SOURCE), 'function source');
  assert(!/\bimport\b|\brequire\(|\beval\(|new Function/.test(RUNTIME_SOURCE), 'no imports or eval');
  assert(!/https?:\/\/(?!www\.w3\.org)[a-z]/.test(RUNTIME_SOURCE), 'no outside URLs');
});

// ------------------------------------------------------------------ manifest + package
test('manifest: same keys and loader config shape as the toolchain manifest', () => {
  const m = manifest(proj('announcements'), 'announcements-board-web-part_0123456789abcdef0123.js');
  eq(Object.keys(m).join(','), 'id,alias,componentType,version,manifestVersion,supportedHosts,supportsThemeVariants,supportsFullBleed,preconfiguredEntries,loaderConfig');
  eq(m.alias, 'AnnouncementsBoardWebPart'); eq(m.version, '1.0.3');
  eq(m.loaderConfig.entryModuleId, 'announcements-board-web-part');
  eq(m.loaderConfig.internalModuleBaseUrls[0], 'HTTPS://SPCLIENTSIDEASSETLIBRARY/');
  eq(m.loaderConfig.scriptResources['@microsoft/sp-webpart-base'].id, '974a7777-0990-4136-8fa6-95d80114c2e0');
  eq(m.preconfiguredEntries[0].groupId, '75e22ed5-fa14-4829-850a-c890608aca2d', 'Communication and collaboration');
  eq(m.preconfiguredEntries[0].properties.maxItems, 6); eq(m.preconfiguredEntries[0].properties.showImages, true);
});
test('package: same parts as `gulp package-solution --ship` (minus the unused icon PNGs)', () => {
  const p = proj('announcements');
  const files = packageFiles(p, { configIds: ['c1', 'c2'] }).files.map((f) => f[0].replace(/_[0-9a-f]{20}\.js$/, '_HASH.js')).sort();
  const expected = ['[Content_Types].xml', 'AppManifest.xml', 'ClientSideAssets.xml', 'ClientSideAssets.xml.config.xml', 'ClientSideAssets/announcements-board-web-part_HASH.js',
    `_rels/.rels`, `_rels/AppManifest.xml.rels`, `_rels/ClientSideAssets.xml.rels`, `_rels/feature_${IDS.feature}.xml.rels`, `feature_${IDS.feature}.xml`, `feature_${IDS.feature}.xml.config.xml`, `${IDS.feature}/WebPart_${IDS.component}.xml`].sort();
  eq(JSON.stringify(files), JSON.stringify(expected));
});
test('package: relationship ids, versions and the client-side asset point at the right files', () => {
  const p = proj('quicklinks', 7);
  const { files, webPart } = packageFiles(p, { configIds: ['c1', 'c2'] });
  const get = (n) => files.find((f) => f[0] === n)[1];
  assert(/Version="1\.0\.7\.0" SkipFeatureDeployment="true"/.test(get('AppManifest.xml')), 'solution version + tenant-wide');
  assert(/Id="r5"[\s\S]*Id="r6"/.test(get('_rels/AppManifest.xml.rels')), 'manifest rels r5, r6');
  assert(get('_rels/ClientSideAssets.xml.rels').includes(`/ClientSideAssets/${webPart.file}`), 'asset rel');
  const el = get(`${IDS.feature}/WebPart_${IDS.component}.xml`);
  const man = JSON.parse(/ComponentManifest="([^"]*)"/.exec(el)[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'));
  eq(man.loaderConfig.scriptResources['quick-links-tiles-web-part'].path, webPart.file);
  eq(man.version, '1.0.7');
  p.spec.tenantWide = false;
  assert(/SkipFeatureDeployment="false"/.test(packageFiles(p).files.find((f) => f[0] === 'AppManifest.xml')[1]), 'site-by-site install');
});
test('package: XML special characters in titles are escaped', () => {
  const p = proj('countdown'); p.spec.title = 'Q&A <Board> "x"'; p.spec.description = "It's & more";
  const files = packageFiles(p).files;
  const app = files.find((f) => f[0] === 'AppManifest.xml')[1];
  assert(app.includes('Q&amp;A &lt;Board&gt;') && !app.includes('<Board>'), 'escaped');
  const el = files.find((f) => f[0].includes('WebPart_'))[1];
  assert(!/ComponentManifest="[^"]*"[^ >]/.test(el.replace(/ComponentManifest="[^"]*" Type/, 'X')), 'attribute intact');
});

// ------------------------------------------------------------------ CSS scoping
test('scopeCss: prefixes selectors, keeps at-rules, maps :root/.root, renames keyframes, drops @import', () => {
  const out = scopeCss(`@import url(https://evil.example/x.css);
.a, .b:hover > i { color: red }
:root { --x: 1 } .root.is-edit .c { x: y }
@media (max-width: 500px) { .a { y: z } }
@container (max-width: 400px) { .b, .c { z: 1 } }
@keyframes spin { to { transform: rotate(1turn) } }
.s { animation: spin 1s linear infinite }
.n:not(.x, .y) { a: b }
/* comment { } */ body { margin: 0 }`, 'sss-1');
  assert(!out.includes('@import'), 'no import');
  assert(out.includes('.sss-1 .a,.sss-1 .b:hover > i{'), 'list'); assert(out.includes('.sss-1{ --x: 1 }'), ':root');
  assert(out.includes('.sss-1.is-edit .c{'), '.root'); assert(out.includes('@media (max-width: 500px){.sss-1 .a{'), 'media');
  assert(out.includes('@container (max-width: 400px){.sss-1 .b,.sss-1 .c{'), 'container');
  assert(out.includes('@keyframes sss-1-spin{') && out.includes('animation: sss-1-spin 1s'), 'keyframes');
  assert(out.includes('.sss-1 .n:not(.x, .y){'), 'commas inside :not() stay'); assert(out.includes('.sss-1{ margin: 0 }'), 'body maps to root');
});

// ------------------------------------------------------------------ spec cleaning + AI request
test('cleanSpec: repairs keys, types, defaults, lists and Title placement', () => {
  const { spec, notes } = cleanSpec({ title: 'X', alias: 'my part!', properties: [
    { key: 'List Name', type: 'list', default: 'Tasks' }, { key: 'listName', type: 'weird', default: 3 }, { key: 'size', type: 'slider', min: 5, max: 1, default: 'x' },
    { key: 'mode', type: 'dropdown', options: ['A', 'B'], default: 'C' }, { key: 'on', type: 'toggle', default: 'Yes' }, { key: 'sssHack', type: 'text' }],
    lists: [{ name: 'Tasks', columns: [{ name: 'Due date', type: 'DateTime' }, { name: 'Title', type: 'Number' }, { name: 'ID', type: 'Text' }, { name: 'Due date', type: 'Text' }], sample: [['2026-01-01', 'Task 1', '9']] }, { name: 'tasks', columns: [] }], css: 1, code: null });
  eq(spec.alias, 'MyPart'); eq(spec.properties.map((p) => p.key).join(','), 'listName,listName2,size,mode,on,xsssHack');
  eq(spec.properties[1].type, 'text'); eq(spec.properties[2].max, 15); eq(spec.properties[2].default, '5');
  eq(spec.properties[3].default, 'A'); eq(spec.properties[4].default, 'true');
  eq(spec.lists.length, 1, 'duplicate list dropped');
  eq(spec.lists[0].columns.map((c) => c.name + ':' + c.type).join(','), 'Title:Text,DueDate:DateTime,IDValue:Text');
  eq(JSON.stringify(spec.lists[0].sample[0]), JSON.stringify(['Task 1', '2026-01-01', '9']), 'sample follows Title to the front');
  assert(notes.some((n) => /weird/.test(n)) && notes.some((n) => /reserved/.test(n)), 'notes');
  eq(spec.group, 'Advanced'); eq(typeof spec.css, 'string');
});
test('checkCode: every example compiles; a syntax error is reported', () => {
  for (const ex of EXAMPLES) eq(checkCode(ex.spec.code), '', ex.key);
  assert(/Unexpected|expected/i.test(checkCode('const x = ;')), 'syntax error');
});
test('examples are already clean (cleanSpec changes nothing)', () => {
  for (const ex of EXAMPLES) { const { spec, notes } = cleanSpec(ex.spec); eq(notes.length, 0, ex.key + ' notes'); eq(JSON.stringify(spec.properties), JSON.stringify(ex.spec.properties.map((p) => ({ ...p, ...(p.type === 'slider' ? { min: p.min, max: p.max } : {}) }))), ex.key + ' props'); }
});
test('AI request: structured output schema, the example in the system prompt, icons listed, change history shape', () => {
  const r = buildRequest({ text: 'A directory', model: 'claude-sonnet-5-5' });
  eq(r.output_config.format.type, 'json_schema'); assert(r.output_config.format.schema === SPEC_SCHEMA, 'schema');
  assert(!('temperature' in r), 'no temperature');
  assert(SYSTEM_PROMPT.includes('"title": "Quick Links Tiles"'), 'example included');
  for (const n of ICON_NAMES) assert(RUNTIME_SOURCE.includes(` ${n}: '`) || RUNTIME_SOURCE.includes(`${n}: '`), 'runtime has icon ' + n);
  const h = changeHistory('A directory', EXAMPLES[0].spec, 'Add search');
  eq(h.map((m) => m.role).join(','), 'user,assistant,user'); assert(JSON.parse(h[1].content).code, 'current spec as the answer');
  const walk = (s) => { if (s.type === 'object') { eq(s.additionalProperties, false); eq(s.required.join(), Object.keys(s.properties).join()); Object.values(s.properties).forEach(walk); } if (s.type === 'array') walk(s.items); };
  walk(SPEC_SCHEMA);
});

// ------------------------------------------------------------------ source project
test('source project: SPFx 1.21.1 layout, same ids and version as the package, code embedded', () => {
  const p = proj('announcements', 4);
  const { files, folder } = sourceFiles(p);
  const get = (n) => files.find((f) => f[0] === n)?.[1];
  eq(folder, 'announcements-board');
  const pkg = JSON.parse(get('package.json')); eq(pkg.version, '1.0.4'); eq(pkg.dependencies['@microsoft/sp-webpart-base'], '1.21.1');
  const sol = JSON.parse(get('config/package-solution.json')).solution; eq(sol.id, IDS.solution); eq(sol.version, '1.0.4.0'); eq(sol.features[0].id, IDS.feature);
  const cfg = JSON.parse(get('config/config.json')); assert(cfg.bundles['announcements-board-web-part'], 'bundle entry');
  const man = JSON.parse(get('src/webparts/announcementsBoard/AnnouncementsBoardWebPart.manifest.json')); eq(man.id, IDS.component); eq(man.version, '*'); assert(!man.loaderConfig, 'no loaderConfig in source');
  const ts = get('src/webparts/announcementsBoard/AnnouncementsBoardWebPart.ts');
  assert(ts.startsWith('/* eslint-disable */\n// @ts-nocheck'), 'lint and type checks off'); assert(ts.includes("import * as wb from '@microsoft/sp-webpart-base';"), 'imports');
  assert(ts.includes("const list = props.listName;"), 'code embedded'); assert(ts.includes('export default R({ core, pp, wb, http, cb }, SPEC, CSS, APP);'), 'default export');
  eq(JSON.parse(get('tsconfig.json')).compilerOptions.target, 'es2017');
});

test('names: kebab and pascal helpers', () => {
  eq(kebab('QuickLinksTilesWebPart'), 'quick-links-tiles-web-part'); eq(pascal('staff directory 2'), 'StaffDirectory2'); eq(pascal('3d view'), 'DView');
  eq(names(proj('countdown')).scope, 'sss-a1b2c3d4');
});

console.log(`\n\n${passed} passed, ${failed.length} failed`);
for (const [n, m] of failed) console.log('✗ ' + n + '\n    ' + m);
process.exit(failed.length ? 1 : 0);
