// bundle.js — turns a project (spec + ids + version) into the files SPFx needs: the web part manifest and the AMD
// bundle. The bundle has the same shape the SPFx 1.21 toolchain produces (define("<id>_<version>", [deps],
// factory) whose default export is the web part class), so SharePoint's loader treats it like any other web part.
import { RUNTIME_SOURCE } from './runtime.js';

export const SPFX_VERSION = '1.21.1';
/** SharePoint framework components the bundle asks the loader for (ids are fixed across tenants). */
export const DEPS = [
  ['@microsoft/sp-core-library', '7263c7d0-1d6a-45ec-8d85-d4d1d234171b', 'core'],
  ['@microsoft/sp-property-pane', 'f9e737b7-f0df-4597-ba8c-3060f82380db', 'pp'],
  ['@microsoft/sp-webpart-base', '974a7777-0990-4136-8fa6-95d80114c2e0', 'wb'],
  ['@microsoft/sp-http', 'c07208f0-ea3b-4c1a-9965-ac1b825211a6', 'http'],
  ['@microsoft/sp-component-base', '467dc675-7cc5-4709-8aac-78e3b71bd2f6', 'cb'],
];
/** Toolbox groups SharePoint knows. The id decides where the web part shows in the + menu. */
export const GROUPS = {
  'Text, media, and content': 'cf066440-0614-43d6-98ae-0b31cf14c7c3',
  'Discovery': '1edbd9a8-0bfb-4aa2-9afd-14b8c45dd489',
  'Communication and collaboration': '75e22ed5-fa14-4829-850a-c890608aca2d',
  'Planning and process': '1bc7927e-4a5e-4520-b540-71305c79c20a',
  'Business and intelligence': '4aca9e90-eff5-4fa1-bac7-728f5f157b66',
  'Site tools': '070951d7-94da-4db8-b06e-9d581f1f55b1',
  'Advanced': '5c03119e-3074-46fd-976b-c60198311f70',
};
export const PROP_TYPES = ['text', 'multiline', 'number', 'toggle', 'dropdown', 'slider', 'list', 'color', 'url'];
export const COL_TYPES = ['Text', 'Note', 'RichText', 'Number', 'Currency', 'DateTime', 'DateTimeFull', 'Boolean', 'Choice', 'MultiChoice', 'URL', 'Image', 'User', 'UserMulti'];

export const kebab = (s) => String(s || '').replace(/([a-z0-9])([A-Z])/g, '$1-$2').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '').toLowerCase() || 'web-part';
export const pascal = (s) => (String(s || '').replace(/[^A-Za-z0-9]+(.)?/g, (_, c) => (c ? c.toUpperCase() : '')).replace(/^[^A-Za-z]+/, '') || 'StudioPart').replace(/^./, (c) => c.toUpperCase());
export const uuid = () => (globalThis.crypto?.randomUUID ? globalThis.crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => { const r = (Math.random() * 16) | 0; return (c === 'x' ? r : (r & 3) | 8).toString(16); }));
/** Stable 20-hex-char content hash (FNV-1a, two lanes) for file names, like webpack's chunk hashes. */
export function hash20(s) {
  let a = 0x811c9dc5, b = 0x01000193 ^ 0x5bd1e995;
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); a = Math.imul(a ^ c, 16777619) >>> 0; b = Math.imul(b ^ c ^ (i & 255), 2246822519) >>> 0; }
  const x = (a.toString(16).padStart(8, '0') + b.toString(16).padStart(8, '0'));
  return (x + (Math.imul(a ^ b, 3266489917) >>> 0).toString(16).padStart(8, "0")).slice(0, 20);
}

// ------------------------------------------------------------------ CSS scoping
/** Prefix every selector with the web part's scope class so generated CSS never leaks onto the page.
 *  `:root`, `.root` and `:host` at the start of a selector mean the web part's outer element. Keyframes are renamed. */
export function scopeCss(css, scope) {
  const src = String(css || '').replace(/\/\*[\s\S]*?\*\//g, '');
  const frames = new Map();
  src.replace(/@(?:-webkit-)?keyframes\s+([A-Za-z_][\w-]*)/g, (_, n) => { frames.set(n, `${scope}-${n}`); return _; });
  const sel = (s) => splitTop(s, ',').map((x) => x.trim()).filter(Boolean).map((x) => {
    const m = /^(:root|\.root|:host|&)(?![\w-])/.exec(x);
    if (m) return `.${scope}${x.slice(m[0].length)}`;
    if (/^(html|body)\b/.test(x)) return `.${scope}`;
    return `.${scope} ${x}`;
  }).join(',');
  const renameFrames = (decl) => frames.size ? decl.replace(/(animation(?:-name)?\s*:)([^;}]*)/g, (_, p, v) => p + v.replace(/[A-Za-z_][\w-]*/g, (w) => frames.get(w) || w)) : decl;
  function walk(s) {
    let out = '', i = 0;
    while (i < s.length) {
      const open = s.indexOf('{', i);
      if (open < 0) break;
      let prelude = s.slice(i, open);
      prelude = prelude.slice(prelude.lastIndexOf(';') + 1).trim(); // drops @import/@charset statements before the rule
      let depth = 1, j = open + 1;
      while (j < s.length && depth) { if (s[j] === '{') depth++; else if (s[j] === '}') depth--; j++; }
      const inner = s.slice(open + 1, j - 1);
      if (/^@(media|container|supports|layer)\b/i.test(prelude)) out += `${prelude}{${walk(inner)}}`;
      else if (/^@(-webkit-)?keyframes\s+/i.test(prelude)) { const n = prelude.split(/\s+/)[1]; out += `${prelude.split(/\s+/)[0]} ${frames.get(n) || n}{${inner}}`; }
      else if (/^@(font-face|page|property)\b/i.test(prelude)) out += `${prelude}{${inner}}`;
      else if (prelude.startsWith('@')) { /* @import and friends are dropped: no outside requests */ }
      else if (prelude) out += `${sel(prelude)}{${renameFrames(inner)}}`;
      i = j;
    }
    return out;
  }
  return walk(src);
}
function splitTop(s, ch) {
  const out = []; let depth = 0, cur = '';
  for (const c of s) { if (c === '(' || c === '[') depth++; else if (c === ')' || c === ']') depth--; if (c === ch && !depth) { out.push(cur); cur = ''; } else cur += c; }
  out.push(cur); return out;
}

// ------------------------------------------------------------------ versions and names
/** 1.0.N.0 for the solution, 1.0.N for the component (SharePoint compares them on upgrade). */
export const solutionVersion = (project) => `1.0.${project.build || 1}.0`;
export const componentVersion = (project) => `1.0.${project.build || 1}`;
export function names(project) {
  const alias = pascal(project.spec.alias || project.spec.title) + 'WebPart';
  const entry = kebab(alias);
  const pkg = kebab(project.spec.alias || project.spec.title);
  return { alias, entry, pkg, solution: `${pkg}-client-side-solution`, scope: 'sss-' + project.ids.component.slice(0, 8) };
}

/** The spec as the runtime sees it (only what it needs, nothing the visitor did not write). */
export function runtimeSpec(project) {
  const s = project.spec, n = names(project);
  return {
    id: project.ids.component, title: s.title, description: s.description, version: componentVersion(project), scope: n.scope,
    properties: (s.properties || []).map((p) => ({ key: p.key, type: p.type, label: p.label, description: p.description || '', group: p.group || '', default: p.default ?? '', options: p.options || [], min: p.min, max: p.max, step: p.step })),
    lists: (s.lists || []).map((l) => ({ name: l.name, description: l.description || '', columns: l.columns || [], sample: l.sample || [] })),
  };
}

/** The AMD module text. `code` is the body of an async function receiving ctx. */
export function bundleText(project) {
  const deps = DEPS.map((d) => d[0]);
  const n = names(project);
  const css = scopeCss(project.spec.css, n.scope);
  const code = String(project.spec.code || '');
  return `define(${JSON.stringify(project.ids.component + '_' + componentVersion(project))},${JSON.stringify(deps)},function(${DEPS.map((d) => d[2]).join(',')}){"use strict";
/* ${String(project.spec.title).replace(/\*\//g, '* /')} — generated by SPFx Studio (satechitsolutions.com/spfx-studio). Runtime + your web part code below. */
var R=${RUNTIME_SOURCE};
var SPEC=${JSON.stringify(runtimeSpec(project))};
var CSS=${JSON.stringify(css)};
var APP=async function(ctx){
var h=ctx.h,icon=ctx.icon,esc=ctx.esc,fmt=ctx.fmt,props=ctx.props,sp=ctx.sp;
{
${code}
}
};
var C=R({${DEPS.map((d) => `${d[2]}:${d[2]}`).join(',')}},SPEC,CSS,APP);
return{__esModule:true,default:C};});
`;
}

export function manifest(project, bundleFile) {
  const s = project.spec, n = names(project);
  const props = {};
  for (const p of s.properties || []) {
    if (p.default === undefined || p.default === '') continue;
    props[p.key] = p.type === 'toggle' ? /^(true|yes|1)$/i.test(String(p.default)) : (p.type === 'number' || p.type === 'slider') ? Number(p.default) : p.default;
  }
  const scriptResources = { [n.entry]: { type: 'path', path: bundleFile } };
  for (const [name, id] of DEPS) scriptResources[name] = { type: 'component', id, version: SPFX_VERSION };
  const groupName = GROUPS[s.group] ? s.group : 'Advanced';
  return {
    id: project.ids.component, alias: n.alias, componentType: 'WebPart', version: componentVersion(project), manifestVersion: 2,
    supportedHosts: ['SharePointWebPart', 'SharePointFullPage'], supportsThemeVariants: true, supportsFullBleed: !!s.fullBleed,
    preconfiguredEntries: [{
      groupId: GROUPS[groupName], group: { default: groupName },
      title: { default: s.title }, description: { default: s.description || s.title },
      officeFabricIconFontName: s.icon || 'Page', properties: props,
    }],
    loaderConfig: { internalModuleBaseUrls: ['HTTPS://SPCLIENTSIDEASSETLIBRARY/'], entryModuleId: n.entry, scriptResources },
  };
}

/** Bundle file name, bundle text and manifest together. */
export function buildWebPart(project) {
  const text = bundleText(project);
  const file = `${names(project).entry}_${hash20(text)}.js`;
  return { file, text, manifest: manifest(project, file) };
}
