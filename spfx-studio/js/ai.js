// ai.js — plain-English description → Claude → a web part spec (properties, lists with sample rows, CSS, code).
// The browser calls the Claude API directly with the visitor's own key (the same key Power App Studio uses),
// with structured outputs so the answer always parses. cleanSpec() then repairs anything out of bounds and
// checkCode() catches syntax errors before the preview runs, so one automatic repair round can fix them.
import { GROUPS, PROP_TYPES, COL_TYPES, pascal } from './bundle.js';
import { EXAMPLES } from './examples.js';

export const API_URL = 'https://api.anthropic.com/v1/messages';
export const API_VERSION = '2026-10-21';
export const MODELS = [
  ['claude-sonnet-5-5', 'Claude Sonnet 5.5 (recommended)'],
  ['claude-opus-5-5', 'Claude Opus 5.5 (most careful, slower)'],
  ['claude-haiku-5-5', 'Claude Haiku 5.5 (fastest, cheapest)'],
];
export const ICON_NAMES = ['add', 'close', 'check', 'chevronRight', 'chevronLeft', 'chevronDown', 'chevronUp', 'search', 'edit', 'trash', 'link', 'external', 'calendar', 'clock', 'person', 'people', 'mail', 'phone', 'location', 'star', 'heart', 'info', 'warning', 'megaphone', 'document', 'folder', 'image', 'video', 'home', 'settings', 'filter', 'sort', 'refresh', 'download', 'upload', 'grid', 'list', 'bell', 'bookmark', 'chart', 'trophy', 'lightbulb', 'globe', 'lock', 'tag', 'flag', 'briefcase', 'book', 'rocket', 'sparkle', 'more', 'play', 'pause', 'arrowRight', 'arrowLeft'];

const obj = (properties) => ({ type: 'object', additionalProperties: false, required: Object.keys(properties), properties });
const str = (description) => ({ type: 'string', ...(description ? { description } : {}) });
const arr = (items) => ({ type: 'array', items });
export const SPEC_SCHEMA = obj({
  title: str('Name shown in the SharePoint toolbox, 2-4 words'),
  alias: str('PascalCase identifier, e.g. TeamDirectory'),
  description: str('One sentence for the toolbox'),
  icon: str('Fluent UI MDL2 icon name for the toolbox, e.g. People, Link, Calendar, Megaphone'),
  group: { type: 'string', enum: Object.keys(GROUPS) },
  fullBleed: { type: 'boolean' },
  properties: arr(obj({ key: str(), type: { type: 'string', enum: PROP_TYPES }, label: str(), description: str(), group: str(), default: str(), options: arr(str()), min: { type: 'number' }, max: { type: 'number' } })),
  lists: arr(obj({ name: str(), description: str(), columns: arr(obj({ name: str(), type: { type: 'string', enum: COL_TYPES }, required: { type: 'boolean' }, choices: arr(str()) })), sample: arr(arr(str())) })),
  css: str(),
  code: str(),
  notes: str(),
});

const example = EXAMPLES[0].spec;
export const SYSTEM_PROMPT = `You build SharePoint Online web parts. Your output is a JSON spec that SPFx Studio packages into a real SPFx 1.21 web part (.sppkg) with no build step. It will run on corporate intranets, often US government tenants (GCC High), so nothing may load from outside SharePoint.

## How your code runs
\`code\` is the body of an async function, run on every render (first load and each property-pane change). Build the UI into ctx.root (an empty div). Await your data before drawing so the web part never flashes half-built content; the runtime shows a spinner after 200 ms and a friendly error box if you throw.
In scope: ctx, h, icon, esc, fmt, props, sp.
- h(tag, attrs, ...children) returns an element. attrs: class, style (object; CSS variables like '--cols' allowed), text, html (trusted content only), dataset, on<event> handlers (onclick, oninput, onkeydown, onsubmit…), aria-* and any other attribute; value/checked/disabled set as properties. Children: nodes, strings, numbers, arrays; null/false are skipped. SVG tags work.
- icon(name, size = 18, extraClass) returns an inline SVG. Names: ${ICON_NAMES.join(', ')}. Use only these.
- esc(s); ctx.clear(el) empties an element.
- fmt.date(d), fmt.dateTime(d), fmt.time(d), fmt.relative(d) ("in 3 days"), fmt.number(n, opts), fmt.initials(name), fmt.truncate(s, n), fmt.plain(html).
- props: property values, typed (number/slider are numbers, toggle is boolean, the rest strings).
- ctx.state: an object that survives re-renders (current tab, filter, page number…).
- ctx.cache(key, fn): memoizes a promise until a property changes or ctx.refresh() runs. Wrap every data load: await ctx.cache('items:' + list, () => sp.items(list, {...})).
- ctx.refresh() clears the cache and re-renders (after add/update/delete). ctx.rerender() re-renders keeping the cache.
- ctx.isCurrent() is false once a newer render started: check it in timers and late callbacks and stop.
- ctx.page {webUrl, webTitle, serverRelativeUrl, editMode, instanceId}; ctx.user {displayName, email, loginName}; ctx.theme (colours).
- ctx.loading(text) returns a spinner element; ctx.showError(e) returns an error element.
- sp is SharePoint REST on the current site (sp.at(url) for another site):
  - sp.items(list, {select, expand, filter, orderBy, top = 500, all}) returns an array. OData $filter syntax (eq, ne, gt, lt, and, or, substringof('x',Field), datetime'2026-01-01T00:00:00Z'). Always pass select. Person columns need expand: ['Owner'] with select 'Owner/Title', 'Owner/EMail' (selecting 'Owner' alone fails). Author/Editor work the same way.
  - sp.item(list, id, {select, expand}); sp.add(list, data) returns the new item; sp.update(list, id, data); sp.remove(list, id).
  - Writing values: Text/Choice/Note strings; Number numbers; Boolean true/false; DateTime ISO strings; URL/Image {Url, Description}; Person: <Column>Id = a user id (from sp.me() or sp.ensureUser(email or login)); MultiChoice arrays.
  - Reading values: URL/Image are {Url, Description}; DateTime ISO strings; Boolean true/false; MultiChoice arrays; Note is plain text; RichText is HTML (fmt.plain it unless you mean to render it).
  - sp.me() returns {Id, Title, Email}; sp.lists() returns [{Title}]; sp.fields(list); sp.listExists(list).
  - Failures throw an Error with .status (404 when the list does not exist).

## Spec rules
- properties are what page editors configure. key is camelCase and unique. type: text | multiline | number | toggle | dropdown (options) | slider (min, max) | list (picks a list on the site) | color (hex, '' means follow the theme) | url. default is always a string ('true'/'false' for toggles, digits for numbers). group names the property-pane group ('Display', 'Data', 'Content'…). Every list you read gets a list-type property whose default is that list's name; use props.<key> as the list name in code.
- lists: every list your code reads or writes. Column names are PascalCase with no spaces (so internal and display names match). The first column is Title (Text). Column types: ${COL_TYPES.join(', ')}. choices only for Choice and MultiChoice. sample: 4-8 realistic rows, one string per column in column order: DateTime as '+N'/'-N' days from today or ISO; User as a full name; Boolean 'true'/'false'; URL and Image as URLs (leave Image empty unless you have a real https image URL on SharePoint); MultiChoice values separated by ';'. The preview runs on these rows. The property pane gets a "Create list" button for each list automatically.
- css is plain CSS, scoped automatically (every selector is prefixed with the web part root; ':root' means the root itself). No @import and no url() to outside hosts. Follow the site theme and section backgrounds with these variables: var(--sss-primary), --sss-primary-dark, --sss-primary-light, --sss-primary-lighter, --sss-text, --sss-muted, --sss-bg, --sss-surface, --sss-border, --sss-link. Text on a var(--sss-primary) background uses color: var(--sss-bg), never #fff (in strong, inverted sections the primary turns white). The root is a size container: use @container (max-width: 480px) {...} for responsive layout, because web parts sit in one-third columns as well as full width. Prefix your classes with 2-3 letters unique to this web part.
- code is modern JavaScript (ES2020, async/await). No import, require, eval, new Function, document.write, fetch to other hosts, CDNs, web fonts or libraries. Touch nothing outside ctx.root (no document.body, no global styles, no window properties). Attach events through h attributes. Stop timers when ctx.isCurrent() is false.
- fullBleed true only for banners or heroes meant for full-width sections.

## Quality bar
Production work a SharePoint designer would be proud of: clear hierarchy, comfortable spacing, rounded cards with subtle borders, hover and focus states, restrained transitions. Fluent look: inherit Segoe UI, 14px body, 20px section headings, 600 weights, the theme colours.
Handle every state: loading (await first), empty (friendly message and what to do), list not configured or missing (in edit mode say how to fix it through the property pane's Lists section; in view mode a short neutral message), and errors (catch status 404 and explain; rethrow anything else).
Accessible: semantic elements, real buttons for actions, labels for inputs, aria-live on regions that change, visible focus, good contrast, keyboard support.
Responsive from 300px to 1400px wide.
Forms: use a <form> with onsubmit that calls e.preventDefault() first, validate, disable the button while saving, show success or failure inline, then ctx.refresh().
Add search, filters, sorting or paging where they clearly help. Don't invent big features nobody asked for; when the request is ambiguous pick the most useful reading and say so in notes.
notes: 2-4 short sentences for the page owner: how to set it up and any choice you made.

## Example of the expected standard
${JSON.stringify({ ...example, fullBleed: false }, null, 1)}`;

export function buildRequest({ text, model = MODELS[0][0], history }) {
  const messages = history && history.length ? history : [{ role: 'user', content: `Web part request:\n\n${String(text).trim()}` }];
  return { model, max_tokens: 32000, system: SYSTEM_PROMPT, messages, output_config: { format: { type: 'json_schema', schema: SPEC_SCHEMA } } };
}
/** Conversation for a change: the original request, the current spec as Claude's answer, then the change. */
export function changeHistory(request, spec, change) {
  return [
    { role: 'user', content: `Web part request:\n\n${String(request || spec.description || spec.title).trim()}` },
    { role: 'assistant', content: JSON.stringify(specForClaude(spec)) },
    { role: 'user', content: `${String(change).trim()}\n\nReturn the complete updated spec. Keep everything that already works unless the change says otherwise.` },
  ];
}
export function fixHistory(request, spec, error) {
  return changeHistory(request, spec, `The preview reported this error:\n\n${error}\n\nFind the cause and fix it.`);
}
export const specForClaude = (s) => ({ title: s.title, alias: s.alias, description: s.description, icon: s.icon, group: s.group, fullBleed: !!s.fullBleed, properties: s.properties, lists: s.lists, css: s.css, code: s.code, notes: s.notes || '' });

export class AiError extends Error {}
export async function callClaude({ key, request, signal, fetchFn = globalThis.fetch }) {
  if (!/^sk-ant-/.test(String(key || '').trim())) throw new AiError('Enter an Anthropic API key (it starts with sk-ant-).');
  let res;
  try {
    res = await fetchFn(API_URL, {
      method: 'POST', signal,
      headers: { 'content-type': 'application/json', 'x-api-key': key.trim(), 'anthropic-version': API_VERSION, 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify(request),
    });
  } catch (e) {
    if (e?.name === 'AbortError') throw new AiError('Cancelled.');
    throw new AiError('Could not reach api.anthropic.com. Your network may block it (common on government networks), or you are offline.');
  }
  let body = null;
  try { body = await res.json(); } catch { /* not JSON */ }
  if (!res.ok) {
    const msg = body?.error?.message || res.statusText || 'request failed';
    if (res.status === 401) throw new AiError('The API key was rejected. Check it in the Claude Console.');
    if (res.status === 429) throw new AiError('Rate limited by the API. Wait a minute and try again.');
    if (res.status === 529 || res.status >= 500) throw new AiError('The Claude API is busy. Try again shortly.');
    throw new AiError(`Claude API error ${res.status}: ${msg}`);
  }
  if (body?.stop_reason === 'refusal') throw new AiError('Claude declined this request.');
  if (body?.stop_reason === 'max_tokens') throw new AiError('The web part was too long and got cut off. Ask for fewer features at once.');
  const text = (body?.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
  let spec;
  try { spec = JSON.parse(text); } catch { throw new AiError('The response was not a valid spec. Try again.'); }
  return { spec, usage: body?.usage || null };
}

// ------------------------------------------------------------------ cleaning
const camel = (s) => { const p = pascal(s); return p[0].toLowerCase() + p.slice(1); };
/** Make a spec safe to build. Returns {spec, notes}: notes say what was changed. */
export function cleanSpec(input) {
  const s = JSON.parse(JSON.stringify(input || {})), notes = [];
  s.title = String(s.title || 'My Web Part').trim().slice(0, 60);
  s.alias = pascal(s.alias || s.title).slice(0, 50);
  s.description = String(s.description || s.title);
  s.icon = /^[A-Za-z0-9]+$/.test(s.icon || '') ? s.icon : 'Page';
  if (!GROUPS[s.group]) s.group = 'Advanced';
  s.fullBleed = !!s.fullBleed;
  const seen = new Set();
  s.properties = (Array.isArray(s.properties) ? s.properties : []).filter((p) => p && p.key).map((p) => {
    let key = camel(p.key); if (/^sss/.test(key)) key = 'x' + key;
    while (seen.has(key)) key += '2';
    seen.add(key);
    const type = PROP_TYPES.includes(p.type) ? p.type : 'text';
    if (type !== p.type) notes.push(`Property ${key}: type "${p.type}" isn't supported, used text.`);
    const out = { key, type, label: String(p.label || key), description: String(p.description || ''), group: String(p.group || 'Settings'), default: p.default === undefined || p.default === null ? '' : String(p.default) };
    if (type === 'dropdown') { out.options = (p.options || []).map(String).filter(Boolean); if (!out.options.length) out.options = ['Default']; if (!out.options.includes(out.default)) out.default = out.options[0]; }
    if (type === 'slider') { out.min = Number.isFinite(p.min) ? p.min : 0; out.max = Number.isFinite(p.max) && p.max > out.min ? p.max : out.min + 10; if (isNaN(Number(out.default))) out.default = String(out.min); }
    if (type === 'number' && out.default !== '' && isNaN(Number(out.default))) out.default = '';
    if (type === 'toggle') out.default = /^(true|yes|1|on)$/i.test(out.default) ? 'true' : 'false';
    if (type === 'color' && out.default && !/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(out.default)) out.default = '';
    return out;
  });
  const listNames = new Set();
  s.lists = (Array.isArray(s.lists) ? s.lists : []).filter((l) => l && l.name).map((l) => {
    const name = String(l.name).trim().slice(0, 60);
    if (listNames.has(name.toLowerCase())) return null;
    listNames.add(name.toLowerCase());
    const cn = new Set();
    let columns = (l.columns || []).filter((c) => c && c.name).map((c) => {
      let n = c.name === 'Title' ? 'Title' : pascal(c.name);
      if (/^(ID|Id|Created|Modified|Author|Editor|ContentType|Attachments)$/.test(n)) { notes.push(`List ${name}: renamed the reserved column ${n} to ${n}Value.`); n += 'Value'; }
      if (cn.has(n)) return null; cn.add(n);
      const type = COL_TYPES.includes(c.type) ? c.type : 'Text';
      return { name: n, type: n === 'Title' ? 'Text' : type, required: !!c.required, choices: ['Choice', 'MultiChoice'].includes(type) ? (c.choices || []).map(String).filter(Boolean) : [] };
    }).filter(Boolean);
    const titleIdx = columns.findIndex((c) => c.name === 'Title');
    let sample = (l.sample || []).filter(Array.isArray).map((r) => r.map((v) => (v === null || v === undefined ? '' : String(v))));
    if (titleIdx < 0) { columns = [{ name: 'Title', type: 'Text', required: true, choices: [] }, ...columns]; sample = sample.map((r) => [r[0] || '', ...r]); }
    else if (titleIdx > 0) { const [t] = columns.splice(titleIdx, 1); columns.unshift(t); sample = sample.map((r) => { const v = r.splice(titleIdx, 1)[0]; return [v, ...r]; }); }
    sample = sample.slice(0, 25).map((r) => columns.map((_, i) => r[i] ?? ''));
    return { name, description: String(l.description || ''), columns, sample };
  }).filter(Boolean);
  s.css = String(s.css || '');
  s.code = String(s.code || '');
  s.notes = String(s.notes || '');
  if (/\b(eval|new\s+Function|document\.write)\s*\(/.test(s.code)) notes.push('The code uses eval, new Function or document.write. Those may be blocked on SharePoint pages.');
  if (/https?:\/\/(?!contoso\.)[a-z0-9.-]+\.(com|net|org|io|dev)\/[^\s'"]*\.(js|css)\b/i.test(s.code + s.css)) notes.push('The code references a script or stylesheet on another site. GCC High pages usually block those.');
  return { spec: s, notes };
}
/** Syntax check the code the way the bundle wraps it. Returns '' or the error message. */
export function checkCode(code) {
  try {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    // eslint-disable-next-line no-new
    new AsyncFunction('ctx', `var h=ctx.h,icon=ctx.icon,esc=ctx.esc,fmt=ctx.fmt,props=ctx.props,sp=ctx.sp;\n{\n${code}\n}`);
    return '';
  } catch (e) { return e.message; }
}
