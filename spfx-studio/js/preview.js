// preview.js — the page that runs inside the sandboxed preview iframe (sandbox="allow-scripts", no same-origin, so
// generated code can't read the API key or anything else from the studio). It stands in for SharePoint: an AMD
// define(), the five SPFx modules the bundle asks for, a page context, a theme service, and an in-memory SharePoint
// REST emulator seeded with the spec's sample rows. The bundle it runs is byte-for-byte the one that goes into the
// .sppkg. Talks to the studio with postMessage only.

/* eslint-disable */
export function previewHost() {
  'use strict';
  var parentWin = window.parent;
  function post(msg) { try { parentWin.postMessage(Object.assign({ __sss: 1 }, msg), '*'); } catch (e) { /* ignore */ } }
  var WEB = { absoluteUrl: 'https://contoso.sharepoint.com/sites/intranet', title: 'Contoso Intranet', serverRelativeUrl: '/sites/intranet' };
  var PEOPLE = [
    { Id: 7, Title: 'Avery Morgan', EMail: 'avery.morgan@contoso.com', LoginName: 'i:0#.f|membership|avery.morgan@contoso.com' },
    { Id: 11, Title: 'Jordan Lee', EMail: 'jordan.lee@contoso.com' }, { Id: 12, Title: 'Priya Natarajan', EMail: 'priya.natarajan@contoso.com' },
    { Id: 13, Title: 'Marcus Bell', EMail: 'marcus.bell@contoso.com' }, { Id: 14, Title: 'Sofia Ramirez', EMail: 'sofia.ramirez@contoso.com' },
    { Id: 15, Title: 'Daniel Okafor', EMail: 'daniel.okafor@contoso.com' }, { Id: 16, Title: 'Hannah Kim', EMail: 'hannah.kim@contoso.com' }
  ];
  var ME = PEOPLE[0];

  // ---------------------------------------------------------------- AMD + SPFx stand-ins
  var modules = {}, defined = null;
  window.define = function (name, deps, factory) {
    var args = deps.map(function (d) { if (!modules[d]) throw new Error('Unknown module ' + d); return modules[d]; });
    defined = factory.apply(null, args);
  };
  window.define.amd = true;
  var DisplayMode = { Read: 1, Edit: 2 };
  modules['@microsoft/sp-core-library'] = { Version: { parse: function (v) { return { toString: function () { return v; } }; } }, DisplayMode: DisplayMode, Environment: { type: 0 }, EnvironmentType: { Local: 0, SharePoint: 3 } };
  var FT = { Button: 11, CheckBox: 2, Dropdown: 6, Label: 7, Slider: 8, TextField: 3, Toggle: 5, ChoiceGroup: 10, Link: 4, HorizontalRule: 12 };
  function field(type) { return function (key, props) { return { type: type, targetProperty: key, properties: props || {} }; }; }
  modules['@microsoft/sp-property-pane'] = {
    PropertyPaneTextField: field(FT.TextField), PropertyPaneToggle: field(FT.Toggle), PropertyPaneDropdown: field(FT.Dropdown), PropertyPaneSlider: field(FT.Slider),
    PropertyPaneButton: field(FT.Button), PropertyPaneLabel: field(FT.Label), PropertyPaneCheckbox: field(FT.CheckBox), PropertyPaneChoiceGroup: field(FT.ChoiceGroup),
    PropertyPaneLink: field(FT.Link), PropertyPaneHorizontalRule: function () { return { type: FT.HorizontalRule, properties: {} }; },
    PropertyPaneFieldType: FT, PropertyPaneButtonType: { Normal: 0, Primary: 1 }
  };
  class BaseClientSideWebPart {
    constructor() { this.properties = {}; this.displayMode = DisplayMode.Read; this.instanceId = 'preview-' + Math.random().toString(16).slice(2, 8); }
    onInit() { return Promise.resolve(); }
    render() {}
    onDispose() {}
    onPropertyPaneConfigurationStart() {}
    onPropertyPaneFieldChanged() {}
    getPropertyPaneConfiguration() { return { pages: [] }; }
  }
  modules['@microsoft/sp-webpart-base'] = { BaseClientSideWebPart: BaseClientSideWebPart };
  modules['@microsoft/sp-http'] = { SPHttpClient: { configurations: { v1: { v: 1 } } } };
  var themeHandlers = [], currentTheme = null;
  modules['@microsoft/sp-component-base'] = { ThemeProvider: { serviceKey: 'ThemeProvider' } };
  var themeProvider = { tryGetTheme: function () { return currentTheme; }, themeChangedEvent: { add: function (o, fn) { themeHandlers.push(fn.bind(o)); } } };

  // ---------------------------------------------------------------- REST emulator
  var DB = {}, log = [];
  var BUILTIN_FIELDS = [['Title', 'Text'], ['Created', 'DateTime'], ['Modified', 'DateTime'], ['Author', 'User'], ['Editor', 'User']];
  function iname(n) { return String(n).replace(/[^A-Za-z0-9_]/g, ''); }
  function person(name) {
    if (!name) return null;
    var p = PEOPLE.filter(function (x) { return x.Title.toLowerCase() === String(name).toLowerCase() || x.EMail === name; })[0];
    if (!p) { p = { Id: 100 + PEOPLE.length, Title: String(name), EMail: String(name).toLowerCase().replace(/[^a-z]+/g, '.') + '@contoso.com' }; PEOPLE.push(p); }
    return { Id: p.Id, Title: p.Title, EMail: p.EMail, Name: 'i:0#.f|membership|' + p.EMail };
  }
  function newList(title, columns, template) {
    var l = { Id: rid(), Title: title, BaseTemplate: template || 100, fields: {}, items: [], nextId: 1, Hidden: false };
    BUILTIN_FIELDS.forEach(function (f) { l.fields[f[0]] = { InternalName: f[0], Title: f[0], TypeAsString: f[1], builtin: true, Required: f[0] === 'Title' }; });
    (columns || []).forEach(function (c) { addField(l, c); });
    DB[title.toLowerCase()] = l; return l;
  }
  function addField(l, c) {
    var n = c.name === 'Title' ? 'Title' : iname(c.name);
    var t = { Note: 'Note', RichText: 'Note', DateTimeFull: 'DateTime', Image: 'URL', Currency: 'Currency' }[c.type] || c.type || 'Text';
    l.fields[n] = { InternalName: n, Title: c.name, TypeAsString: t, studioType: c.type, Required: !!c.required, Choices: c.choices || [] };
  }
  function rid() { return 'xxxxxxxx-xxxx-4xxx-8xxx-xxxxxxxxxxxx'.replace(/x/g, function () { return (Math.random() * 16 | 0).toString(16); }); }
  function coerce(f, v, fromSample) {
    if (v === undefined || v === null || v === '') return null;
    switch (f.TypeAsString) {
      case 'Number': case 'Currency': return Number(v);
      case 'Boolean': return typeof v === 'boolean' ? v : /^(true|yes|1)$/i.test(String(v));
      case 'DateTime': { var d; if (fromSample && /^\s*[+-]\d+\s*$/.test(String(v))) { d = new Date(); d.setHours(9, 0, 0, 0); d.setDate(d.getDate() + Number(v)); } else d = new Date(v); if (isNaN(d)) return null; return d.toISOString().replace(/\.\d{3}Z$/, 'Z'); }
      case 'URL': return typeof v === 'object' ? v : { Url: String(v), Description: f.studioType === 'Image' ? '' : String(v) };
      case 'MultiChoice': return Array.isArray(v) ? v : String(v).split(/\s*;\s*/).filter(Boolean);
      case 'User': return typeof v === 'object' ? v : person(v);
      case 'UserMulti': return Array.isArray(v) ? v : String(v).split(/\s*;\s*/).filter(Boolean).map(person);
      default: return String(v);
    }
  }
  function seed(lists, keep) {
    if (!keep) { DB = {}; newList('Documents', [], 101); newList('Site Pages', [], 119); }
    (lists || []).forEach(function (def) {
      var l = DB[def.name.toLowerCase()];
      if (l && keep) { (def.columns || []).forEach(function (c) { var n = c.name === 'Title' ? 'Title' : iname(c.name); if (!l.fields[n]) addField(l, c); }); return; }
      l = newList(def.name, def.columns);
      (def.sample || []).forEach(function (row, ri) {
        var it = newItem(l, ME);
        (def.columns || []).forEach(function (c, i) { var n = c.name === 'Title' ? 'Title' : iname(c.name); it[n] = coerce(l.fields[n], row[i], true); });
        var d = new Date(Date.now() - (def.sample.length - ri) * 86400000 * 2); it.Created = d.toISOString().replace(/\.\d{3}Z$/, 'Z'); it.Modified = it.Created;
      });
    });
  }
  function newItem(l, who) {
    var now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
    var it = { Id: l.nextId++, Title: null, Created: now, Modified: now, Author: person(who.Title), Editor: person(who.Title) };
    l.items.push(it); return it;
  }
  function ser(l, it, select, expand) {
    var out = {}, exp = {}, sel = select && select.length ? select : null;
    (expand || []).forEach(function (e) { exp[e.toLowerCase()] = 1; });
    var userFields = Object.keys(l.fields).filter(function (k) { return /^User/.test(l.fields[k].TypeAsString); });
    if (sel) {
      sel.forEach(function (s) {
        var parts = s.split('/'), k = parts[0];
        if (k === '*') { Object.assign(out, ser(l, it, null, expand)); return; }
        if (k === 'Id' || k === 'ID') { out.Id = it.Id; out.ID = it.Id; return; }
        var f = l.fields[k] || (/Id$/.test(k) && l.fields[k.slice(0, -2)]);
        if (!f) throw httpErr(400, "The field or property '" + k + "' does not exist.");
        if (parts.length > 1) {
          if (!exp[k.toLowerCase()]) throw httpErr(400, "The query to field '" + s + "' is not valid. The $select query string must specify the target fields and the $expand query string must contains " + k + '.');
          var v = it[k]; out[k] = out[k] || (Array.isArray(v) ? v.map(function () { return {}; }) : v ? {} : null);
          if (Array.isArray(v)) v.forEach(function (p, i) { out[k][i][parts[1]] = p[parts[1]]; }); else if (v) out[k][parts[1]] = v[parts[1]];
          return;
        }
        if (/^User/.test(f.TypeAsString) && l.fields[k]) { if (!exp[k.toLowerCase()]) throw httpErr(400, "The field '" + k + "' is a lookup. Select " + k + 'Id, or expand it and select ' + k + '/Title.'); out[k] = it[k]; return; }
        if (f !== l.fields[k]) { out[k] = idOf(it[k.slice(0, -2)]); return; }
        out[k] = it[k] === undefined ? null : it[k];
      });
    } else {
      Object.keys(it).forEach(function (k) { if (userFields.indexOf(k) < 0) out[k] = it[k]; });
      userFields.forEach(function (k) { out[k + 'Id'] = idOf(it[k]); });
      Object.keys(l.fields).forEach(function (k) { if (!(k in out) && userFields.indexOf(k) < 0) out[k] = null; });
    }
    out.Id = it.Id; out.ID = it.Id;
    return out;
  }
  function idOf(v) { return Array.isArray(v) ? v.map(function (p) { return p.Id; }) : v ? v.Id : null; }
  function httpErr(status, msg) { var e = new Error(msg); e.status = status; return e; }

  // --- a small OData $filter evaluator (eq ne gt ge lt le and or not, substringof, startswith, datetime'…')
  function tokenize(s) {
    var t = [], re = /\s*(datetime'[^']*'|'(?:[^']|'')*'|\(|\)|,|-?\d+(?:\.\d+)?(?![\w/])|[A-Za-z_][\w/]*)/g, m, last = 0;
    while ((m = re.exec(s)) && m.index === last) { t.push(m[1]); last = re.lastIndex; }
    if (s.slice(last).trim()) throw httpErr(400, 'Could not parse $filter near: ' + s.slice(last));
    return t;
  }
  function parseFilter(s) {
    var t = tokenize(s), i = 0;
    function peek() { return t[i]; } function next() { return t[i++]; }
    function expr() { var a = and(); while (/^or$/i.test(peek() || '')) { next(); var b = and(), A = a; a = (function (x, y) { return function (r) { return x(r) || y(r); }; })(A, b); } return a; }
    function and() { var a = unary(); while (/^and$/i.test(peek() || '')) { next(); var b = unary(), A = a; a = (function (x, y) { return function (r) { return x(r) && y(r); }; })(A, b); } return a; }
    function unary() { if (/^not$/i.test(peek() || '')) { next(); var a = unary(); return function (r) { return !a(r); }; } return cmp(); }
    function cmp() {
      if (peek() === '(') { next(); var e = expr(); next(); return e; }
      var left = val();
      var op = (peek() || '').toLowerCase();
      if (['eq', 'ne', 'gt', 'ge', 'lt', 'le'].indexOf(op) < 0) return function (r) { return !!left(r); };
      next(); var right = val();
      return function (r) {
        var a = norm(left(r)), b = norm(right(r));
        if (Array.isArray(a)) return op === 'eq' ? a.indexOf(b) >= 0 : op === 'ne' ? a.indexOf(b) < 0 : false;
        switch (op) { case 'eq': return a === b; case 'ne': return a !== b; case 'gt': return a > b; case 'ge': return a >= b; case 'lt': return a < b; case 'le': return a <= b; }
      };
    }
    function val() {
      var tok = next();
      if (tok === undefined) throw httpErr(400, 'Unexpected end of $filter');
      if (/^datetime'/i.test(tok)) { var d = tok.slice(9, -1); return function () { return new Date(d).getTime(); }; }
      if (tok[0] === "'") { var s2 = tok.slice(1, -1).replace(/''/g, "'"); return function () { return s2; }; }
      if (/^-?\d/.test(tok)) { var n = Number(tok); return function () { return n; }; }
      if (/^(true|false)$/i.test(tok)) { var b = /^true$/i.test(tok); return function () { return b; }; }
      if (/^null$/i.test(tok)) return function () { return null; };
      if (/^(substringof|startswith|endswith)$/i.test(tok) && peek() === '(') {
        next(); var a1 = val(); next(); var a2 = val(); next();
        var fn = tok.toLowerCase();
        return function (r) {
          var x = String(norm(a1(r)) ?? '').toLowerCase(), y = String(norm(a2(r)) ?? '').toLowerCase();
          return fn === 'substringof' ? y.indexOf(x) >= 0 : fn === 'startswith' ? x.indexOf(y) === 0 : x.slice(-y.length) === y;
        };
      }
      return function (r) { return get(r, tok); };
    }
    var f = expr();
    if (i < t.length) throw httpErr(400, 'Could not parse $filter near: ' + t.slice(i).join(' '));
    return f;
  }
  function norm(v) {
    if (v === true) return 1; if (v === false) return 0;
    if (typeof v === 'string' && /^\d{4}-\d\d-\d\dT/.test(v)) return new Date(v).getTime();
    return v;
  }
  function get(it, path) {
    var parts = path.split('/'), k = parts[0];
    if (k === 'ID') k = 'Id';
    if (/Id$/.test(k) && it[k] === undefined && it[k.slice(0, -2)] !== undefined) return idOf(it[k.slice(0, -2)]);
    var v = it[k];
    if (parts.length > 1) { if (Array.isArray(v)) return v.map(function (p) { return p[parts[1]]; }); return v ? v[parts[1]] : null; }
    if (v && typeof v === 'object' && !Array.isArray(v) && 'Url' in v) return v.Url;
    return v;
  }

  function route(method, url, headers, body) {
    var u = new URL(url), path = decodeURIComponent(u.pathname), q = {};
    u.searchParams.forEach(function (v, k) { q[k] = v; });
    var verb = (headers['X-HTTP-Method'] || headers['x-http-method'] || method).toUpperCase();
    var verbose = /odata=verbose/.test(headers['Accept'] || headers['accept'] || '');
    var api = path.replace(WEB.serverRelativeUrl, '').replace(/^\/_api\//, '');
    if (path.indexOf(WEB.serverRelativeUrl + '/_api/') !== 0) throw httpErr(404, 'The preview site is ' + WEB.absoluteUrl + '. Other sites are not available in the preview.');
    var wrap = function (o) { return verbose ? { d: Array.isArray(o) ? { results: o } : o } : Array.isArray(o) ? { value: o } : o; };
    var m;
    if (api === 'web' || api === 'web/') return wrap({ Title: WEB.title, Url: WEB.absoluteUrl, ServerRelativeUrl: WEB.serverRelativeUrl });
    if (api === 'web/currentuser') return wrap({ Id: ME.Id, Title: ME.Title, Email: ME.EMail, LoginName: ME.LoginName });
    if (api === 'web/ensureuser') { var p = person((body && (body.logonName || body.logonname)) || 'Guest'); return wrap({ Id: p.Id, Title: p.Title, Email: p.EMail, LoginName: p.Name }); }
    if (api === 'web/lists' && verb === 'GET') {
      var ls = Object.keys(DB).map(function (k) { return DB[k]; }).filter(function (l) { return !(q.$filter && /Hidden eq false/i.test(q.$filter) && l.Hidden); });
      if (q.$filter && /BaseTemplate eq (\d+)/.test(q.$filter)) { var bt = Number(RegExp.$1); ls = ls.filter(function (l) { return l.BaseTemplate === bt; }); }
      return wrap(ls.sort(function (a, b) { return a.Title.localeCompare(b.Title); }).map(function (l) { return { Title: l.Title, Id: l.Id, ItemCount: l.items.length, BaseTemplate: l.BaseTemplate }; }));
    }
    if (api === 'web/lists' && verb === 'POST') {
      if (DB[String(body.Title).toLowerCase()]) throw httpErr(500, 'A list, survey, discussion board, or document library with the specified title already exists in this Web site.');
      var nl = newList(body.Title, [], body.BaseTemplate); post({ type: 'db' }); return wrap({ Id: nl.Id, Title: nl.Title });
    }
    if ((m = /^web\/lists\/getbytitle\('((?:[^']|'')*)'\)(.*)$/i.exec(api))) {
      var title = m[1].replace(/''/g, "'"), rest = m[2], l = DB[title.toLowerCase()];
      if (!l) throw httpErr(404, "List '" + title + "' does not exist at site with URL '" + WEB.absoluteUrl + "'.");
      if (rest === '' ) return wrap({ Id: l.Id, Title: l.Title, ItemCount: l.items.length });
      if (rest === '/fields' || rest === '/fields/') return wrap(Object.keys(l.fields).map(function (k) { var f = l.fields[k]; return { Title: f.Title, InternalName: f.InternalName, TypeAsString: f.TypeAsString, Required: f.Required, Choices: f.Choices || [] }; }).filter(function (f) { return !/^(Created|Modified|Author|Editor)$/.test(f.InternalName) || !/ReadOnlyField eq false/.test(q.$filter || ''); }));
      if (rest === '/fields/createfieldasxml') {
        var xml = body.parameters.SchemaXml, gx = function (a) { var r = new RegExp(a + "='([^']*)'").exec(xml); return r ? r[1] : ''; };
        var type = gx('Type'), fmt = gx('Format');
        var choices = []; xml.replace(/<CHOICE>([^<]*)<\/CHOICE>/g, function (_, c) { choices.push(c); });
        addField(l, { name: gx('DisplayName') || gx('Name'), type: type === 'URL' && fmt === 'Image' ? 'Image' : type, required: gx('Required') === 'TRUE', choices: choices });
        l.fields[gx('Name')] = l.fields[iname(gx('DisplayName') || gx('Name'))]; post({ type: 'db' });
        return wrap({ InternalName: gx('Name') });
      }
      if (rest === '/items' && verb === 'GET') {
        var items = l.items.slice();
        if (q.$filter) { var f = parseFilter(q.$filter); items = items.filter(function (r) { return f(r); }); }
        if (q.$orderby) {
          var keys = q.$orderby.split(',').map(function (s) { var p2 = s.trim().split(/\s+/); return [p2[0], /desc/i.test(p2[1] || '') ? -1 : 1]; });
          items.sort(function (a, b) { for (var i = 0; i < keys.length; i++) { var x = norm(get(a, keys[i][0])), y = norm(get(b, keys[i][0])); if (x === y) continue; if (x === null || x === undefined) return 1; if (y === null || y === undefined) return -1; return (x > y ? 1 : -1) * keys[i][1]; } return 0; });
        }
        var skip = Number(q.$skip || 0), top = Number(q.$top || 100);
        var page = items.slice(skip, skip + top);
        var sel = q.$select ? q.$select.split(',').map(function (s) { return s.trim(); }) : null, exp = q.$expand ? q.$expand.split(',').map(function (s) { return s.trim(); }) : [];
        var out = { value: page.map(function (it) { return ser(l, it, sel, exp); }) };
        if (skip + top < items.length) { u.searchParams.set('$skip', skip + top); out['odata.nextLink'] = u.toString(); }
        return verbose ? { d: { results: out.value } } : out;
      }
      if (rest === '/items' && verb === 'POST') { var it = newItem(l, ME); write(l, it, body); post({ type: 'db' }); return wrap(ser(l, it)); }
      if ((m = /^\/items\((\d+)\)$/.exec(rest))) {
        var item = l.items.filter(function (r) { return r.Id === Number(m[1]); })[0];
        if (!item) throw httpErr(404, 'Item does not exist. It may have been deleted by another user.');
        if (verb === 'GET') return wrap(ser(l, item, q.$select ? q.$select.split(',').map(function (s) { return s.trim(); }) : null, q.$expand ? q.$expand.split(',') : []));
        if (verb === 'MERGE') { write(l, item, body); item.Modified = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'); item.Editor = person(ME.Title); post({ type: 'db' }); return null; }
        if (verb === 'DELETE') { l.items.splice(l.items.indexOf(item), 1); post({ type: 'db' }); return null; }
      }
    }
    throw httpErr(400, 'The preview does not emulate ' + verb + ' ' + path + '. In SharePoint this call may work; test it there.');
  }
  function write(l, it, body) {
    Object.keys(body || {}).forEach(function (k) {
      if (k === '__metadata') return;
      var name = k, f = l.fields[k];
      if (!f && /Id$/.test(k) && l.fields[k.slice(0, -2)]) {
        name = k.slice(0, -2); f = l.fields[name];
        var v = body[k], arr = v && v.results ? v.results : v;
        it[name] = Array.isArray(arr) ? arr.map(function (id) { var p = PEOPLE.filter(function (x) { return x.Id === id; })[0]; return p ? person(p.Title) : { Id: id, Title: 'User ' + id }; })
          : arr === null ? null : (function () { var p = PEOPLE.filter(function (x) { return x.Id === arr; })[0]; return p ? person(p.Title) : { Id: arr, Title: 'User ' + arr }; })();
        return;
      }
      if (!f) throw httpErr(400, "The property '" + k + "' does not exist on type 'SP.Data." + l.Title.replace(/\W/g, '') + "ListItem'. Make sure to only use property names that are defined by the type.");
      if (/^User/.test(f.TypeAsString)) throw httpErr(400, "Set the person column '" + k + "' with " + k + 'Id (the user id), not ' + k + '.');
      var val = body[k]; if (val && val.results) val = val.results;
      it[k] = coerce(f, val);
    });
  }
  var spHttpClient = {
    _do: function (method, url, opts) {
      opts = opts || {}; var headers = opts.headers || {}, body = null;
      try { body = opts.body ? JSON.parse(opts.body) : null; } catch (e) { body = null; }
      return new Promise(function (resolve) {
        setTimeout(function () {
          var status = 200, data;
          try { data = route(method, url, headers, body); if (data === null) status = 204; else if (method === 'POST' && !headers['X-HTTP-Method']) status = 201; }
          catch (e) { status = e.status || 500; data = { 'odata.error': { code: '-1, Microsoft.SharePoint.SPException', message: { lang: 'en-US', value: e.message } } }; }
          var verb = headers['X-HTTP-Method'] || method;
          log.push({ verb: verb, url: url.replace(WEB.absoluteUrl, ''), status: status }); post({ type: 'rest', verb: verb, url: url.replace(WEB.absoluteUrl, ''), status: status, error: status >= 400 ? data['odata.error'].message.value : '' });
          var text = data === null || status === 204 ? '' : JSON.stringify(data);
          resolve({ ok: status < 400, status: status, statusText: status < 400 ? 'OK' : 'Error', text: function () { return Promise.resolve(text); }, json: function () { return text ? Promise.resolve(JSON.parse(text)) : Promise.reject(new Error('No content')); }, headers: { get: function () { return null; } } });
        }, 90 + Math.random() * 120);
      });
    },
    get: function (url, cfg, opts) { return this._do('GET', url, opts); },
    post: function (url, cfg, opts) { return this._do('POST', url, opts); },
    fetch: function (url, cfg, opts) { return this._do((opts && opts.method) || 'GET', url, opts); }
  };

  // ---------------------------------------------------------------- mounting
  var wp = null, Cls = null, host = document.getElementById('host'), paneFields = {};
  window.__sssReport = function (e) { post({ type: 'error', message: (e && e.message) || String(e), stack: e && e.stack ? String(e.stack).split('\n').slice(0, 6).join('\n') : '' }); };
  window.addEventListener('error', function (e) { window.__sssReport(e.error || e.message); });
  window.addEventListener('unhandledrejection', function (e) { window.__sssReport(e.reason); });
  function serialPane() {
    paneFields = {};
    var cfg; try { wp.onPropertyPaneConfigurationStart(); cfg = wp.getPropertyPaneConfiguration(); } catch (e) { window.__sssReport(e); return; }
    var pages = (cfg.pages || []).map(function (pg) {
      return { header: pg.header && pg.header.description, groups: (pg.groups || []).map(function (g) {
        return { name: g.groupName, fields: (g.groupFields || []).map(function (f) {
          var p = f.properties || {}; if (f.targetProperty) paneFields[f.targetProperty] = f;
          return { type: f.type, key: f.targetProperty, label: p.label || p.text, text: p.text, description: p.description, multiline: !!p.multiline, options: (p.options || []).map(function (o) { return { key: o.key, text: o.text }; }), min: p.min, max: p.max, step: p.step, onText: p.onText, offText: p.offText, value: f.targetProperty ? wp.properties[f.targetProperty] : undefined };
        }) };
      }) };
    });
    post({ type: 'pane', pages: pages });
  }
  function mount(msg) {
    if (wp) { try { wp.onDispose(); } catch (e) { /* ignore */ } }
    host.innerHTML = ''; defined = null;
    [].slice.call(document.querySelectorAll('style[id^="css-sss-"]')).forEach(function (s) { s.remove(); });
    var s = document.createElement('script'); s.textContent = msg.bundle; document.head.appendChild(s); s.remove();
    if (!defined || !defined.default) { post({ type: 'error', message: 'The bundle did not define a web part (see the error above).' }); return; }
    Cls = defined.default;
    wp = new Cls();
    wp.domElement = document.createElement('div'); host.appendChild(wp.domElement);
    wp.properties = Object.assign({}, msg.props || {});
    wp.displayMode = msg.edit ? DisplayMode.Edit : DisplayMode.Read;
    wp.context = { pageContext: { web: WEB, site: { absoluteUrl: WEB.absoluteUrl }, user: { displayName: ME.Title, email: ME.EMail, loginName: ME.LoginName } }, spHttpClient: spHttpClient, serviceScope: { consume: function () { return themeProvider; } }, propertyPane: { refresh: function () { serialPane(); }, isPropertyPaneOpen: function () { return true; } }, instanceId: wp.instanceId };
    Promise.resolve(wp.onInit()).then(function () { wp.render(); serialPane(); post({ type: 'mounted' }); }, window.__sssReport);
  }
  function setTheme(t) { currentTheme = t; document.body.style.background = (t && t.semanticColors && t.semanticColors.bodyBackground) || '#fff'; themeHandlers.forEach(function (fn) { try { fn({ theme: t }); } catch (e) { window.__sssReport(e); } }); }
  window.addEventListener('message', function (e) {
    var msg = e.data; if (!msg || !msg.__sss) return;
    if (msg.type === 'load') { setTheme(msg.theme || null); if (msg.reseed || !Object.keys(DB).length) seed(msg.lists); else seed(msg.lists, true); themeHandlers = []; mount(msg); }
    else if (msg.type === 'prop' && wp) {
      var f = paneFields[msg.key], old = wp.properties[msg.key];
      var err = f && f.properties && f.properties.onGetErrorMessage ? f.properties.onGetErrorMessage(msg.value) : '';
      post({ type: 'fieldError', key: msg.key, message: err || '' });
      if (err) return;
      wp.properties[msg.key] = msg.value; wp.onPropertyPaneFieldChanged(msg.key, old, msg.value); wp.render();
    }
    else if (msg.type === 'button' && wp) { var b = paneFields[msg.key]; if (b && b.properties.onClick) { try { b.properties.onClick(); } catch (x) { window.__sssReport(x); } } }
    else if (msg.type === 'mode' && wp) { wp.displayMode = msg.edit ? DisplayMode.Edit : DisplayMode.Read; wp.render(); }
    else if (msg.type === 'theme') setTheme(msg.theme);
    else if (msg.type === 'reseed') { seed(msg.lists); if (wp) { wp.onDispose(); wp.render(); } }
    else if (msg.type === 'dump') post({ type: 'dump', db: Object.keys(DB).map(function (k) { var l = DB[k]; return { title: l.Title, fields: Object.keys(l.fields), items: l.items.length }; }) });
  });
  // links and forms: keep the preview on its page and tell the studio
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[href]');
    if (!a || e.defaultPrevented) return;
    var href = a.getAttribute('href');
    if (!href || href === '#' || /^javascript:/i.test(href)) { e.preventDefault(); return; }
    if (/^mailto:|^tel:/i.test(href)) { e.preventDefault(); post({ type: 'link', href: href }); return; }
    e.preventDefault(); post({ type: 'link', href: a.href, newTab: a.target === '_blank' });
  });
  window.addEventListener('submit', function (e) {
    if (!e.defaultPrevented) { e.preventDefault(); post({ type: 'warn', message: 'A form was submitted without e.preventDefault(). On a SharePoint page that reloads the page.' }); }
  });
  if (window.ResizeObserver) new ResizeObserver(function () { post({ type: 'height', height: Math.ceil(document.documentElement.scrollHeight) }); }).observe(document.body);
  post({ type: 'ready' });
}

/** The srcdoc for the preview iframe. */
export function previewDoc() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
html,body{margin:0;font-family:"Segoe UI","Segoe UI Web (West European)",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;font-size:14px;color:#323130;-webkit-font-smoothing:antialiased}
body{padding:24px 28px 40px;background:#fff;transition:background .2s}
#host{max-width:100%}
a{color:inherit}
</style></head><body><div id="host"></div><script>(${previewHost.toString()})();<\/script></body></html>`;
}
