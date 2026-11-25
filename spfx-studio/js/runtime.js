// runtime.js — the small library every generated web part carries inside its bundle.
// spfxRuntime() is serialised with toString() into the AMD module (see bundle.js), so it must not touch anything
// outside its own arguments. It turns a spec (properties, lists, CSS, app code) into a real SPFx web part class:
// property pane, list provisioning buttons, theme variables, SharePoint REST helpers, error handling.
// The same text runs in the preview (against preview.js's stand-in SPFx modules) and in SharePoint.

/* eslint-disable */
export function spfxRuntime(m, SPEC, CSS, APP) {
  'use strict';
  var BaseWP = m.wb.BaseClientSideWebPart, PP = m.pp, V1 = m.http.SPHttpClient.configurations.v1;
  var ThemeProvider = m.cb && m.cb.ThemeProvider;
  var SCOPE = SPEC.scope; // e.g. sss-a1b2c3d4: every CSS rule in CSS is already prefixed with .SCOPE
  var SVGNS = 'http://www.w3.org/2000/svg';
  var SVG_TAGS = { svg: 1, path: 1, circle: 1, rect: 1, line: 1, polyline: 1, polygon: 1, g: 1, ellipse: 1, defs: 1, linearGradient: 1, stop: 1, text: 1 };

  // ---------------------------------------------------------------- icons (inline SVG, no font or CDN needed)
  var ICONS = {
    add: 'M12 5v14M5 12h14', close: 'M6 6l12 12M18 6L6 18', check: 'M5 12l5 5 9-10', chevronRight: 'M9 6l6 6-6 6',
    chevronLeft: 'M15 6l-6 6 6 6', chevronDown: 'M6 9l6 6 6-6', chevronUp: 'M6 15l6-6 6 6', search: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM20 20l-4-4',
    edit: 'M4 20h4L19 9l-4-4L4 16v4zM14 6l4 4', trash: 'M5 7h14M10 7V5h4v2M7 7l1 13h8l1-13', link: 'M10 14a4 4 0 0 0 6 0l3-3a4 4 0 0 0-6-6l-1 1M14 10a4 4 0 0 0-6 0l-3 3a4 4 0 0 0 6 6l1-1',
    external: 'M14 4h6v6M20 4l-9 9M18 14v5H5V6h5', calendar: 'M4 6h16v14H4zM4 10h16M8 3v4M16 3v4', clock: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM12 7v5l3 2',
    person: 'M12 4a4 4 0 1 0 0 8 4 4 0 0 0 0-8zM4 21c1-4 4-6 8-6s7 2 8 6', people: 'M9 5a3.5 3.5 0 1 0 0 7 3.5 3.5 0 0 0 0-7zM2 20c.8-3.5 3.4-5.5 7-5.5s6.2 2 7 5.5M16 5.5a3 3 0 0 1 0 6M18 14.8c2 .7 3.3 2.4 4 5.2',
    mail: 'M3 6h18v12H3zM3 7l9 6 9-6', phone: 'M5 4h4l2 5-3 2a11 11 0 0 0 5 5l2-3 5 2v4a2 2 0 0 1-2 2A17 17 0 0 1 3 6a2 2 0 0 1 2-2z',
    location: 'M12 21s-7-6.2-7-12a7 7 0 0 1 14 0c0 5.8-7 12-7 12zM12 6.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z', star: 'M12 3l2.8 5.7 6.2.9-4.5 4.4 1 6.2L12 17.3 6.5 20.2l1-6.2L3 9.6l6.2-.9z',
    heart: 'M12 20s-8-4.8-8-10.5A4.5 4.5 0 0 1 12 7a4.5 4.5 0 0 1 8 2.5C20 15.2 12 20 12 20z', info: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM12 11v6M12 7.5v.5',
    warning: 'M12 3l10 18H2zM12 10v5M12 18v.5', megaphone: 'M3 10v4h3l7 4V6L6 10zM16 9a4 4 0 0 1 0 6M6 14l1 5h3l-1-4', document: 'M6 3h8l5 5v13H6zM14 3v5h5M9 13h7M9 17h7',
    folder: 'M3 6h6l2 2h10v11H3z', image: 'M3 5h18v14H3zM3 16l5-5 4 4 3-3 6 6M15.5 8.5h.01', video: 'M3 6h13v12H3zM16 10l5-3v10l-5-3', home: 'M3 11l9-7 9 7v9h-6v-6H9v6H3z',
    settings: 'M12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6zM19 12l2-1-1-3-2 .2-1.5-1.5.2-2-3-1-1 2h-2l-1-2-3 1 .2 2L5.5 7.2 3.5 7l-1 3 2 1v2l-2 1 1 3 2-.2 1.5 1.5-.2 2 3 1 1-2h2l1 2 3-1-.2-2 1.5-1.5 2 .2 1-3-2-1z',
    filter: 'M3 5h18l-7 8v6l-4 2v-8z', sort: 'M7 4v16M3 16l4 4 4-4M17 20V4M13 8l4-4 4 4', refresh: 'M20 11a8 8 0 1 0-2 6M20 5v6h-6', download: 'M12 4v11M7 10l5 5 5-5M4 20h16',
    upload: 'M12 20V9M7 14l5-5 5 5M4 4h16', grid: 'M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z', list: 'M8 6h12M8 12h12M8 18h12M4 6h.01M4 12h.01M4 18h.01',
    bell: 'M6 16V11a6 6 0 0 1 12 0v5l2 2H4zM10 20a2 2 0 0 0 4 0', bookmark: 'M6 3h12v18l-6-4-6 4z', chart: 'M4 20V10M10 20V4M16 20v-7M22 20H2',
    trophy: 'M8 4h8v5a4 4 0 0 1-8 0zM8 6H4v2a4 4 0 0 0 4 4M16 6h4v2a4 4 0 0 1-4 4M12 13v4M8 21h8M9 17h6', lightbulb: 'M9 18h6M10 21h4M12 3a6 6 0 0 0-4 10.5c.7.7 1 1.5 1 2.5h6c0-1 .3-1.8 1-2.5A6 6 0 0 0 12 3z',
    globe: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM3 12h18M12 3c3 3.5 3 14.5 0 18M12 3c-3 3.5-3 14.5 0 18', lock: 'M6 11h12v10H6zM8 11V7a4 4 0 0 1 8 0v4',
    tag: 'M3 3h8l10 10-8 8L3 11zM7.5 7.5h.01', flag: 'M5 21V4M5 4h11l-2 4 2 4H5', briefcase: 'M3 8h18v12H3zM8 8V5h8v3M3 13h18', book: 'M4 4h7a3 3 0 0 1 3 3v14a2 2 0 0 0-2-2H4zM20 4h-6M20 4v15h-6',
    rocket: 'M12 15l-3-3c1-5 4-8 10-9-1 6-4 9-9 10zM9 12l-4 1-2 3 4 1M12 15l-1 4 3 2 1-4M15 9h.01', sparkle: 'M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8zM19 16l.8 2.2L22 19l-2.2.8L19 22l-.8-2.2L16 19l2.2-.8z',
    more: 'M5 12h.01M12 12h.01M19 12h.01', play: 'M7 4l13 8-13 8z', pause: 'M7 4h4v16H7zM13 4h4v16h-4z', arrowRight: 'M4 12h16M14 6l6 6-6 6', arrowLeft: 'M20 12H4M10 6l-6 6 6 6'
  };
  function icon(name, size, extra) {
    var s = document.createElementNS(SVGNS, 'svg');
    s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('width', size || 18); s.setAttribute('height', size || 18);
    s.setAttribute('fill', 'none'); s.setAttribute('stroke', 'currentColor'); s.setAttribute('stroke-width', '1.8');
    s.setAttribute('stroke-linecap', 'round'); s.setAttribute('stroke-linejoin', 'round'); s.setAttribute('aria-hidden', 'true');
    s.setAttribute('class', 'sss-icon' + (extra ? ' ' + extra : ''));
    var p = document.createElementNS(SVGNS, 'path'); p.setAttribute('d', ICONS[name] || ICONS.info); s.appendChild(p);
    return s;
  }

  // ---------------------------------------------------------------- DOM helpers
  function h(tag, attrs) {
    var el = SVG_TAGS[tag] ? document.createElementNS(SVGNS, tag) : document.createElement(tag);
    if (attrs && (typeof attrs !== 'object' || attrs.nodeType || Array.isArray(attrs))) { append(el, [].slice.call(arguments, 1)); return el; }
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'class' || k === 'className') { if (SVG_TAGS[tag]) el.setAttribute('class', v); else el.className = v; }
      else if (k === 'style' && typeof v === 'object') Object.keys(v).forEach(function (s) { if (s.indexOf('--') === 0) el.style.setProperty(s, v[s]); else el.style[s] = v[s]; });
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'text') el.textContent = v;
      else if (k === 'html') el.innerHTML = v;
      else if (k === 'dataset' && typeof v === 'object') Object.keys(v).forEach(function (d) { el.dataset[d] = v[d]; });
      else if ((k === 'value' || k === 'checked' || k === 'selected' || k === 'disabled') && !SVG_TAGS[tag]) el[k] = v;
      else el.setAttribute(k, v === true ? '' : String(v));
    });
    append(el, [].slice.call(arguments, 2));
    return el;
  }
  function append(el, kids) {
    kids.forEach(function (c) {
      if (c === null || c === undefined || c === false || c === true) return;
      if (Array.isArray(c)) append(el, c);
      else el.appendChild(c.nodeType ? c : document.createTextNode(String(c)));
    });
    return el;
  }
  function esc(s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); return el; }

  // ---------------------------------------------------------------- formatting
  function toDate(d) { if (!d) return null; var x = d instanceof Date ? d : new Date(d); return isNaN(x) ? null : x; }
  var fmt = {
    date: function (d, o) { d = toDate(d); return d ? d.toLocaleDateString(undefined, o || { year: 'numeric', month: 'short', day: 'numeric' }) : ''; },
    dateTime: function (d) { d = toDate(d); return d ? d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : ''; },
    time: function (d) { d = toDate(d); return d ? d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' }) : ''; },
    number: function (n, o) { return n === null || n === undefined || n === '' || isNaN(n) ? '' : Number(n).toLocaleString(undefined, o); },
    relative: function (d) {
      d = toDate(d); if (!d) return '';
      var s = Math.round((d.getTime() - Date.now()) / 1000), a = Math.abs(s), u = [[60, 'second', 1], [3600, 'minute', 60], [86400, 'hour', 3600], [604800, 'day', 86400], [2629800, 'week', 604800], [31557600, 'month', 2629800], [Infinity, 'year', 31557600]];
      for (var i = 0; i < u.length; i++) if (a < u[i][0]) { var v = Math.round(s / u[i][2]); try { return new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' }).format(v, u[i][1]); } catch (e) { return fmt.date(d); } }
      return fmt.date(d);
    },
    initials: function (name) { return String(name || '').split(/\s+/).filter(Boolean).slice(0, 2).map(function (w) { return w[0].toUpperCase(); }).join(''); },
    truncate: function (s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, '') + '…' : s; },
    plain: function (html) { var d = document.createElement('div'); d.innerHTML = String(html || ''); return (d.textContent || '').replace(/\s+/g, ' ').trim(); }
  };
  function debounce(fn, ms) { var t; return function () { var a = arguments, me = this; clearTimeout(t); t = setTimeout(function () { fn.apply(me, a); }, ms || 250); }; }

  // ---------------------------------------------------------------- SharePoint REST
  function quote(s) { return "'" + encodeURIComponent(String(s).replace(/'/g, "''")) + "'"; }
  function listPath(name) { return "/_api/web/lists/getbytitle(" + quote(name) + ")"; }
  var FIELD_XML = {
    Text: function (n) { return "<Field Type='Text' MaxLength='255'"; },
    Note: function () { return "<Field Type='Note' NumLines='6' RichText='FALSE'"; },
    RichText: function () { return "<Field Type='Note' NumLines='8' RichText='TRUE' RichTextMode='FullHtml'"; },
    Number: function () { return "<Field Type='Number'"; },
    Currency: function () { return "<Field Type='Currency' LCID='1033'"; },
    DateTime: function () { return "<Field Type='DateTime' Format='DateOnly'"; },
    DateTimeFull: function () { return "<Field Type='DateTime' Format='DateTime'"; },
    Boolean: function () { return "<Field Type='Boolean'"; },
    Choice: function () { return "<Field Type='Choice' Format='Dropdown'"; },
    MultiChoice: function () { return "<Field Type='MultiChoice'"; },
    URL: function () { return "<Field Type='URL' Format='Hyperlink'"; },
    Image: function () { return "<Field Type='URL' Format='Image'"; },
    User: function () { return "<Field Type='User' UserSelectionMode='PeopleOnly'"; },
    UserMulti: function () { return "<Field Type='UserMulti' UserSelectionMode='PeopleOnly' Mult='TRUE'"; }
  };
  function xmlEsc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]; }); }
  function internalName(n) { return String(n).replace(/[^A-Za-z0-9_]/g, ''); }
  function fieldXml(c) {
    var t = FIELD_XML[c.type] ? c.type : 'Text', iname = internalName(c.name);
    var x = FIELD_XML[t](c.name) + " Name='" + iname + "' StaticName='" + iname + "' DisplayName='" + xmlEsc(c.name) + "' Required='" + (c.required ? 'TRUE' : 'FALSE') + "'";
    var choices = (c.choices || []).filter(Boolean);
    if ((t === 'Choice' || t === 'MultiChoice') && choices.length) return x + '><CHOICES>' + choices.map(function (ch) { return '<CHOICE>' + xmlEsc(ch) + '</CHOICE>'; }).join('') + '</CHOICES>' + (t === 'Choice' ? '<Default>' + xmlEsc(choices[0]) + '</Default>' : '') + '</Field>';
    if (t === 'Boolean') return x + '><Default>0</Default></Field>';
    return x + ' />';
  }

  function spClient(wp, siteUrl) {
    var http = wp.context.spHttpClient;
    var base = String(siteUrl || wp.context.pageContext.web.absoluteUrl).replace(/\/+$/, '');
    function call(method, path, body, verbose, extra) {
      var url = /^https?:/i.test(path) ? path : base + path;
      var type = verbose ? 'application/json;odata=verbose' : 'application/json;odata=nometadata';
      var headers = { 'Accept': type, 'odata-version': '' };
      if (extra) Object.keys(extra).forEach(function (k) { headers[k] = extra[k]; });
      var opts = { headers: headers };
      if (body !== undefined) { headers['Content-Type'] = type; opts.body = JSON.stringify(body); }
      return (method === 'GET' ? http.get(url, V1, opts) : http.post(url, V1, opts)).then(function (r) {
        if (r.status === 204) return null;
        return r.text().then(function (t) {
          var j = null; try { j = t ? JSON.parse(t) : null; } catch (e) { /* not JSON */ }
          if (!r.ok) {
            var msg = j && (j['odata.error'] || j.error) ? ((j['odata.error'] || j.error).message || {}).value || (j['odata.error'] || j.error).message : '';
            var err = new Error((msg || 'SharePoint returned ' + r.status) + (r.status === 404 ? ' (not found)' : '')); err.status = r.status; throw err;
          }
          return j;
        });
      });
    }
    function query(o) {
      o = o || {}; var q = [];
      if (o.select) q.push('$select=' + [].concat(o.select).join(','));
      if (o.expand) q.push('$expand=' + [].concat(o.expand).join(','));
      if (o.filter) q.push('$filter=' + encodeURIComponent(o.filter));
      if (o.orderBy) q.push('$orderby=' + encodeURIComponent([].concat(o.orderBy).join(',')));
      if (o.top !== undefined || !o.noTop) q.push('$top=' + (o.top || 500));
      return q.length ? '?' + q.join('&') : '';
    }
    var api = {
      siteUrl: base,
      at: function (url) { return spClient(wp, url); },
      get: function (path) { return call('GET', path); },
      post: function (path, body, headers) { return call('POST', path, body, false, headers); },
      items: function (list, o) {
        o = o || {};
        var all = [];
        function page(url) {
          return call('GET', url).then(function (j) {
            all = all.concat((j && j.value) || []);
            var next = j && (j['odata.nextLink'] || j['@odata.nextLink']);
            return o.all && next && all.length < 5000 ? page(next) : all;
          });
        }
        return page(listPath(list) + '/items' + query(o));
      },
      item: function (list, id, o) { return call('GET', listPath(list) + '/items(' + Number(id) + ')' + query(Object.assign({ noTop: true }, o || {}))); },
      add: function (list, data) { return call('POST', listPath(list) + '/items', data); },
      update: function (list, id, data) { return call('POST', listPath(list) + '/items(' + Number(id) + ')', data, false, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }); },
      remove: function (list, id) { return call('POST', listPath(list) + '/items(' + Number(id) + ')', undefined, false, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }); },
      me: function () { return call('GET', '/_api/web/currentuser?$select=Id,Title,Email,LoginName'); },
      ensureUser: function (login) { return call('POST', '/_api/web/ensureuser', { logonName: login }); },
      lists: function () { return call('GET', "/_api/web/lists?$select=Title,Id,ItemCount,BaseTemplate&$filter=Hidden eq false&$orderby=Title").then(function (j) { return (j && j.value) || []; }); },
      fields: function (list) { return call('GET', listPath(list) + "/fields?$select=Title,InternalName,TypeAsString,Required,Choices&$filter=Hidden eq false and ReadOnlyField eq false").then(function (j) { return (j && j.value) || []; }); },
      listExists: function (list) { return call('GET', listPath(list) + '?$select=Id').then(function () { return true; }, function (e) { if (e.status === 404) return false; throw e; }); },
      ensureList: function (def, progress) {
        var name = def.name, made = [];
        progress = progress || function () {};
        return api.listExists(name).then(function (exists) {
          if (exists) return null;
          progress('Creating list ' + name + '…');
          made.push('list');
          return call('POST', '/_api/web/lists', { __metadata: { type: 'SP.List' }, Title: name, BaseTemplate: 100, Description: def.description || '' }, true);
        }).then(function () {
          return api.fields(name);
        }).then(function (have) {
          var names = {}; have.forEach(function (f) { names[f.InternalName.toLowerCase()] = 1; names[String(f.Title).toLowerCase()] = 1; });
          var cols = (def.columns || []).filter(function (c) { return c.name && c.name !== 'Title' && !names[internalName(c.name).toLowerCase()] && !names[c.name.toLowerCase()]; });
          return cols.reduce(function (p, c) {
            return p.then(function () {
              progress('Adding column ' + c.name + '…'); made.push(c.name);
              return call('POST', listPath(name) + '/fields/createfieldasxml', { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: fieldXml(c), Options: 8 } }, true);
            });
          }, Promise.resolve());
        }).then(function () {
          if (!def.sample || !def.sample.length || made[0] !== 'list') return made;
          progress('Adding sample items…');
          return def.sample.reduce(function (p, row) { return p.then(function () { return api.add(name, sampleToItem(def, row)); }); }, Promise.resolve()).then(function () { return made; });
        });
      }
    };
    return api;
  }
  function sampleToItem(def, row) {
    var o = {};
    (def.columns || []).forEach(function (c, i) {
      var v = row[i]; if (v === undefined || v === null || v === '') return;
      var k = c.name === 'Title' ? 'Title' : internalName(c.name);
      if (c.type === 'User' || c.type === 'UserMulti') return; // people need real accounts
      if (c.type === 'Number' || c.type === 'Currency') o[k] = Number(v);
      else if (c.type === 'Boolean') o[k] = /^(true|yes|1)$/i.test(String(v));
      else if (c.type === 'URL' || c.type === 'Image') o[k] = { Url: String(v), Description: c.type === 'Image' ? '' : String(v) };
      else if (c.type === 'MultiChoice') o[k] = String(v).split(/\s*;\s*/);
      else o[k] = String(v);
    });
    return o;
  }

  // ---------------------------------------------------------------- theme
  function readTheme(wp) {
    var t = wp._theme || (window.__themeState__ && window.__themeState__.theme) || {};
    var sc = t.semanticColors || {}, pal = t.palette || t;
    var pick = function () { for (var i = 0; i < arguments.length; i++) if (arguments[i]) return arguments[i]; };
    return {
      primary: pick(pal.themePrimary, '#0f6cbd'), primaryDark: pick(pal.themeDark, pal.themeDarker, '#0c3b5e'), primaryLight: pick(pal.themeLighter, '#deecf9'),
      primaryLighter: pick(pal.themeLighterAlt, '#eff6fc'), text: pick(sc.bodyText, pal.neutralPrimary, '#323130'), muted: pick(sc.bodySubtext, pal.neutralSecondary, '#605e5c'),
      background: pick(sc.bodyBackground, pal.white, '#ffffff'), surface: pick(sc.bodyStandoutBackground, pal.neutralLighterAlt, '#faf9f8'),
      border: pick(sc.bodyDivider, pal.neutralLight, '#edebe9'), link: pick(sc.link, pal.themePrimary, '#0f6cbd'),
      isInverted: !!t.isInverted
    };
  }
  function applyTheme(root, th) {
    var v = { '--sss-primary': th.primary, '--sss-primary-dark': th.primaryDark, '--sss-primary-light': th.primaryLight, '--sss-primary-lighter': th.primaryLighter,
      '--sss-text': th.text, '--sss-muted': th.muted, '--sss-bg': th.background, '--sss-surface': th.surface, '--sss-border': th.border, '--sss-link': th.link };
    Object.keys(v).forEach(function (k) { root.style.setProperty(k, v[k]); });
  }
  function injectCss() {
    var id = 'css-' + SCOPE + '-' + SPEC.version;
    if (document.getElementById(id)) return;
    var s = document.createElement('style'); s.id = id;
    s.textContent = '.' + SCOPE + '{container-type:inline-size;color:var(--sss-text);font-family:inherit;}.' + SCOPE + ' *{box-sizing:border-box}.' + SCOPE + ' .sss-icon{flex:none;vertical-align:middle}' +
      '.' + SCOPE + ' .sss-error{border:1px solid #f1bbbc;background:#fdf3f4;color:#a4262c;border-radius:8px;padding:12px 14px;font-size:13px;line-height:1.45}.' + SCOPE + ' .sss-error pre{white-space:pre-wrap;margin:8px 0 0;font-size:11px;color:#605e5c}' +
      '.' + SCOPE + ' .sss-spin{display:flex;align-items:center;gap:10px;color:var(--sss-muted);font-size:13px;padding:16px 4px}.' + SCOPE + ' .sss-spin:before{content:"";width:16px;height:16px;border-radius:50%;border:2px solid var(--sss-border);border-top-color:var(--sss-primary);animation:sss-rot .8s linear infinite}@keyframes sss-rot{to{transform:rotate(360deg)}}' +
      CSS;
    document.head.appendChild(s);
  }

  // ---------------------------------------------------------------- property values
  function propValue(p, raw) {
    if (p.type === 'number' || p.type === 'slider') { var n = Number(raw); return raw === '' || raw === null || raw === undefined || isNaN(n) ? (p.default === '' ? null : Number(p.default)) : n; }
    if (p.type === 'toggle') return raw === undefined || raw === null ? /^(true|yes|1)$/i.test(String(p.default)) : !!raw;
    if (raw === undefined || raw === null) return p.default === undefined ? '' : p.default;
    return raw;
  }

  // ---------------------------------------------------------------- the web part
  class StudioWebPart extends BaseWP {
    constructor() { super(); this._state = {}; this._cache = {}; this._pane = {}; this._run = 0; this._siteLists = null; }
    onInit() {
      var me = this;
      try {
        if (ThemeProvider && this.context.serviceScope) {
          var tp = this.context.serviceScope.consume(ThemeProvider.serviceKey);
          this._theme = tp.tryGetTheme();
          tp.themeChangedEvent.add(this, function (a) { me._theme = a.theme; me.render(); });
        }
      } catch (e) { /* theme service not available */ }
      return Promise.resolve();
    }
    get dataVersion() { return m.core.Version.parse('1.0'); }
    get disableReactivePropertyChanges() { return false; }
    _props() {
      var out = {}, me = this;
      (SPEC.properties || []).forEach(function (p) { out[p.key] = propValue(p, me.properties[p.key]); });
      Object.keys(this.properties || {}).forEach(function (k) { if (!(k in out)) out[k] = me.properties[k]; });
      return out;
    }
    _ctx(root, run) {
      var me = this, pc = this.context.pageContext, user = pc.user || {};
      var editMode = this.displayMode === (m.core.DisplayMode ? m.core.DisplayMode.Edit : 2);
      var ctx = {
        root: root, props: this._props(), state: this._state, h: h, esc: esc, icon: icon, clear: clear, fmt: fmt, debounce: debounce,
        sp: spClient(this), lists: SPEC.lists || [],
        user: { displayName: user.displayName || '', email: user.email || '', loginName: user.loginName || '' },
        page: { webUrl: pc.web.absoluteUrl, webTitle: pc.web.title, serverRelativeUrl: pc.web.serverRelativeUrl, editMode: editMode, instanceId: this.instanceId || '' },
        theme: readTheme(this),
        isCurrent: function () { return run === me._run; },
        cache: function (key, fn) { if (!(key in me._cache)) me._cache[key] = Promise.resolve().then(fn).catch(function (e) { delete me._cache[key]; throw e; }); return me._cache[key]; },
        refresh: function () { me._cache = {}; me.render(); },
        rerender: function () { me.render(); },
        setProp: function (key, value) { me.properties[key] = value; me.render(); },
        listDef: function (name) { return (SPEC.lists || []).filter(function (l) { return !name || l.name === name; })[0] || null; },
        ensureList: function (name) { var d = ctx.listDef(name); return d ? ctx.sp.ensureList(d) : Promise.reject(new Error('No list called ' + name + ' in this web part')); },
        loading: function (text) { return h('div', { class: 'sss-spin', role: 'status' }, text || 'Loading…'); },
        showError: function (e, where) { return errorBox(e, where, editMode); }
      };
      return ctx;
    }
    render() {
      var me = this, run = ++this._run;
      injectCss();
      var root = this.domElement.querySelector(':scope > .' + SCOPE);
      if (!root) { clear(this.domElement); root = h('div', { class: SCOPE + ' sss-root' }); this.domElement.appendChild(root); }
      var ctx = this._ctx(root, run);
      applyTheme(root, ctx.theme);
      root.classList.toggle('is-edit', ctx.page.editMode);
      var body = h('div', { class: 'sss-body' });
      var slow = setTimeout(function () { if (run === me._run && !body.firstChild) { clear(root).appendChild(ctx.loading()); } }, 200);
      var done = function () { clearTimeout(slow); if (run === me._run) { clear(root); root.appendChild(body); } };
      var fail = function (e) { clearTimeout(slow); if (run !== me._run) return; clear(root).appendChild(errorBox(e, 'render', ctx.page.editMode)); report(e); };
      try {
        ctx.root = body;
        var r = APP.call(null, ctx);
        if (r && typeof r.then === 'function') r.then(done, fail); else done();
      } catch (e) { fail(e); }
    }
    onDispose() { this._run++; this._cache = {}; }
    onPropertyPaneConfigurationStart() {
      var me = this;
      if (!this._siteLists && (SPEC.properties || []).some(function (p) { return p.type === 'list'; })) {
        spClient(this).lists().then(function (ls) { me._siteLists = ls.map(function (l) { return l.Title; }); me._refreshPane(); }, function () { me._siteLists = []; me._refreshPane(); });
      }
    }
    onPropertyPaneFieldChanged(path, oldValue, newValue) { if (oldValue !== newValue) this._cache = {}; }
    _refreshPane() { try { this.context.propertyPane.refresh(); } catch (e) { /* pane closed */ } }
    getPropertyPaneConfiguration() {
      var me = this, groups = {}, order = [];
      (SPEC.properties || []).forEach(function (p) {
        var g = p.group || 'Settings';
        if (!groups[g]) { groups[g] = []; order.push(g); }
        groups[g].push(paneField(p, me));
      });
      var pages = [{ header: { description: SPEC.description || '' }, groups: order.map(function (g) { return { groupName: g, groupFields: groups[g] }; }) }];
      if ((SPEC.lists || []).length) {
        var fields = [PP.PropertyPaneLabel('sssListsHelp', { text: 'These buttons create the lists this web part uses on this site, with every column it needs. Lists that already exist are left alone (missing columns are added).' })];
        SPEC.lists.forEach(function (l, i) {
          fields.push(PP.PropertyPaneButton('sssCreateList' + i, {
            text: 'Create "' + l.name + '" list', buttonType: PP.PropertyPaneButtonType ? PP.PropertyPaneButtonType.Normal : 0, icon: 'Add',
            onClick: function () {
              me._pane[i] = 'Working…'; me._refreshPane();
              spClient(me).ensureList(l, function (msg) { me._pane[i] = msg; me._refreshPane(); }).then(function (made) {
                me._pane[i] = made === null || (made && !made.length) ? 'Already there with every column.' : 'Done: ' + (made[0] === 'list' ? 'created the list' + (made.length > 1 ? ' and ' + (made.length - 1) + ' column(s)' : '') : 'added ' + made.join(', ')) + '.';
                me._cache = {}; me._refreshPane(); me.render();
              }, function (e) { me._pane[i] = 'Could not create it: ' + e.message + ' You need Manage Lists permission on this site.'; me._refreshPane(); });
              return '';
            }
          }));
          if (me._pane[i]) fields.push(PP.PropertyPaneLabel('sssListStatus' + i, { text: me._pane[i] }));
        });
        pages[0].groups.push({ groupName: 'Lists', groupFields: fields });
      }
      pages[0].groups.push({ groupName: 'About', groupFields: [PP.PropertyPaneLabel('sssAbout', { text: SPEC.title + ' ' + SPEC.version + '. Built with SPFx Studio.' })] });
      return { pages: pages };
    }
  }
  function paneField(p, wp) {
    var label = p.label || p.key, base = { label: label, description: p.description || undefined };
    switch (p.type) {
      case 'multiline': return PP.PropertyPaneTextField(p.key, Object.assign(base, { multiline: true, rows: 5, resizable: true }));
      case 'toggle': return PP.PropertyPaneToggle(p.key, { label: label, onText: 'On', offText: 'Off' });
      case 'dropdown': return PP.PropertyPaneDropdown(p.key, { label: label, options: (p.options || []).map(function (o) { return { key: o, text: o }; }) });
      case 'slider': return PP.PropertyPaneSlider(p.key, { label: label, min: Number(p.min) || 0, max: Number(p.max) || 100, step: Number(p.step) || 1, showValue: true });
      case 'list': {
        var names = (wp._siteLists || []).slice(), cur = wp.properties[p.key] || p.default;
        if (cur && names.indexOf(cur) < 0) names.unshift(cur);
        (SPEC.lists || []).forEach(function (l) { if (names.indexOf(l.name) < 0) names.push(l.name); });
        return PP.PropertyPaneDropdown(p.key, { label: label + (wp._siteLists ? '' : ' (loading lists…)'), options: names.map(function (n) { return { key: n, text: n }; }) });
      }
      case 'number': return PP.PropertyPaneTextField(p.key, Object.assign(base, { onGetErrorMessage: function (v) { return v === '' || !isNaN(Number(v)) ? '' : 'Enter a number'; }, deferredValidationTime: 300 }));
      case 'color': return PP.PropertyPaneTextField(p.key, Object.assign(base, { description: p.description || 'A colour such as #0f6cbd. Leave empty to follow the site theme.', onGetErrorMessage: function (v) { return !v || /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(v) ? '' : 'Use a hex colour like #0f6cbd'; } }));
      case 'url': return PP.PropertyPaneTextField(p.key, Object.assign(base, { onGetErrorMessage: function (v) { return !v || /^(https?:\/\/|\/)/i.test(v) ? '' : 'Start with https:// or /'; } }));
      default: return PP.PropertyPaneTextField(p.key, base);
    }
  }
  function errorBox(e, where, editMode) {
    var msg = (e && e.message) || String(e);
    return h('div', { class: 'sss-error', role: 'alert' },
      h('b', null, SPEC.title + ' could not ' + (where === 'render' ? 'load' : where) + '. '),
      editMode || /401|403|not found|permission/i.test(msg) ? msg : 'Refresh the page to try again. If it keeps happening, tell the page owner.',
      editMode && e && e.stack ? h('pre', null, String(e.stack).split('\n').slice(0, 4).join('\n')) : null);
  }
  function report(e) { try { if (window.__sssReport) window.__sssReport(e); else if (window.console) console.error('[' + SPEC.title + ']', e); } catch (x) { /* ignore */ } }

  return StudioWebPart;
}

export const RUNTIME_SOURCE = spfxRuntime.toString();
