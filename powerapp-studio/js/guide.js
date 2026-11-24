// guide.js — "Export guide": one self-contained HTML page for the customer, like the hand-made
// WAL Requests Guide. Lifecycle diagram, screens (thumbnails rendered by the studio), data model,
// who-can-do-what, automations, and numbered test cases with Pass/Fail, notes and Copy results.
// No external requests: system fonts, inline SVG, inline CSS and script.
import { walk, COLTYPES, FORM_FACTORS } from './core.js';
import { FLOW_KINDS, recipientsText, deepLinkInfo } from './flows.js';
import { statusCol } from './templates.js';

const esc = (t) => String(t ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pretty = (n) => String(n || '').replace(/^scr/, '').replace(/([a-z])([A-Z])/g, '$1 $2');
const single = (n) => pretty(n).replace(/ies$/, 'y').replace(/s$/, '').toLowerCase();
const adminOnly = (c) => /\bvarIsAdmin\b/.test(String(c.fx?.Visible || ''));
const btnLabel = (c) => (c.props?.text || c.fx?.Text || c.name).toString().replace(/^If\(.*?"([^"]+)".*$/, '$1');

/** Everything the app lets people do, read from the controls' actions. */
export function appActions(project) {
  const out = [];
  for (const scr of project.screens) for (const [c] of walk(scr.controls)) {
    for (const [ev, acts] of Object.entries(c.actions || {})) for (const a of acts || []) {
      if (['save', 'delete', 'setField', 'runFlow', 'newItem'].includes(a.type)) out.push({ screen: scr, ctl: c, ev, a, admin: adminOnly(c) });
    }
  }
  return out;
}

/** Numbered test cases. area: Create, Edit, Validation, Decision, Delete, Button flow, Flow, Link, Start, Access, Filter. */
export function guideCases(project) {
  const cases = [];
  const add = (area, title, steps, expect) => cases.push({ area, title, steps, expect });
  const acts = appActions(project);
  const screenName = (s) => pretty(s.name);
  if (project.app.features?.loading && project.screens.some((s) => /loading/i.test(s.name))) {
    add('Start', 'The app opens on the loading screen', ['Open the app from its web link.'], 'The loading screen shows its steps, then moves to the home screen on its own within a few seconds.');
  }
  const saveScreens = new Map();
  for (const x of acts.filter((x) => x.a.type === 'save')) if (!saveScreens.has(x.screen.id + x.a.list)) saveScreens.set(x.screen.id + x.a.list, x);
  for (const x of saveScreens.values()) {
    const l = project.lists.find((y) => y.name === x.a.list); if (!l) continue;
    const req = l.columns.filter((c) => c.required).map((c) => pretty(c.name));
    add('Create', `Add a new ${single(l.name)}`, [`Open ${screenName(x.screen)} with a "New" button.`, 'Fill in every field.', `Press ${btnLabel(x.ctl)}.`],
      `A success message appears${project.app.features?.confirm ? ' (or the confirmation page)' : ''}, and the new ${single(l.name)} appears at the top of the list in SharePoint and in the app.`);
    if (req.length) add('Validation', `Required fields block saving (${single(l.name)})`, [`Start a new ${single(l.name)}.`, `Leave ${req[0]} empty.`, `Try to press ${btnLabel(x.ctl)}.`],
      `${btnLabel(x.ctl)} stays disabled and the missing field is pointed out. Required: ${req.join(', ')}.`);
    add('Edit', `Edit an existing ${single(l.name)}`, [`Open any ${single(l.name)} and choose Edit.`, 'Change one field.', `Press ${btnLabel(x.ctl)}.`], 'The change is saved; the details show the new value and SharePoint has one updated item, not a new one.');
  }
  const decisions = new Map();
  for (const x of acts.filter((x) => x.a.type === 'setField')) {
    const k = `${x.a.list}|${x.a.field}|${x.a.value}`;
    if (!decisions.has(k)) decisions.set(k, x);
  }
  for (const x of decisions.values()) {
    add('Decision', `${btnLabel(x.ctl)}: set ${pretty(x.a.field)} to ${x.a.value || 'today / me'}`,
      [`${x.admin ? 'As a member of the admin group, open' : 'Open'} a ${single(x.a.list)} on ${screenName(x.screen)}.`, `Press ${btnLabel(x.ctl)}.`],
      `${pretty(x.a.field)} becomes ${x.a.value || 'the new value'} in the app and in SharePoint${x.admin ? '. People outside the admin group do not see this button' : ''}.`);
  }
  const deletes = new Map();
  for (const x of acts.filter((x) => x.a.type === 'delete')) if (!deletes.has(x.a.list)) deletes.set(x.a.list, x);
  for (const x of deletes.values()) {
    const twoClick = /confirm/i.test(JSON.stringify(x.ctl.fx || {}));
    add('Delete', `Delete a ${single(x.a.list)}`, [`Open a test ${single(x.a.list)} on ${screenName(x.screen)}.`, `Press ${btnLabel(x.ctl)}${twoClick ? ' once, then press it again when it says "Confirm delete"' : ''}.`],
      `${twoClick ? 'Nothing is deleted on the first press. ' : ''}After confirming, the item disappears from the app and from the SharePoint list.`);
  }
  for (const x of acts.filter((x) => x.a.type === 'runFlow')) {
    const f = (project.flows || []).find((y) => y.id === x.a.flow); if (!f) continue;
    add('Button flow', `Run "${f.name}" from the app`, [`Open a ${single(f.list)} on ${screenName(x.screen)}.`, `Press ${btnLabel(x.ctl)} and wait.`], `The app shows "${x.a.message || 'Done'}" and the email goes to ${recipientsText(f)}.`);
  }
  for (const f of (project.flows || []).filter((y) => y.enabled !== false)) {
    const when = { created: `Add a new ${single(f.list)}.`, statusChanged: `Change ${pretty(f.statusField)} of a ${single(f.list)} to ${(f.statusValues || []).join(' or ') || 'another value'}.`,
      reminder: `Leave a ${single(f.list)}${f.statusField ? ` with ${pretty(f.statusField)} = ${f.reminderValue}` : ''} older than ${f.days} day${f.days === 1 ? '' : 's'}, and wait for ${(f.hour % 12) || 12}:00 ${f.hour < 12 ? 'AM' : 'PM'}.`,
      button: `Run it from its button in the app.` }[f.kind];
    add('Flow', `Flow "${f.name}" sends its email`, [when, 'Wait about a minute, then check the flow\'s run history in Power Automate.'], `One email to ${recipientsText(f)} with the subject "${f.subject}".${/\{Link\}/.test(f.body) ? ' Its link opens the item in the app.' : ''}`);
  }
  const ch = project.app.chain;
  if (ch && project.lists.some((l) => l.name === ch.list)) {
    const one = single(ch.list), names = ch.stages.map((s) => s.name), who = (s) => (s.approver === 'group' ? `a member of the ${s.groupLabel || s.name} group` : `the person in ${pretty(s.column)}`);
    add('Approval chain', `${one} goes through ${names.join(' → ')}`, [`As a requester, add a ${one} and name a different person in each approver field.`, ...ch.stages.map((s, i) => `Sign in as ${who(s)}, open Approvals, pick the ${one} and press Approve${i ? '' : ' (a comment is optional)'}.`)],
      `After each approval the stepper on the details screen colours that stage green and moves to the next; the ${one} is in only the next approver's queue. After ${names[names.length - 1]}, ${pretty(ch.field)} is ${ch.values.approved}.`);
    add('Approval chain', `Rejecting needs a comment and ends the chain`, [`Sign in as the stage ${names.length > 1 ? 2 : 1} approver of a waiting ${one}.`, 'Press Reject with the comment empty, then type a comment and press Reject.'],
      `Reject stays disabled until there is a comment. Then ${pretty(ch.field)} is ${ch.values.rejected}, the stage turns red, later stages stay grey and the history shows who rejected, when, and the comment. The requester gets the outcome email.`);
    add('Approval chain', 'Nobody approves their own request', ['Sign in as someone who is an approver, and create a ' + one + ' naming yourself as the stage 1 approver.', 'Open Approvals and the details screen.'], `The ${one} is not in your queue and the details show "Waiting on…" instead of Approve and Reject.`);
    if (ch.resubmit) add('Approval chain', 'A rejected request can be resubmitted', [`As the requester, open your rejected ${one}.`, 'Edit it if needed, then press Resubmit.'], `Every stage is Pending again, ${pretty(ch.field)} is ${ch.values.pending}, the history keeps the earlier decisions with a "Resubmitted" line on top, and the stage 1 approver gets a new email.`);
  }
  const dl = deepLinkInfo(project);
  if (dl) add('Link', 'An email link opens the right item', ['Open a link from one of the app\'s emails (it ends in &ID= and a number).'], `The app opens ${screenName(dl.screen)} with that ${single(dl.list)} selected, not the home screen.`);
  if (project.app.features?.adminGroup) add('Access', 'Admin-only buttons stay hidden for others', ['Sign in as someone who is not in the admin group.', 'Open an item.'], 'Approve, reject and delete are not shown. Members of the admin group see them.');
  let mine = false;
  for (const s of project.screens) for (const [c] of walk(s.controls)) if (/CreatedByMe|'Created By'\.Email/.test(JSON.stringify(c.data || {}) + JSON.stringify(c.fx || {}))) mine = true;
  if (mine) add('Filter', '"Mine" shows only my items', ['Create one item as yourself; have a colleague create another.', 'Switch the list to Mine.'], 'Only your item is listed; All shows both.');
  return cases.map((c, i) => ({ id: 'TC' + String(i + 1).padStart(2, '0'), ...c }));
}

/** Lifecycle of the main Choice column, with transitions from the app's buttons and the flows that email. */
export function lifecycleSvg(project, theme) {
  const list = project.lists.find((l) => statusCol(l)) || null;
  const col = list && statusCol(list);
  if (!col || !col.choices.length) return '';
  const C = theme.colors;
  const n = col.choices.length, W = 170, G = 50, H = 64, top = 96, width = n * W + (n - 1) * G + 160;
  const x = (i) => 140 + i * (W + G);
  const trans = [];
  for (const a of appActions(project)) if (a.a.type === 'setField' && a.a.list === list.name && a.a.field === col.name) {
    const to = col.choices.indexOf(a.a.value); if (to < 0) continue;
    const from = to === 0 ? -1 : 0;
    if (!trans.some((t) => t.to === to && t.from === from)) trans.push({ from, to, label: btnLabel(a.ctl) });
  }
  if (!trans.length) for (let i = 1; i < n; i++) trans.push({ from: i - 1, to: i, label: '' });
  const mails = (v) => (project.flows || []).filter((f) => f.enabled !== false && f.list === list.name && f.kind === 'statusChanged' && f.statusField === col.name && (!(f.statusValues || []).length || f.statusValues.includes(v)));
  const created = (project.flows || []).filter((f) => f.enabled !== false && f.list === list.name && f.kind === 'created');
  let svg = `<svg viewBox="0 0 ${width} 260" role="img" aria-label="Lifecycle of ${esc(pretty(col.name))}" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:${width}px;height:auto">
    <defs><marker id="ar" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${C.TextMuted}"/></marker></defs>
    <rect x="10" y="${top + 10}" width="96" height="${H - 20}" rx="22" fill="${C.SurfaceAlt}" stroke="${C.Border}"/><text x="58" y="${top + H / 2 + 5}" text-anchor="middle" font-size="13" fill="${C.Text}">New item</text>
    <path d="M108 ${top + H / 2} L${x(0) - 4} ${top + H / 2}" stroke="${C.TextMuted}" stroke-width="2" fill="none" marker-end="url(#ar)"/>`;
  if (created.length) svg += `<text x="${(108 + x(0)) / 2}" y="${top + H / 2 + 24}" text-anchor="middle" font-size="11" fill="${C.Primary}">✉ ${esc(created.map((f) => f.name).join(', ')).slice(0, 40)}</text>`;
  col.choices.forEach((v, i) => {
    const m = mails(v);
    svg += `<rect x="${x(i)}" y="${top}" width="${W}" height="${H}" rx="12" fill="${i === 0 ? C.Primary : C.Surface}" stroke="${i === 0 ? C.Primary : C.Border}" stroke-width="2"/>
      <text x="${x(i) + W / 2}" y="${top + H / 2 + 5}" text-anchor="middle" font-size="15" font-weight="700" fill="${i === 0 ? C.OnPrimary : C.Text}">${esc(v)}</text>`;
    if (m.length) svg += `<text x="${x(i) + W / 2}" y="${top + H + 24}" text-anchor="middle" font-size="11" fill="${C.Primary}">✉ ${esc(m.map((f) => f.name).join(', ')).slice(0, 44)}</text>`;
  });
  trans.forEach((t, k) => {
    if (t.from < 0) return;
    const x1 = x(t.from) + W / 2, x2 = x(t.to) + W / 2, lift = 34 + Math.abs(t.to - t.from) * 18 + k * 6, y = top;
    svg += `<path d="M${x1} ${y} C ${x1} ${y - lift}, ${x2} ${y - lift}, ${x2} ${y - 4}" stroke="${C.TextMuted}" stroke-width="2" fill="none" marker-end="url(#ar)"/>`;
    if (t.label) svg += `<text x="${(x1 + x2) / 2}" y="${y - lift * 0.78}" text-anchor="middle" font-size="12" font-weight="600" fill="${C.Text}">${esc(t.label)}</text>`;
  });
  return svg + '</svg>';
}

function rolesRows(project) {
  const rows = [], seen = new Set();
  const gated = !!project.app.features?.adminGroup;
  const push = (what, admin) => { if (seen.has(what)) return; seen.add(what); rows.push({ what, everyone: !admin || !gated, admin: true }); };
  for (const l of project.lists) push(`See ${pretty(l.name).toLowerCase()}`, false);
  for (const x of appActions(project)) {
    const l = single(x.a.list || (project.flows || []).find((f) => f.id === x.a.flow)?.list || '');
    if (x.a.type === 'newItem' || x.a.type === 'save') push(`Add and edit a ${l}`, x.admin);
    if (x.a.type === 'delete') push(`Delete a ${l}`, x.admin);
    if (x.a.type === 'setField') push(`${btnLabel(x.ctl)} a ${l}`, x.admin);
    if (x.a.type === 'runFlow') push(`${btnLabel(x.ctl)}`, x.admin);
  }
  return { rows, gated };
}

/** The whole guide page. thumbs: {screenId: html of the rendered screen}; css: rules the thumbnails need. */
export function guideHtml(project, { thumbs = {}, css = '', date = new Date() } = {}) {
  const t = project.theme, C = t.colors;
  const ff = FORM_FACTORS[project.app.formFactor || 'tablet'];
  const cases = guideCases(project);
  const flows = (project.flows || []).filter((f) => f.enabled !== false);
  const { rows, gated } = rolesRows(project);
  const lc = lifecycleSvg(project, t);
  const stamp = date.toISOString().slice(0, 10);
  const storeKey = 'pas-guide:' + (project.app.name || 'app') + ':' + stamp;
  const tw = 360, scale = tw / ff.w;
  const scrDesc = (s) => {
    let g = 0, inp = 0, b = 0; const targets = new Set();
    for (const [c] of walk(s.controls)) {
      if (c.type === 'gallery') g++;
      if (['textinput', 'dropdown', 'datepicker', 'checkbox', 'toggle', 'person'].includes(c.type)) inp++;
      if (c.type === 'button' && (c.actions?.OnSelect || []).length) b++;
      for (const a of Object.values(c.actions || {}).flat()) if (a.screen) { const tt = project.screens.find((x) => x.id === a.screen); if (tt && tt !== s) targets.add(pretty(tt.name)); }
    }
    return [g ? `${g} list${g > 1 ? 's' : ''}` : '', inp ? `${inp} input${inp > 1 ? 's' : ''}` : '', b ? `${b} button${b > 1 ? 's' : ''}` : '', targets.size ? 'goes to ' + [...targets].join(', ') : ''].filter(Boolean).join(' · ');
  };
  const clean = (h) => String(h || '').replace(/\s(src|href)\s*=\s*(["'])(?:https?:)?\/\/[^"']*\2/gi, '').replace(/url\(\s*["']?(?:https?:)?\/\/[^)]*\)/gi, 'none').replace(/<script[\s\S]*?<\/script>/gi, '');
  const kindWhen = (f) => ({ created: `New ${single(f.list)} added`, statusChanged: `${pretty(f.statusField)} changes${(f.statusValues || []).length ? ' to ' + f.statusValues.join(' or ') : ''}`, reminder: `Daily at ${(f.hour % 12) || 12}:00 ${f.hour < 12 ? 'AM' : 'PM'} for items${f.statusField ? ` with ${pretty(f.statusField)} = ${f.reminderValue}` : ''} over ${f.days} day${f.days === 1 ? '' : 's'}`, button: 'A button in the app' }[f.kind]);
  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(project.app.name)} guide</title>
<style>
:root{--p:${C.Primary};--on:${C.OnPrimary};--ink:${C.Text};--mut:${C.TextMuted};--bd:${C.Border};--bg:#f6f8fb;--card:#fff;--hd:${C.HeaderBg};--ok:${C.Success};--bad:${C.Danger}}
*{box-sizing:border-box;min-width:0}body{margin:0;background:var(--bg);color:#0f172a;font:16px/1.55 'Segoe UI',system-ui,-apple-system,Roboto,Arial,sans-serif}
header{background:linear-gradient(110deg,var(--hd),var(--p));color:#fff;padding:44px max(5vw,16px) 36px}
header h1{margin:0 0 6px;font-size:clamp(30px,4vw,44px);line-height:1.08;letter-spacing:-.01em}header p{margin:0;opacity:.85;max-width:70ch}
nav{position:sticky;top:0;z-index:5;background:#fff;border-bottom:1px solid #e2e8f0;display:flex;gap:4px;overflow-x:auto;padding:8px max(5vw,16px)}
nav a{color:#334155;text-decoration:none;font-weight:600;font-size:14px;padding:6px 10px;border-radius:7px;white-space:nowrap}nav a:hover{background:#eef2f7}
main{max-width:1180px;margin:0 auto;padding:12px max(4vw,16px) 64px}section{margin-top:40px}h2{font-size:24px;margin:0 0 6px}h2+p{color:#475569;margin:0 0 16px;max-width:75ch}
.card{background:var(--card);border:1px solid #e2e8f0;border-radius:12px;padding:16px;margin-bottom:12px}
.fig{overflow-x:auto;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:10px}
table{width:100%;border-collapse:collapse;font-size:14px;background:#fff}th,td{text-align:left;padding:9px 10px;border-bottom:1px solid #e2e8f0;vertical-align:top}th{font-size:12px;color:#475569;background:#f8fafc}
td.c{text-align:center;font-weight:700}.yes{color:var(--ok)}.no{color:#94a3b8}
.scr{display:grid;grid-template-columns:repeat(auto-fill,minmax(${tw + 20}px,1fr));gap:16px}
.thumb{width:${tw}px;height:${Math.round(ff.h * scale)}px;overflow:hidden;border-radius:8px;border:1px solid #e2e8f0;position:relative;background:#fff;max-width:100%}
.thumb>.screen{position:absolute;left:0;top:0;width:${ff.w}px;height:${ff.h}px;transform:scale(${scale});transform-origin:0 0;pointer-events:none;overflow:hidden}
.scr h3{margin:10px 0 2px;font-size:16px}.scr small{color:#64748b}
.tc{display:grid;grid-template-columns:64px minmax(0,1fr) auto;gap:14px;align-items:start}
.tc .id{font:700 13px ui-monospace,Consolas,monospace;color:var(--p);padding-top:3px}.tc h3{margin:0 0 4px;font-size:16px}.tc ol{margin:4px 0 6px;padding-left:20px;color:#334155}.tc .ex{font-size:14px;color:#0f172a}.tc .ex b{color:#475569;font-weight:600}
.area{display:inline-block;font-size:11px;font-weight:700;color:#475569;background:#eef2f7;border-radius:99px;padding:1px 8px;margin-left:6px;vertical-align:2px}
.pf{display:flex;gap:6px}.pf button{min-height:40px;min-width:68px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;font-weight:700;cursor:pointer}
.pf button[aria-pressed=true].pass{background:var(--ok);border-color:var(--ok);color:#fff}.pf button[aria-pressed=true].fail{background:var(--bad);border-color:var(--bad);color:#fff}
.tc textarea{grid-column:2/4;width:100%;min-height:40px;border:1px solid #e2e8f0;border-radius:8px;padding:6px 8px;font:inherit;font-size:14px;resize:vertical}
.bar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:12px}.bar .sum{font-weight:700}.bar button{min-height:40px;border-radius:8px;border:1px solid var(--p);background:var(--p);color:var(--on);font-weight:700;padding:0 14px;cursor:pointer}.bar button.alt{background:#fff;color:var(--p)}
code{background:#eef2f7;border-radius:4px;padding:1px 5px;font-size:.9em}
:focus-visible{outline:2px solid var(--p);outline-offset:2px}
@media(max-width:640px){.tc{grid-template-columns:1fr}.tc textarea{grid-column:1}.pf{justify-content:flex-start}table{font-size:13px}}
@media print{nav,.bar button,.pf button:not([aria-pressed=true]){display:none}.card{break-inside:avoid}}
${css}
</style></head><body>
<header><h1>${esc(project.app.name)}</h1><p>How the app works, who can do what, what emails it sends, and a test plan to sign off on. Generated ${esc(stamp)}.</p></header>
<nav aria-label="Sections"><a href="#lifecycle">Lifecycle</a><a href="#screens">Screens</a><a href="#data">Data</a><a href="#roles">Who can do what</a>${flows.length ? '<a href="#flows">Emails and flows</a>' : ''}<a href="#tests">Test plan</a></nav>
<main>
${lc ? `<section id="lifecycle"><h2>Lifecycle</h2><p>Every item moves through these states. Arrows show the buttons that move it; ✉ marks the emails sent along the way.</p><div class="fig">${lc}</div></section>` : ''}
<section id="screens"><h2>Screens</h2><p>${project.screens.length} screens, in the order people meet them.</p><div class="scr">${project.screens.map((s) => `<div><div class="thumb">${thumbs[s.id] ? `<div class="screen">${clean(thumbs[s.id])}</div>` : ''}</div><h3>${esc(pretty(s.name))}</h3><small>${esc(scrDesc(s))}</small></div>`).join('')}</div></section>
<section id="data"><h2>Data</h2><p>SharePoint lists the app reads and writes. Starred columns are required.</p>${project.lists.map((l) => `<div class="card"><h3 style="margin:0 0 8px">${esc(pretty(l.name))}</h3><table><thead><tr><th>Column</th><th>Type</th><th>Choices / notes</th></tr></thead><tbody>${l.columns.map((c) => `<tr><td>${esc(c.name)}${c.required ? ' *' : ''}</td><td>${esc(COLTYPES[c.type]?.label || c.type)}</td><td>${esc(c.type === 'Choice' ? c.choices.join(', ') : c.type === 'Lookup' ? 'Points to ' + c.lookupList : c.builtin ? 'Built-in' : '')}</td></tr>`).join('')}</tbody></table></div>`).join('')}</section>
<section id="roles"><h2>Who can do what</h2><p>${gated ? 'Some actions are limited to members of the admin group (Microsoft 365 group <code>' + esc(project.app.features.adminGroup) + '</code>). Real protection comes from SharePoint list permissions; the app hides the buttons.' : 'Everyone who can open the app and the lists can do these. Use SharePoint list permissions to restrict who can edit.'}</p>
<table><thead><tr><th>Action</th><th style="text-align:center">Everyone</th>${gated ? '<th style="text-align:center">Admin group</th>' : ''}</tr></thead><tbody>${rows.map((r) => `<tr><td>${esc(r.what)}</td><td class="c ${r.everyone ? 'yes' : 'no'}">${r.everyone ? '✓' : '–'}</td>${gated ? `<td class="c yes">✓</td>` : ''}</tr>`).join('')}</tbody></table></section>
${flows.length ? `<section id="flows"><h2>Emails and flows</h2><p>Power Automate flows that run alongside the app.</p><table><thead><tr><th>Flow</th><th>When</th><th>To</th><th>Subject</th></tr></thead><tbody>${flows.map((f) => `<tr><td><b>${esc(f.name)}</b><br><small>${esc(FLOW_KINDS[f.kind]?.label || '')}</small></td><td>${esc(kindWhen(f))}</td><td>${esc(recipientsText(f))}</td><td>${esc(f.subject)}</td></tr>`).join('')}</tbody></table></section>` : ''}
<section id="tests"><h2>Test plan</h2><p>${cases.length} test cases. Mark each one, add notes, then Copy results into an email or ticket. Marks are kept in this browser.</p>
<div class="bar"><span class="sum" id="sum"></span><button id="copy" type="button">Copy results</button><button id="reset" class="alt" type="button">Clear marks</button></div>
${cases.map((c) => `<div class="card tc" data-id="${c.id}"><span class="id">${c.id}</span><div><h3>${esc(c.title)}<span class="area">${esc(c.area)}</span></h3><ol>${c.steps.map((s) => `<li>${esc(s)}</li>`).join('')}</ol><div class="ex"><b>Expected:</b> ${esc(c.expect)}</div></div>
<div class="pf"><button type="button" class="pass" aria-pressed="false">Pass</button><button type="button" class="fail" aria-pressed="false">Fail</button></div><textarea aria-label="Notes for ${c.id}" placeholder="Notes"></textarea></div>`).join('')}
</section></main>
<script>{
const KEY=${JSON.stringify(storeKey)};let st={};
try{st=JSON.parse(localStorage.getItem(KEY)||'{}')||{}}catch(e){st={}}
const save=()=>{try{localStorage.setItem(KEY,JSON.stringify(st))}catch(e){}};
const cards=[...document.querySelectorAll('.tc')];
const paint=()=>{let p=0,f=0;for(const c of cards){const s=st[c.dataset.id]||{};c.querySelector('.pass').setAttribute('aria-pressed',String(s.r==='pass'));c.querySelector('.fail').setAttribute('aria-pressed',String(s.r==='fail'));if(s.r==='pass')p++;if(s.r==='fail')f++;}
document.getElementById('sum').textContent=p+' passed, '+f+' failed, '+(cards.length-p-f)+' not run';};
for(const c of cards){const id=c.dataset.id,t=c.querySelector('textarea');t.value=(st[id]||{}).n||'';
c.querySelector('.pass').onclick=()=>{st[id]={...(st[id]||{}),r:(st[id]||{}).r==='pass'?'':'pass'};save();paint()};
c.querySelector('.fail').onclick=()=>{st[id]={...(st[id]||{}),r:(st[id]||{}).r==='fail'?'':'fail'};save();paint()};
t.oninput=()=>{st[id]={...(st[id]||{}),n:t.value};save()};}
document.getElementById('copy').onclick=()=>{const lines=[${JSON.stringify(project.app.name + ' test results')},document.getElementById('sum').textContent,''];
for(const c of cards){const s=st[c.dataset.id]||{};lines.push(c.dataset.id+'  '+(s.r?s.r.toUpperCase():'NOT RUN')+'  '+c.querySelector('h3').firstChild.textContent+(s.n?'  ('+s.n+')':''));}
const txt=lines.join('\\n');const done=()=>{const b=document.getElementById('copy');b.textContent='Copied';setTimeout(()=>b.textContent='Copy results',1500)};
try{navigator.clipboard.writeText(txt).then(done,()=>{prompt('Copy the results:',txt)})}catch(e){prompt('Copy the results:',txt)}};
document.getElementById('reset').onclick=()=>{st={};save();for(const c of cards)c.querySelector('textarea').value='';paint()};
paint();}</script>
</body></html>`;
}
