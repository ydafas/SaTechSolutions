// setup.js — "Finish setup": the project-level values an app needs before it deploys cleanly
// (site URL, app link, email-link screen, admin group, AppUsage list). Each item maps to the
// lint codes it clears, so the checklist can sit next to the errors and fix them in place.
import { walk } from './core.js';
import { deepLinkInfo } from './flows.js';

export const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const isHttps = (v) => /^https:\/\/[^\s/]+\.[^\s]+$/i.test(String(v || '').trim());
const enabledFlows = (p) => (p.flows || []).filter((f) => f.enabled !== false);
const usesLink = (p) => enabledFlows(p).some((f) => /\{Link\}/.test(`${f.subject} ${f.body}`));
function usesAdminGate(p) {
  for (const s of p.screens) for (const [c] of walk(s.controls)) if (/\bvarIsAdmin\b/.test(Object.values(c.fx || {}).join(' '))) return true;
  return false;
}

/**
 * Items that apply to this project. Each: {id, label, help, kind ('url'|'deeplink'|'guid'|'confirm'),
 * value, done, codes (lint codes it clears), hint (why it is not done yet)}.
 */
export function setupItems(p) {
  const items = [];
  const flows = enabledFlows(p);
  if (flows.length) {
    const v = p.app.siteUrl || '';
    items.push({ id: 'siteUrl', kind: 'url', label: 'SharePoint site URL', value: v, done: isHttps(v), codes: ['PS045'],
      help: 'The site that holds your lists, e.g. https://contoso.sharepoint.com/sites/Team. Every flow reads and watches lists there.',
      hint: v && !isHttps(v) ? 'Use the full https:// address of the site.' : '' });
  }
  if (usesLink(p)) {
    const v = p.app.appLink || '';
    items.push({ id: 'appLink', kind: 'url', label: 'App web link', value: v, done: isHttps(v), codes: ['PS047'],
      help: 'After the first save in Studio: Apps → your app → Details → Web link. Emails use it for {Link}.',
      hint: v && !isHttps(v) ? 'Paste the full https:// web link.' : '' });
    const dl = deepLinkInfo(p);
    items.push({ id: 'deepLink', kind: 'deeplink', label: 'Screen that opens an item from an email link', value: p.app.deepLink || { screen: '', list: '' }, done: !!dl, codes: ['PS047'],
      help: 'Links ending in &ID=12 open this screen with item 12 selected. Usually the details screen.' });
  }
  if (usesAdminGate(p) || p.app.features?.adminGroup) {
    const v = p.app.features?.adminGroup || p.app.setup?.adminDraft || '';
    items.push({ id: 'adminGroup', kind: 'guid', label: 'Admin group ID', value: v, done: GUID.test(v), codes: ['PS041'],
      help: 'Object ID of the Microsoft 365 group whose members see approve, reject and delete (Entra admin center → Groups → the group → Object ID).',
      hint: v && !GUID.test(v) ? 'It should look like 00000000-0000-0000-0000-000000000000.' : '' });
  }
  for (const [i, st] of (p.app.chain?.stages || []).entries()) {
    if (st.approver !== 'group') continue;
    const v = st.groupId || '';
    items.push({ id: 'chainGroup:' + i, kind: 'guid', label: `${st.name} approver group ID`, value: v, done: GUID.test(v), codes: ['PS054'],
      help: `Object ID of the Microsoft 365 group whose members approve the ${st.name} stage (Entra admin center → Groups → the group → Object ID). Members are checked with Office 365 Groups when the app starts.`,
      hint: v && !GUID.test(v) ? 'It should look like 00000000-0000-0000-0000-000000000000.' : '' });
  }
  if (p.app.features?.usage) {
    const v = !!p.app.setup?.usageConfirmed;
    items.push({ id: 'usageList', kind: 'confirm', label: 'AppUsage list exists', value: v, done: v, codes: [],
      help: 'Usage tracking writes one row per user to a SharePoint list called AppUsage (the provisioning script creates it). Give users Contribute on it.' });
  }
  return items;
}
export const openSetupItems = (p) => setupItems(p).filter((i) => !i.done);

/** Apply one checklist field to the project. Returns true when the value was accepted. */
export function applySetup(p, id, value) {
  p.app.setup = p.app.setup && typeof p.app.setup === 'object' ? p.app.setup : {};
  const cg = /^chainGroup:(\d)$/.exec(id);
  if (cg) {
    const st = p.app.chain?.stages?.[+cg[1]];
    if (!st) return false;
    st.groupId = String(value || '').trim().slice(0, 40);
    return !st.groupId || GUID.test(st.groupId);
  }
  switch (id) {
    case 'siteUrl': p.app.siteUrl = String(value || '').trim(); return true;
    case 'appLink': p.app.appLink = String(value || '').trim(); return true;
    case 'deepLink': {
      const d = { screen: '', list: '', ...(p.app.deepLink || {}), ...(value || {}) };
      if (value?.screen && !value.list) {
        const s = p.screens.find((x) => x.id === value.screen);
        const b = s && [...walk(s.controls)].map(([c]) => c.bind?.list || c.field?.list).find(Boolean);
        if (b) d.list = b;
      }
      if (d.screen || d.list) p.app.deepLink = d; else delete p.app.deepLink;
      return true;
    }
    case 'adminGroup': {
      const v = String(value || '').trim();
      p.app.features = { ...(p.app.features || {}) };
      if (GUID.test(v)) { p.app.features.adminGroup = v; delete p.app.setup.adminDraft; return true; }
      if (!v) { p.app.features.adminGroup = ''; delete p.app.setup.adminDraft; return true; }
      p.app.setup.adminDraft = v; return false;
    }
    case 'usageList': p.app.setup.usageConfirmed = !!value; return true;
  }
  return false;
}
/** Lint codes the checklist can fix, for "Fix in Finish setup" links in the lint report. */
export const SETUP_CODES = new Set(['PS041', 'PS045', 'PS047', 'PS054']);
