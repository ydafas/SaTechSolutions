// share.js — read-only "Play" preview links. The whole project rides in the URL fragment (#play=…), which
// browsers never send to a server: compressed with deflate-raw when the browser has CompressionStream,
// plain JSON otherwise, then base64url. What goes in is a trimmed copy: no API keys, no flow test samples,
// no setup drafts, sample rows capped, group IDs and fixed email addresses replaced, and the SharePoint site
// URL and app link only when the person ticks the box.
import { clone, normalizeProject } from './core.js';

export const PREFIX = '#play=';
export const MAX_ROWS = 25;          // imported sample rows kept per list
export const WARN_CHARS = 16000;     // some mail and chat apps cut links around here
export const MAX_CHARS = 100000;     // refuse: browsers and servers start dropping fragments this long
const MAX_JSON = 6000000;            // decompressed size we are willing to parse
const ZERO_GUID = '00000000-0000-0000-0000-000000000000';
const HIDDEN = 'address hidden in preview';

/** The trimmed copy that goes into a link. */
export function shareProject(project, { includeLinks = false } = {}) {
  const p = clone(project);
  delete p.deployed;
  p.app = p.app || {};
  for (const k of ['setup', 'flowSample', 'approvalsDraft']) delete p.app[k];
  if (!includeLinks) { delete p.app.siteUrl; delete p.app.appLink; }
  if (p.app.features?.adminGroup) p.app.features.adminGroup = ZERO_GUID;
  for (const st of p.app.chain?.stages || []) { if (st.groupId) st.groupId = ZERO_GUID; if (st.notify) st.notify = HIDDEN; }
  for (const f of p.flows || []) {
    delete f.sample;
    for (const r of f.recipients || []) if (r.kind === 'address') r.value = HIDDEN;
  }
  for (const l of p.lists || []) if (Array.isArray(l.rows)) l.rows = l.rows.slice(0, MAX_ROWS);
  // belt and braces: nothing that looks like an Anthropic key, wherever it hides
  return JSON.parse(JSON.stringify(p).replace(/sk-ant-[A-Za-z0-9_-]{8,}/g, ''));
}

// ------------------------------------------------------------------ bytes ↔ base64url
const toB64u = (bytes) => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); };
const fromB64u = (t) => { const s = atob(t.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (t.length % 4)) % 4)); const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i); return b; };
async function pipe(bytes, stream) {
  const out = new Response(new Blob([bytes]).stream().pipeThrough(stream)).body.getReader();
  const parts = []; let n = 0;
  for (;;) { const { done, value } = await out.read(); if (done) break; n += value.length; if (n > MAX_JSON) throw new Error('This preview link is too large to open.'); parts.push(value); }
  const all = new Uint8Array(n); let o = 0; for (const p of parts) { all.set(p, o); o += p.length; }
  return all;
}

/** Build the link fragment. Returns {hash, chars, compressed, warn, tooLong}. */
export async function encodeShare(project, { includeLinks = false, compress = typeof CompressionStream !== 'undefined' } = {}) {
  const json = JSON.stringify({ v: 1, p: shareProject(project, { includeLinks }) });
  const raw = new TextEncoder().encode(json);
  let body, compressed = false;
  if (compress) { try { body = 'z' + toB64u(await pipe(raw, new CompressionStream('deflate-raw'))); compressed = true; } catch { body = null; } }
  if (!body) body = 'j' + toB64u(raw);
  const hash = PREFIX + body;
  return { hash, chars: hash.length, compressed, warn: hash.length > WARN_CHARS, tooLong: hash.length > MAX_CHARS };
}

export const isShareHash = (h) => typeof h === 'string' && h.startsWith(PREFIX);

/** Read a #play=… fragment back into a (normalized, untrusted-safe) project. Throws with a plain message. */
export async function decodeShare(hash) {
  if (!isShareHash(hash)) throw new Error('This is not a preview link.');
  const body = hash.slice(PREFIX.length);
  if (body.length > MAX_CHARS * 1.1) throw new Error('This preview link is too long.');
  let bytes;
  try { bytes = fromB64u(body.slice(1)); } catch { throw new Error('This preview link is damaged (it may have been cut off when it was pasted).'); }
  if (body[0] === 'z') {
    if (typeof DecompressionStream === 'undefined') throw new Error('This browser cannot open compressed preview links. Try a current Edge or Chrome.');
    try { bytes = await pipe(bytes, new DecompressionStream('deflate-raw')); } catch (e) { throw new Error(/too large/.test(e.message) ? e.message : 'This preview link is damaged (it may have been cut off when it was pasted).'); }
  } else if (body[0] !== 'j') throw new Error('This preview link is from a newer version of the Studio.');
  if (bytes.length > MAX_JSON) throw new Error('This preview link is too large to open.');
  let d;
  try { d = JSON.parse(new TextDecoder().decode(bytes)); } catch { throw new Error('This preview link is damaged (it may have been cut off when it was pasted).'); }
  if (!d || d.v !== 1 || !d.p || typeof d.p !== 'object') throw new Error('This preview link is from a newer version of the Studio.');
  const p = normalizeProject(d.p);
  for (const l of p.lists) if (Array.isArray(l.rows)) l.rows = l.rows.slice(0, MAX_ROWS);
  return p;
}
