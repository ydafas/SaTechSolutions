// keystore.js — where the visitor's own Anthropic API key and model choice live (this tab, or this browser if
// they tick Remember). Never written to project files, exports or anything sent anywhere except api.anthropic.com.
export const KEY = 'pas.aikey', DRAFT = 'pas.aidraft', MODEL = 'pas.aimodel', CHANGE_DRAFT = 'pas.aichange';
export const store = {
  get(k) { try { return localStorage.getItem(k) ?? sessionStorage.getItem(k) ?? ''; } catch { return ''; } },
  set(k, v, remember) { try { (remember ? localStorage : sessionStorage).setItem(k, v); (remember ? sessionStorage : localStorage).removeItem(k); } catch { /* storage blocked */ } },
  del(k) { try { localStorage.removeItem(k); sessionStorage.removeItem(k); } catch { /* storage blocked */ } },
  remembered(k) { try { return !!localStorage.getItem(k); } catch { return false; } },
};
