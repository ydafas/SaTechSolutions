// a11y.js — WCAG contrast checks for the theme and a one-click fix that nudges the failing color.
import { hexToRgba, isHex } from './core.js';

/** Theme color pairs the generated screens put text on: [foreground, background, where it shows]. */
export const CONTRAST_PAIRS = [
  ['Text', 'Bg', 'Body text on the page background'],
  ['Text', 'Surface', 'Text on cards and forms'],
  ['TextMuted', 'Surface', 'Captions and hints on cards'],
  ['OnPrimary', 'Primary', 'Button text'],
  ['OnHeader', 'HeaderBg', 'Header title and icons'],
];
export const AA = 4.5;

const lin = (c) => { c /= 255; return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; };
/** WCAG relative luminance of a #RRGGBB color (alpha ignored). */
export function luminance(hex) { const [r, g, b] = hexToRgba(hex); return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b); }
/** WCAG contrast ratio (1–21). */
export function contrast(a, b) { const x = luminance(a), y = luminance(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); }

/** Every pair below AA for a theme's colors. */
export function contrastIssues(colors) {
  return CONTRAST_PAIRS.filter(([f, b]) => isHex(colors[f]) && isHex(colors[b])).map(([f, b, where]) => ({ fg: f, bg: b, where, ratio: contrast(colors[f], colors[b]) })).filter((x) => x.ratio < AA);
}

// HSL helpers keep the hue while the lightness moves
function toHsl(hex) {
  const [r, g, b] = hexToRgba(hex).map((v, i) => (i < 3 ? v / 255 : v));
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min;
  if (!d) return [0, 0, l];
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h * 60, s, l];
}
function toHex([h, s, l]) {
  const f = (n) => { const k = (n + h / 30) % 12, a = s * Math.min(l, 1 - l); return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1)))); };
  return '#' + [f(0), f(8), f(4)].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}
/** Move a color's lightness (darker or lighter, away from `other`) just far enough to reach `target`. Returns the new hex or null. */
function nudge(hex, other, target) {
  const [h, s, l] = toHsl(hex), dir = luminance(hex) >= luminance(other) ? 1 : -1;
  for (let i = 1; i <= 100; i++) {
    const nl = Math.max(0, Math.min(1, l + dir * i * 0.01));
    const c = toHex([h, s, nl]);
    if (contrast(c, other) >= target) return c;
    if (nl === 0 || nl === 1) break;
  }
  return null;
}
/**
 * The smallest change that makes fg on bg pass AA: nudge the text color first, and the background only when the text
 * color cannot get there (already near white or black). Returns {key, value} for the color to change.
 */
export function fixContrast(colors, fg, bg, target = AA + 0.05) {
  const a = nudge(colors[fg], colors[bg], target);
  if (a) return { key: fg, value: a };
  const b = nudge(colors[bg], colors[fg], target);
  if (b) return { key: bg, value: b };
  return { key: fg, value: luminance(colors[bg]) > 0.18 ? '#000000' : '#FFFFFF' };
}
