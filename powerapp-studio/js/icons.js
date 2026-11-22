// icons.js — a small original icon set (24×24, stroked). Used in the designer and
// emitted into the app as SVG data URIs inside Image controls (no media upload needed).

export const ICONS = {
  home: 'M3 11l9-7 9 7M5 10v10h5v-6h4v6h5V10',
  back: 'M15 5l-7 7 7 7M8 12h12',
  forward: 'M9 5l7 7-7 7M16 12H4',
  chevron: 'M9 6l6 6-6 6',
  plus: 'M12 5v14M5 12h14',
  edit: 'M4 20h4l11-11-4-4L4 16v4zM13 7l4 4',
  trash: 'M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13M10 11v6M14 11v6',
  search: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM16.5 16.5L21 21',
  check: 'M5 12.5l4.5 4.5L19 7',
  close: 'M6 6l12 12M18 6L6 18',
  user: 'M12 4a4 4 0 1 0 0 8 4 4 0 0 0 0-8zM4 21c1-4 4.5-6 8-6s7 2 8 6',
  users: 'M9 5a3.5 3.5 0 1 0 0 7 3.5 3.5 0 0 0 0-7zM2.5 20c.8-3.5 3.6-5.5 6.5-5.5s5.7 2 6.5 5.5M16 5.5a3 3 0 0 1 0 6M18 14.5c1.8.6 3.1 2.4 3.5 5.5',
  calendar: 'M4 6h16v14H4zM4 10h16M8 3v5M16 3v5',
  clock: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM12 7v5l3 2',
  mail: 'M3 6h18v12H3zM3 7l9 6 9-6',
  bell: 'M6 16V11a6 6 0 0 1 12 0v5l2 2H4l2-2zM10 20a2 2 0 0 0 4 0',
  list: 'M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01',
  grid: 'M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z',
  chart: 'M4 20V10M10 20V4M16 20v-7M22 20H2',
  filter: 'M4 5h16l-6 7v6l-4 2v-8L4 5z',
  refresh: 'M20 11a8 8 0 0 0-14.5-4.5L4 8M4 4v4h4M4 13a8 8 0 0 0 14.5 4.5L20 16M20 20v-4h-4',
  download: 'M12 4v11M7 10l5 5 5-5M5 20h14',
  upload: 'M12 20V9M7 14l5-5 5 5M5 4h14',
  star: 'M12 3.5l2.6 5.5 6 .8-4.4 4.1 1.1 6L12 17l-5.3 2.9 1.1-6L3.4 9.8l6-.8L12 3.5z',
  info: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM12 11v6M12 7.5h.01',
  alert: 'M12 3l10 18H2L12 3zM12 10v5M12 18h.01',
  lock: 'M6 11h12v10H6zM8.5 11V8a3.5 3.5 0 0 1 7 0v3',
  file: 'M6 3h8l4 4v14H6zM14 3v4h4M9 13h6M9 17h6',
  folder: 'M3 6h6l2 2h10v11H3z',
  menu: 'M4 6h16M4 12h16M4 18h16',
  settings: 'M4 7h10M18 7h2M4 17h2M10 17h10M14 5v4M8 15v4',
  logout: 'M15 4h4v16h-4M10 8l-4 4 4 4M6 12h10',
  link: 'M10 14a4 4 0 0 0 5.6 0l3-3a4 4 0 0 0-5.6-5.6l-1 1M14 10a4 4 0 0 0-5.6 0l-3 3a4 4 0 0 0 5.6 5.6l1-1',
  tag: 'M3 12V4h8l10 10-8 8L3 12zM7.5 8h.01',
  box: 'M3 7l9-4 9 4v10l-9 4-9-4V7zM3 7l9 4 9-4M12 11v10',
  money: 'M3 7h18v10H3zM12 9.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5zM6 10v4M18 10v4',
  flag: 'M5 21V4M5 4h11l-2 4 2 4H5',
  pin: 'M12 21s-7-6.5-7-12a7 7 0 0 1 14 0c0 5.5-7 12-7 12zM12 6.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z',
  send: 'M21 3L3 10.5l7 2.5 2.5 7L21 3zM10 13l5-5',
  eye: 'M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12zM12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6z',
};
export const ICON_NAMES = Object.keys(ICONS);

/** SVG markup with single-quoted attributes (safe inside a Power Fx string). */
export const iconSvg = (name, color, stroke = 2) =>
  `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'><path d='${ICONS[name] || ICONS.info}' fill='none' stroke='${color}' stroke-width='${stroke}' stroke-linecap='round' stroke-linejoin='round'/></svg>`;
