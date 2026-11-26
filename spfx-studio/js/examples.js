// examples.js — hand-written starting points. They work without an API key, they are the fixtures for the tests,
// and the first one is shown to Claude as the quality bar for generated web parts.

const ICON_CHOICES = ['link', 'document', 'people', 'calendar', 'mail', 'megaphone', 'book', 'briefcase', 'globe', 'settings', 'chart', 'lightbulb', 'home', 'folder', 'video', 'trophy'];

export const EXAMPLES = [
  {
    key: 'quicklinks',
    blurb: 'Icon tiles or a compact list from a SharePoint list, with sort order and new-tab links.',
    spec: {
      title: 'Quick Links Tiles', alias: 'QuickLinksTiles', icon: 'Link', group: 'Text, media, and content',
      description: 'Icon tiles that link to the tools and pages people use most, managed in a SharePoint list.',
      properties: [
        { key: 'title', type: 'text', label: 'Heading', description: 'Leave empty to hide the heading', group: 'Display', default: 'Quick links' },
        { key: 'layout', type: 'dropdown', label: 'Layout', description: '', group: 'Display', default: 'Tiles', options: ['Tiles', 'Compact list'] },
        { key: 'columns', type: 'slider', label: 'Tiles per row', description: '', group: 'Display', default: '4', min: 2, max: 6 },
        { key: 'showDescriptions', type: 'toggle', label: 'Show descriptions', description: '', group: 'Display', default: 'true' },
        { key: 'accent', type: 'color', label: 'Accent colour', description: '', group: 'Display', default: '' },
        { key: 'listName', type: 'list', label: 'Links list', description: '', group: 'Data', default: 'QuickLinks' },
      ],
      lists: [{
        name: 'QuickLinks', description: 'Links shown by the Quick Links Tiles web part',
        columns: [
          { name: 'Title', type: 'Text', required: true, choices: [] },
          { name: 'Link', type: 'URL', required: true, choices: [] },
          { name: 'IconName', type: 'Choice', required: false, choices: ICON_CHOICES },
          { name: 'Description', type: 'Text', required: false, choices: [] },
          { name: 'SortOrder', type: 'Number', required: false, choices: [] },
          { name: 'OpenInNewTab', type: 'Boolean', required: false, choices: [] },
        ],
        sample: [
          ['Employee handbook', 'https://contoso.sharepoint.com/sites/hr/handbook', 'book', 'Policies, benefits and leave', '1', 'false'],
          ['Submit a ticket', 'https://contoso.service-now.com', 'settings', 'IT help desk requests', '2', 'true'],
          ['Org chart', 'https://contoso.sharepoint.com/sites/hr/org', 'people', 'Who reports to whom', '3', 'false'],
          ['Training calendar', 'https://contoso.sharepoint.com/sites/learning', 'calendar', 'Upcoming courses and sign-ups', '4', 'false'],
          ['Travel requests', 'https://contoso.sharepoint.com/sites/travel', 'globe', 'Book and approve travel', '5', 'true'],
          ['Dashboards', 'https://contoso.sharepoint.com/sites/bi', 'chart', 'Monthly metrics', '6', 'false'],
        ],
      }],
      css: `.ql-title{font-size:20px;font-weight:600;margin:0 0 14px;color:var(--sss-text)}
.ql-tiles{display:grid;grid-template-columns:repeat(var(--cols),minmax(0,1fr));gap:12px}
@container (max-width:640px){.ql-tiles{grid-template-columns:repeat(2,minmax(0,1fr))}}
@container (max-width:340px){.ql-tiles{grid-template-columns:1fr}}
.ql-tile{position:relative;display:flex;flex-direction:column;gap:10px;padding:16px;border:1px solid var(--sss-border);border-radius:10px;background:var(--sss-bg);color:var(--sss-text);text-decoration:none;transition:border-color .15s,box-shadow .15s,transform .15s}
.ql-tile:hover{border-color:var(--accent);box-shadow:0 6px 18px rgba(0,0,0,.08);transform:translateY(-1px)}
.ql-tile:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.ql-ic{width:40px;height:40px;border-radius:9px;display:grid;place-items:center;background:color-mix(in srgb,var(--accent) 12%,transparent);color:var(--accent)}
.ql-txt{display:flex;flex-direction:column;gap:3px;min-width:0}
.ql-txt b{font-weight:600;font-size:14px}
.ql-txt small{color:var(--sss-muted);font-size:12px;line-height:1.4}
.ql-ext{position:absolute;top:12px;right:12px;color:var(--sss-muted)}
.ql-list{display:flex;flex-direction:column;border:1px solid var(--sss-border);border-radius:10px;overflow:hidden}
.ql-list .ql-tile{flex-direction:row;align-items:center;border:0;border-radius:0;border-bottom:1px solid var(--sss-border);padding:10px 14px}
.ql-list .ql-tile:last-child{border-bottom:0}
.ql-list .ql-tile:hover{transform:none;box-shadow:none;background:var(--sss-surface)}
.ql-list .ql-ic{width:32px;height:32px}
.ql-list .ql-ext{position:static;margin-left:auto}
.ql-setup,.ql-empty{display:flex;gap:12px;align-items:center;padding:18px;border:1px dashed var(--sss-border);border-radius:10px;color:var(--sss-muted)}
.ql-setup p,.ql-empty p{margin:2px 0 0;font-size:13px}
.ql-setup b{color:var(--sss-text)}`,
      code: `const list = props.listName;
const wrap = h('section', { class: 'ql', 'aria-label': props.title || 'Quick links' });
if (props.title) wrap.appendChild(h('h2', { class: 'ql-title' }, props.title));
ctx.root.appendChild(wrap);
const setup = (msg) => h('div', { class: 'ql-setup' }, icon('settings', 22), h('div', null,
  h('b', null, 'Connect a list'),
  h('p', null, msg || (ctx.page.editMode ? 'Pick a list in the property pane, or create the QuickLinks list from its Lists section.' : 'This web part is not set up yet.'))));
if (!list) { wrap.appendChild(setup()); return; }
let items;
try {
  items = await ctx.cache('items:' + list, () => sp.items(list, {
    select: ['Id', 'Title', 'Link', 'IconName', 'Description', 'SortOrder', 'OpenInNewTab'],
    orderBy: ['SortOrder asc', 'Title asc'], top: 200 }));
} catch (e) {
  if (e.status === 404) { wrap.appendChild(setup('The list "' + list + '" is not on this site yet.' + (ctx.page.editMode ? ' Create it from the Lists section of the property pane.' : ''))); return; }
  throw e;
}
if (!items.length) {
  wrap.appendChild(h('div', { class: 'ql-empty' }, icon('link', 26), h('div', null, h('b', null, 'No links yet'),
    ctx.page.editMode ? h('p', null, 'Add items to the ' + list + ' list and they appear here.') : null)));
  return;
}
const compact = props.layout === 'Compact list';
const grid = h('div', { class: compact ? 'ql-list' : 'ql-tiles', role: 'list',
  style: { '--cols': String(props.columns || 4), '--accent': props.accent || 'var(--sss-primary)' } });
for (const it of items) {
  const url = it.Link && it.Link.Url;
  const newTab = !!it.OpenInNewTab;
  grid.appendChild(h('a', { class: 'ql-tile', role: 'listitem', href: url || '#', target: newTab ? '_blank' : null, rel: newTab ? 'noopener' : null,
      title: it.Description || it.Title, 'aria-label': it.Title + (newTab ? ' (opens in a new tab)' : '') },
    h('span', { class: 'ql-ic' }, icon(it.IconName || 'link', compact ? 18 : 22)),
    h('span', { class: 'ql-txt' }, h('b', null, it.Title), props.showDescriptions && it.Description ? h('small', null, it.Description) : null),
    newTab ? icon('external', 14, 'ql-ext') : null));
}
wrap.appendChild(grid);`,
      notes: 'Pick or create the QuickLinks list in the property pane. Icons come from the built-in icon set, so nothing loads from outside SharePoint.',
    },
  },
  {
    key: 'announcements',
    blurb: 'Pinned and recent news cards with category filters, "New" badges and read-more.',
    spec: {
      title: 'Announcements Board', alias: 'AnnouncementsBoard', icon: 'Megaphone', group: 'Communication and collaboration',
      description: 'Pinned and recent announcements from a SharePoint list, with category filters and expiry dates.',
      properties: [
        { key: 'heading', type: 'text', label: 'Heading', description: '', group: 'Display', default: 'Announcements' },
        { key: 'maxItems', type: 'slider', label: 'How many to show', description: '', group: 'Display', default: '6', min: 1, max: 20 },
        { key: 'showImages', type: 'toggle', label: 'Show images', description: '', group: 'Display', default: 'true' },
        { key: 'showFilters', type: 'toggle', label: 'Show category filters', description: '', group: 'Display', default: 'true' },
        { key: 'newDays', type: 'number', label: 'Mark as New for (days)', description: '', group: 'Display', default: '7' },
        { key: 'listName', type: 'list', label: 'Announcements list', description: '', group: 'Data', default: 'Announcements' },
      ],
      lists: [{
        name: 'Announcements', description: 'Items shown by the Announcements Board web part',
        columns: [
          { name: 'Title', type: 'Text', required: true, choices: [] },
          { name: 'Body', type: 'Note', required: false, choices: [] },
          { name: 'Category', type: 'Choice', required: false, choices: ['General', 'HR', 'IT', 'Facilities', 'Events'] },
          { name: 'PublishDate', type: 'DateTime', required: false, choices: [] },
          { name: 'Expires', type: 'DateTime', required: false, choices: [] },
          { name: 'Pinned', type: 'Boolean', required: false, choices: [] },
          { name: 'Image', type: 'Image', required: false, choices: [] },
          { name: 'Contact', type: 'User', required: false, choices: [] },
        ],
        sample: [
          ['Open enrollment starts Monday', 'Review your health, dental and vision options before the window closes on the 30th. Benefits counselors are available in Building 2 every Tuesday and Thursday.', 'HR', '-1', '+21', 'true', '', 'Priya Natarajan'],
          ['New VPN client rollout', 'The new VPN client installs automatically this week. Restart when prompted; your settings carry over.', 'IT', '-3', '+14', 'false', '', 'Marcus Bell'],
          ['Holiday party RSVP', 'Join us in the atrium on the 18th from 3 to 5 pm. RSVP by the 12th so we can plan food.', 'Events', '-5', '+20', 'false', '', 'Sofia Ramirez'],
          ['Parking garage resurfacing', 'Level 2 of the north garage is closed through Friday. Use the south garage or level 3.', 'Facilities', '-9', '+3', 'false', '', 'Daniel Okafor'],
          ['Quarterly town hall recording', 'Missed the town hall? The recording and slides are on the Communications site.', 'General', '-14', '', 'false', '', 'Hannah Kim'],
        ],
      }],
      css: `.an-head{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px}
.an-head h2{font-size:20px;font-weight:600;margin:0;margin-right:auto}
.an-chips{display:flex;gap:6px;flex-wrap:wrap}
.an-chip{border:1px solid var(--sss-border);background:var(--sss-bg);color:var(--sss-text);border-radius:999px;padding:4px 12px;font:inherit;font-size:12px;cursor:pointer}
.an-chip[aria-pressed="true"]{background:var(--sss-primary);border-color:var(--sss-primary);color:var(--sss-bg)}
.an-list{display:grid;gap:12px}
.an-card{display:grid;grid-template-columns:auto 1fr;gap:14px;padding:14px;border:1px solid var(--sss-border);border-radius:10px;background:var(--sss-bg)}
.an-card.is-pinned{border-color:var(--sss-primary);box-shadow:inset 3px 0 0 var(--sss-primary)}
.an-img{width:96px;height:72px;border-radius:8px;object-fit:cover;background:var(--sss-surface)}
.an-badge{width:44px;height:44px;border-radius:10px;display:grid;place-items:center;background:var(--sss-primary-lighter);color:var(--sss-primary)}
.an-meta{display:flex;gap:8px;align-items:center;flex-wrap:wrap;font-size:12px;color:var(--sss-muted);margin-bottom:4px}
.an-cat{font-weight:600;color:var(--sss-primary);text-transform:uppercase;letter-spacing:.04em;font-size:11px}
.an-new{background:#107c10;color:#fff;border-radius:4px;padding:1px 6px;font-size:11px;font-weight:600}
.an-pin{display:inline-flex;align-items:center;gap:3px;color:var(--sss-primary);font-weight:600}
.an-card h3{font-size:15px;font-weight:600;margin:0 0 4px}
.an-body{margin:0;font-size:13px;line-height:1.5;color:var(--sss-text)}
.an-more{border:0;background:none;color:var(--sss-link);font:inherit;font-size:13px;padding:0;margin-top:6px;cursor:pointer}
.an-more:hover{text-decoration:underline}
.an-contact{font-size:12px;color:var(--sss-muted);margin-top:6px}
.an-contact a{color:var(--sss-link)}
.an-empty{padding:22px;border:1px dashed var(--sss-border);border-radius:10px;color:var(--sss-muted);text-align:center}
@container (max-width:480px){.an-card{grid-template-columns:1fr}.an-img{width:100%;height:140px}}`,
      code: `const list = props.listName;
const head = h('div', { class: 'an-head' }, props.heading ? h('h2', null, props.heading) : null);
ctx.root.appendChild(head);
if (!list) { ctx.root.appendChild(h('div', { class: 'an-empty' }, ctx.page.editMode ? 'Pick the announcements list in the property pane.' : 'Nothing to show yet.')); return; }
const today = new Date().toISOString().slice(0, 10);
let rows;
try {
  rows = await ctx.cache('rows:' + list, () => sp.items(list, {
    select: ['Id', 'Title', 'Body', 'Category', 'PublishDate', 'Expires', 'Pinned', 'Image', 'Created', 'Contact/Title', 'Contact/EMail'],
    expand: ['Contact'],
    filter: "(Expires ge datetime'" + today + "T00:00:00Z') or (Expires eq null)",
    orderBy: ['Pinned desc', 'PublishDate desc'], top: 100 }));
} catch (e) {
  if (e.status === 404) { ctx.root.appendChild(h('div', { class: 'an-empty' }, 'The list "' + list + '" was not found.' + (ctx.page.editMode ? ' Create it from the Lists section of the property pane.' : ''))); return; }
  throw e;
}
const cats = [...new Set(rows.map((r) => r.Category).filter(Boolean))];
const state = ctx.state;
if (state.cat && !cats.includes(state.cat)) state.cat = '';
const listEl = h('div', { class: 'an-list', 'aria-live': 'polite' });
const draw = () => {
  ctx.clear(listEl);
  const shown = rows.filter((r) => !state.cat || r.Category === state.cat).slice(0, props.maxItems || 6);
  if (!shown.length) { listEl.appendChild(h('div', { class: 'an-empty' }, 'No announcements' + (state.cat ? ' in ' + state.cat : '') + '.')); return; }
  for (const r of shown) {
    const when = r.PublishDate || r.Created;
    const isNew = when && (Date.now() - new Date(when).getTime()) < (props.newDays || 7) * 86400000;
    const img = props.showImages && r.Image && r.Image.Url;
    const body = fmt.plain(r.Body);
    const long = body.length > 160;
    const p = h('p', { class: 'an-body' }, long ? fmt.truncate(body, 160) : body);
    listEl.appendChild(h('article', { class: 'an-card' + (r.Pinned ? ' is-pinned' : '') },
      img ? h('img', { class: 'an-img', src: img, alt: '' }) : h('div', { class: 'an-badge' }, icon('megaphone', 22)),
      h('div', null,
        h('div', { class: 'an-meta' }, r.Category ? h('span', { class: 'an-cat' }, r.Category) : null, h('span', null, fmt.date(when)),
          isNew ? h('span', { class: 'an-new' }, 'New') : null, r.Pinned ? h('span', { class: 'an-pin' }, icon('bookmark', 12), 'Pinned') : null),
        h('h3', null, r.Title), p,
        long ? h('button', { class: 'an-more', type: 'button', 'aria-expanded': 'false', onclick: (e) => {
          const open = e.currentTarget.getAttribute('aria-expanded') === 'true';
          p.textContent = open ? fmt.truncate(body, 160) : body;
          e.currentTarget.setAttribute('aria-expanded', String(!open)); e.currentTarget.textContent = open ? 'Read more' : 'Show less';
        } }, 'Read more') : null,
        r.Contact && r.Contact.Title ? h('div', { class: 'an-contact' }, 'Contact: ', r.Contact.EMail ? h('a', { href: 'mailto:' + r.Contact.EMail }, r.Contact.Title) : r.Contact.Title) : null)));
  }
};
if (props.showFilters && cats.length > 1) {
  const chips = h('div', { class: 'an-chips', role: 'group', 'aria-label': 'Filter by category' });
  const mk = (label, value) => h('button', { class: 'an-chip', type: 'button', 'aria-pressed': String((state.cat || '') === value),
    onclick: (e) => { state.cat = value; chips.querySelectorAll('.an-chip').forEach((b) => b.setAttribute('aria-pressed', String(b === e.currentTarget))); draw(); } }, label);
  chips.appendChild(mk('All', ''));
  cats.forEach((c) => chips.appendChild(mk(c, c)));
  head.appendChild(chips);
}
ctx.root.appendChild(listEl);
draw();`,
      notes: 'Expired items (Expires before today) are hidden. Pinned items come first.',
    },
  },
  {
    key: 'countdown',
    blurb: 'A live countdown to an event date, with a message once it starts. No list needed.',
    spec: {
      title: 'Event Countdown', alias: 'EventCountdown', icon: 'Clock', group: 'Text, media, and content',
      description: 'A live countdown to an event, with a message once it starts.',
      properties: [
        { key: 'eventName', type: 'text', label: 'Event name', description: '', group: 'Event', default: 'Holiday Party' },
        { key: 'eventDate', type: 'text', label: 'Date and time', description: 'For example 2026-12-18 15:00', group: 'Event', default: '2026-12-18 15:00' },
        { key: 'location', type: 'text', label: 'Location', description: '', group: 'Event', default: 'Main atrium' },
        { key: 'link', type: 'url', label: 'RSVP link', description: '', group: 'Event', default: '' },
        { key: 'afterText', type: 'text', label: 'Message once it starts', description: '', group: 'Event', default: 'Happening now. See you there!' },
        { key: 'style', type: 'dropdown', label: 'Style', description: '', group: 'Display', default: 'Bold', options: ['Bold', 'Light'] },
        { key: 'showSeconds', type: 'toggle', label: 'Show seconds', description: '', group: 'Display', default: 'true' },
      ],
      lists: [],
      css: `.cd{border-radius:14px;padding:22px 24px;display:grid;gap:14px}
.cd.bold{background:linear-gradient(135deg,var(--sss-primary),color-mix(in srgb,var(--sss-primary) 70%,#000));color:var(--sss-bg)}
.cd.light{background:var(--sss-primary-lighter);color:var(--sss-text);border:1px solid var(--sss-border)}
.cd-top{display:flex;align-items:center;gap:10px;font-size:13px;opacity:.9}
.cd h2{margin:0;font-size:22px;font-weight:600;line-height:1.25}
.cd-units{display:grid;grid-template-columns:repeat(var(--n),minmax(0,1fr));gap:10px;max-width:520px}
.cd-unit{border-radius:10px;padding:10px 6px;text-align:center}
.bold .cd-unit{background:color-mix(in srgb,var(--sss-bg) 14%,transparent)}
.light .cd-unit{background:var(--sss-bg);border:1px solid var(--sss-border)}
.cd-num{display:block;font-size:30px;font-weight:700;font-variant-numeric:tabular-nums;line-height:1.1}
.cd-lbl{font-size:11px;text-transform:uppercase;letter-spacing:.06em;opacity:.85}
.cd-foot{display:flex;gap:14px;align-items:center;flex-wrap:wrap;font-size:13px}
.cd-foot span{display:inline-flex;align-items:center;gap:6px}
.cd-btn{display:inline-flex;align-items:center;gap:6px;border-radius:8px;padding:7px 14px;font-weight:600;text-decoration:none}
.bold .cd-btn{background:var(--sss-bg);color:var(--sss-primary)}
.light .cd-btn{background:var(--sss-primary);color:var(--sss-bg)}
.cd-now{font-size:18px;font-weight:600}
@container (max-width:420px){.cd-num{font-size:22px}.cd{padding:16px}}`,
      code: `const when = new Date(String(props.eventDate || '').replace(' ', 'T'));
const box = h('div', { class: 'cd ' + (props.style === 'Light' ? 'light' : 'bold'), role: 'timer', 'aria-label': 'Countdown to ' + (props.eventName || 'the event') });
ctx.root.appendChild(box);
if (isNaN(when)) { box.appendChild(h('div', null, ctx.page.editMode ? 'Enter the event date in the property pane, like 2026-12-18 15:00.' : 'Date to be announced.')); return; }
box.appendChild(h('div', { class: 'cd-top' }, icon('calendar', 16), fmt.dateTime(when)));
box.appendChild(h('h2', null, props.eventName || 'Upcoming event'));
const units = [['days', 86400], ['hours', 3600], ['minutes', 60]].concat(props.showSeconds ? [['seconds', 1]] : []);
const nums = {};
const grid = h('div', { class: 'cd-units', style: { '--n': String(units.length) } },
  units.map(([u]) => h('div', { class: 'cd-unit' }, nums[u] = h('span', { class: 'cd-num' }, '0'), h('span', { class: 'cd-lbl' }, u))));
box.appendChild(grid);
const foot = h('div', { class: 'cd-foot' },
  props.location ? h('span', null, icon('location', 16), props.location) : null,
  props.link ? h('a', { class: 'cd-btn', href: props.link, target: '_blank', rel: 'noopener' }, 'RSVP', icon('arrowRight', 14)) : null);
box.appendChild(foot);
const tick = () => {
  if (!ctx.isCurrent()) return clearInterval(timer);
  let left = Math.max(0, Math.floor((when - Date.now()) / 1000));
  if (left === 0) { grid.replaceWith(h('div', { class: 'cd-now' }, props.afterText || 'Happening now')); return clearInterval(timer); }
  for (const [u, s] of units) { nums[u].textContent = String(Math.floor(left / s)); left %= s; }
};
const timer = setInterval(tick, 1000);
tick();`,
      notes: "The countdown uses the viewer's own time zone.",
    },
  },
];
