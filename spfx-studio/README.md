# SPFx Studio

Describe a SharePoint web part in plain English. Claude writes it, the studio previews it against sample data with the real property pane, and you download a deployable **.sppkg**. You don't need Node, gulp or a build server. The package loads nothing from outside SharePoint, so it can be used in GCC High.

Live at `/spfx-studio/` on the site. Everything runs in the visitor's browser. Calls to Claude go straight to `api.anthropic.com` with the visitor's own key, which is shared with Power App Studio on the same site.

## How it works

1. **Describe**: plain-English requirements → Claude (structured output) → a *spec*. The spec holds:
   - the toolbox title, alias, group and icon
   - property-pane fields
   - the SharePoint lists the web part uses, with sample rows
   - CSS
   - the web part code (the body of an async function)

   `cleanSpec()` repairs anything out of bounds. `checkCode()` syntax-checks the code, and if it fails, one automatic repair round goes back to Claude.
2. **Preview**: the bundle that goes into the package runs, unchanged, in a sandboxed iframe (`allow-scripts`, no same-origin, so it can't read the key). Stand-ins replace SharePoint:
   - an AMD `define`
   - the five SPFx modules the bundle imports
   - a page context and a theme service
   - an in-memory **SharePoint REST emulator** seeded with the sample rows. It handles `$select`/`$expand`/`$filter`/`$orderby`/`$top`, paging, add/update (MERGE)/delete, `createfieldasxml`, current user and ensure user. It also returns SharePoint's real errors for common mistakes, such as selecting `Person/Title` without `$expand`.

   The preview toolbar sets the section width (one-third, half, one column, full width), the section background (white, soft, strong/inverted), the site theme colour and edit mode. The REST log shows every call. Runtime errors get an **Ask Claude to fix it** button.
3. **Refine**: *Ask for a change* sends the current spec back to Claude with the change. The **Code** tab edits the JS, CSS, properties and lists directly. **Version history** keeps the last 25 states.
4. **Download .sppkg**, or a **source project** (a standard SPFx 1.21.1 project with the same ids, if you want to keep going with the toolchain).

## The package

`js/sppkg.js` writes the same Open Packaging Conventions zip that `gulp package-solution --ship` writes for SPFx 1.21.1:

- `AppManifest.xml`
- one feature with the web part's element manifest
- the `ClientSideAssets` feature carrying the bundle (`includeClientSideAssets: true`)
- `_rels`
- `[Content_Types].xml`

The bundle (`js/bundle.js`) is an AMD module named `<componentId>_<version>`. It depends on the five framework components below, and its default export is the web part class, the same contract webpack's output follows:

- `sp-core-library`
- `sp-property-pane`
- `sp-webpart-base`
- `sp-http`
- `sp-component-base`

`tests/fixtures/toolchain-announcements.js` is the toolchain's own bundle for the Announcements example. It runs in the preview host the same way the studio's bundle does.

- **Tenant-wide by default**: `skipFeatureDeployment`. You can untick it in the Package tab.
- **Upgrades**: the solution and web part ids stay with the project. Each download of a changed web part bumps the version (`1.0.N.0`, web part `1.0.N`), so uploading it over the old package upgrades every page that uses it.
- **Lists**: each list in the spec gets a **Create "…" list** button in the web part's property pane. The button creates the list and its columns through REST (`createfieldasxml`, added to the default view) and adds the sample rows. Existing lists are left alone; missing columns are added.

## The runtime contract (what generated code can use)

`js/runtime.js` is serialised into every bundle. Generated code sees:

| Name | What it is |
|---|---|
| `ctx.root` | Element to build into (empty on each render) |
| `h(tag, attrs, ...children)` | Element builder (events as `onclick`, styles as objects, SVG tags work) |
| `icon(name, size)` | 55 built-in line icons as inline SVG (no font or CDN) |
| `props` | Property values, typed |
| `sp` | `items`, `item`, `add`, `update`, `remove`, `me`, `ensureUser`, `lists`, `fields`, `listExists`, `ensureList`, `at(siteUrl)` |
| `fmt` | `date`, `dateTime`, `time`, `relative`, `number`, `initials`, `truncate`, `plain` |
| `ctx.cache(key, fn)` / `ctx.refresh()` / `ctx.rerender()` | Data caching across renders |
| `ctx.state` | State that survives re-renders |
| `ctx.page`, `ctx.user`, `ctx.theme` | Page context, current user, theme colours |
| `ctx.isCurrent()` | False once a newer render has started; stop timers and late callbacks |

CSS is scoped to the web part automatically. Theme variables are `--sss-primary`, `--sss-primary-dark`, `--sss-primary-light`, `--sss-primary-lighter`, `--sss-text`, `--sss-muted`, `--sss-bg`, `--sss-surface`, `--sss-border` and `--sss-link`. They follow the site theme and section variants. The root element is a size container, so `@container` queries work.

## Files

| File | Purpose |
|---|---|
| `index.html`, `studio.css` | The studio shell |
| `js/app.js` | UI: Build/Code/Package tabs, preview bridge, property pane, downloads, history |
| `js/ai.js` | Claude request (schema, system prompt with the Quick Links example), `cleanSpec`, `checkCode` |
| `js/runtime.js` | The library every web part carries |
| `js/bundle.js` | AMD bundle, manifest, CSS scoping, names and versions |
| `js/sppkg.js` | The .sppkg |
| `js/source.js` | SPFx 1.21.1 source project |
| `js/preview.js` | Preview host: SPFx stand-ins and the REST emulator |
| `js/examples.js` | Quick Links Tiles, Announcements Board, Event Countdown |
| `tests/run.mjs` | `node spfx-studio/tests/run.mjs` |
| `tests/harness.html` | Mounts an example (`?ex=`) or any bundle (`?bundle=`) in the preview host |

## Deploying a web part

1. Open the SharePoint admin center → **More features** → **Apps** → **Open** → **Apps for SharePoint**.
2. **Upload** the .sppkg. Keep **Make this solution available to all sites** checked, then click **Deploy**.
3. Edit a page, click **+**, and search for the web part's title.
4. In the property pane, go to **Lists** and click **Create "…" list**. This needs Manage Lists permission.
