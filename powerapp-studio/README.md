# Power App Studio

A browser-only Power Apps canvas app designer for satechitsolutions.com. There is no build step and no server: GitHub Pages serves these files as they are.

## What it does

1. **Requirements wizard**: pick a scenario, shape the SharePoint lists, then choose the screens and a theme. It generates wired screens: home, browse, details, edit and approvals.
2. **Designer**: drag-and-drop canvas, gallery templates, inputs bound to columns, an actions builder, Visible rules, Power Fx overrides, layers, undo, copy and paste, alignment tools.
3. **Theme**: presets, 16 color tokens, font, corner radius. These are emitted as one `appTheme` named formula.
4. **Play**: runs the app against sample data, including search, open, edit, save, delete, approve/reject, variables and people pickers.
5. **Get YAML**: App.Formulas, App.OnStart, App.StartScreen and one paste-ready YAML block per screen. It also produces a lint report, a provisioning script (PnP PowerShell, or a GCC High / DoD flavour; see below), the project JSON and a zip of everything.
6. **Import**: paste YAML copied from Power Apps Studio (or a `Src/*.pa.yaml` file) to keep editing it here.
7. **Flows**: Power Automate flows for the app's lists, exported as legacy flow packages. There are four templates: email on a new item, email when a Choice column changes, a daily reminder, and a flow a button runs. The email editor uses `{Column}` tokens and `{Link}`, which deep-links into the app. Play mode shows which flows would fire.
8. **Describe your app** (uses Claude): write requirements in plain English. Claude returns a plan of lists, screens, approvals and flows, which you review and refine before the studio builds it. It uses the visitor's own Anthropic API key.
9. **Paste assistant**: Get YAML opens a step-by-step deploy. Each step has one Copy or Download button and a "Done, next" button. It records a hash of what each step pasted, so the next deploy offers "Only what changed".
10. **Finish setup**: one checklist for the values an app needs before it deploys cleanly: site URL, app web link, email-link screen, admin group ID and AppUsage list. Fixing a field clears its lint error at once.
11. **Customer guide**: Export guide writes one self-contained HTML page with a lifecycle diagram, screen thumbnails, data model, who can do what, emails and flows, and numbered test cases with Pass/Fail, notes and Copy results.
12. **Demo and tour**: first-time visitors can open a finished Request Tracker with flows and a 4-step tour.
13. **Ask for a change** (uses Claude): describe a change to the app you already have ("add a required Cost Center to the form and details"). Claude returns small edit operations, you tick the ones to apply, and the studio changes only the screens they touch. One Undo reverts the whole change.
14. **Import from CSV or Excel**: the Data tab turns a `.csv` or `.xlsx` file into a SharePoint list with guessed column types, keeps up to 200 rows as Play sample data, and can add list, details and form screens for it.
15. **Contrast check and accessibility labels**: the Theme tab checks the theme's text colors against WCAG AA and fixes a failing pair in one click. An opt-in App setting adds `AccessibleLabel` to inputs and tap buttons.
16. **Check against a Studio export**: read an `.msapp`, a solution `.zip` or `Src/*.pa.yaml` files, list the controls and screens that exist only in Studio, and bring them in exactly as Studio wrote them.
17. **Multi-stage approval chains**: 1 to 4 ordered stages, each with its own approver (a Person column on the item or a Microsoft 365 group), decision, date and comment, plus an append-only history. Flows email each approver in turn and the requester at the end.
18. **My blocks**: save a selection of controls as a reusable block, insert it in any app with fresh names, and export or import blocks as a file.
19. **Share a Play preview**: a link that carries the app inside its `#play=` fragment and opens a Play-only viewer for stakeholders.

## What the generator builds

The wizard, the Describe dialog and **New screen** all use the same templates (`js/templates.js`). They follow the hand-built apps in the powerapps-canvas-yaml skill:

- **Home**: a gradient hero with a greeting and the main action, KPI cards with icon chips (one per status value), a *My items* list and a *Waiting on me* queue, each with an empty state. Phones get a two-column layout.
- **Browse**: search, an **All / Mine** switch, a status filter ("All statuses" plus the choices), a row count, card rows (title, meta line, status pill, chevron), an empty state, and a floating + button on phones. Each combination is its own delegable `Filter` chosen with `If` (Mine is `'Created By'.Email = User().Email`).
- **Details**: a summary card ("Created by … on … · #ID"), a status stepper, field sections, Approve / Reject while the item is pending, and a **two-click Delete** ("Confirm delete" on the first click). It shows an empty state when nothing is selected.
- **New / edit**: section cards (Details, People and dates, Notes) in two or three columns and a bottom action bar. Saving with blank required fields sets `varTriedSave<List>`, which turns those inputs' borders red and shows a **Still needed: …** strip. Saving a new item opens a **confirmation page** with *Add another* and *Done*.
- **Approvals**: a queue with a count and a review panel, with an "All caught up" empty state.
- **Loading screen** (optional): an original hexagon monogram drawn as a SMIL SVG and gated on `locInited`, plus a separate progress ring. A hidden self-starting timer refreshes the lists, records usage and checks access, then navigates home from `OnTimerEnd`, never from `OnVisible`. It becomes the start screen.
- **Admins** (optional): give a Microsoft 365 group ID and `App.OnStart` sets `varIsAdmin` with `IfError(... Office365Groups.ListGroupMembers(id).value, Lower(mail) = Lower(User().Email) ...)`. Approve, Reject, Delete and the Approvals tile then show only for admins. Play has a "Signed in as an admin" toggle.
- **Usage tracking** (optional): an `AppUsage` list (provisioned by the PnP script) and the skill's `IfError` upsert, one row per user.

Spacing and type follow one scale: 16/24 px padding, 12–26 px type, the theme radius on every card, and the theme font in HTML (`appTheme.FontCss`, which has no quotes so it can sit inside `style='…'`). Every HtmlViewer obeys the 5 px padding rule: the first box's height plus its margins is at most Height − 10. Lint PS051 checks this.

Wizard step 3 ("Extras") and the Claude plan's `options` object choose these features: `loadingScreen`, `confirmationPage`, `statusStepper`, `showMineFirst`, `trackUsage` and `adminGroupId`. Older projects and plans without them get the defaults: loading, confirmation and stepper on, All first, no usage, no admin group.

In the designer, conditional overlays (confirmation page, empty states, "Still needed") are evaluated as in Play. Tick **Show hidden** in the stage bar to see them all, or select one in Layers.

### Needs a real-tenant paste test

**Round 3 additions to test in a tenant.** `AccessibleLabel` is not in the skill's field-tested references. That is why it is opt-in (App → *Add accessibility labels*) and off by default, and default YAML does not change. Paste one screen with it on before you rely on it. Screens edited by *Ask for a change* reuse the generator's own controls, so they need no new properties. Controls imported from a Studio export are re-emitted exactly as Studio wrote them, except that Width and Height are added when the export left them out.

These templates passed the harness and the skill's `screen_tool.py check`, but have not yet been pasted into a tenant: the stepper HTML built with `With()`, the confirmation overlay on a multi-screen app, the Mine/All `If` branches over `SortByColumns(Filter(...))`, the loading-screen emblem data URI with the timer step machine, and `appTheme.FontCss` inside HtmlViewer styles. The admin check (`mail` field) and the AppUsage block are copied from verified skill patterns.

## Files

| File | Role |
|---|---|
| `index.html`, `studio.css` | Shell and styles (SaTech ink and cyan) |
| `js/core.js` | Project model and JSDoc types, control catalog, column types, theme presets, sample data, `renameCopy` (unique names for copies), `normalizeProject` (makes loaded and shared projects safe) |
| `js/fx.js` | Power Fx builders for the field-tested shapes: `ifErrorFx`, `upsertBaseFx`, `rebaseFx`, `spUserFx` / `meUserFx`, `recordFx`, `equalsFx`, `notifyFx`, `navigateFx` |
| `js/emit.js` | YAML and Power Fx generation, deploy checklist, provisioning scripts (PnP and GCC High), lint (`LINT_RULES`) |
| `js/scenarios.js` | Round 6: `buildScenario` (one call from a gallery card to a finished app with options, chain and flows), gallery summaries |
| `js/attach.js` | Round 6: attachments placeholders, form names, loader line, Paste assistant text |
| `js/versionstore.js` | Round 6: Version history in IndexedDB, with a localStorage fallback |
| `js/render.js` | Control rendering, HTML sanitizer, mini Power Fx evaluator, Play runtime |
| `js/templates.js` | Scenarios, screen templates, insertable blocks |
| `js/import.js` | YAML import with name-clash renaming |
| `js/ui.js` | Designer UI: canvas, panels, inspector, dialogs, Play |
| `js/flows.js` | Flow model and templates, Logic Apps definitions, legacy package files (`flowPackageFiles`), tokens to expressions, deep link (`deepLinkInfo`), flow lint |
| `js/ai.js` | Describe-your-app: plan JSON schema, system prompt, browser call to the Claude API, `specToProject` (cleans the plan and builds the project, wiring button flows and deep links) |
| `js/describeui.js` | The Describe dialog: requirements, key, model, review, refine, build |
| `js/keystore.js` | Where the visitor's API key and model choice live (tab or browser storage, never in project files) |
| `js/chain.js` | Multi-stage approvals: stage columns, decision and resubmit Patches, queue filter, stepper template, group checks, chain flows, lint, Play simulation |
| `js/blocks.js` | My blocks: make, sanitize, insert (fresh names, remapped formulas), store, export and import |
| `js/share.js` | Play preview links: trim and sanitize a copy, deflate-raw and base64url into `#play=`, decode safely |
| `js/edit.js` | Edits to an existing project that keep untouched screens byte-identical: list and column renames that rewrite the generator's baked formulas, add a column to a form, details screen or gallery rows (grows the section, reflows to three columns when it no longer fits), remove a column, admin gating, rebuild status visuals |
| `js/change.js` | Ask for a change: `CHANGE_SCHEMA` (structured outputs), `CHANGE_PROMPT`, `projectSummary()` (names and types only, no rows or setup values), `reviewOps()`, `applyOps()` |
| `js/changeui.js` | The Ask-for-a-change dialog: request, review with per-edit checkboxes and a dry run, apply as one undo step, new lint findings |
| `js/tabular.js` | CSV parser and writer, `.xlsx` reader (via JSZip: shared strings, inline strings, dates from cell styles), type inference, `listFromTable()`, `listToCsv()` |
| `js/tableui.js` | Data tab → Import CSV or Excel |
| `js/a11y.js` | WCAG contrast ratio, theme pairs, `fixContrast()` |
| `js/drift.js`, `js/driftui.js` | Check against a Studio export: `readStudioExport()`, `driftReport()`, `importDrift()` |
| `js/deploy.js` | Paste assistant model: ordered `deploySteps()`, FNV-1a `hash()`, `stepStatus()` (new / changed / unchanged), `markDeployed()`, `pendingSteps()`, `removedSinceDeploy()` |
| `js/setup.js` | Finish setup: `setupItems()` (each mapped to the lint codes it clears), `applySetup()`, GUID check |
| `js/guide.js` | Customer guide: `guideCases()` test cases from the app's actions and flows, `lifecycleSvg()`, `guideHtml()` |
| `js/assistui.js` | UI for the Paste assistant, Finish setup card and dialog, guide export (thumbnails from the designer renderer) and the tour |
| `js/feedback.js`, `js/feedbackui.js` | Report a paste error: `analyze()` maps Studio and Power Automate error text to a control, cause and fix; `applyAction()` (switch a property off or on, mark a list large); findings log and `exportFindings()` |
| `js/delegation.js` | Expected list size (`SIZES`), `capCounts()` ("2000+" KPI counts), `delegationIssues()` (PS059), `needsRowLimit()` |
| `js/history.js`, `js/historyui.js` | Version history: `takeSnapshot()`, `restoreSnapshot()`, `diffProjects()` / `describeDiff()`, `mergeSnapshots()` |
| `js/bundle.js` | App bundle zip in the skill's layout: `bundleFiles()`, `bundleZip()`, `schemaDigest()`, `readBundle()` (embedded project, or rebuilt from `canonical/` and `schema/`) |
| `js/flowsui.js` | Flows tab, new-flow picker, flow editor with email preview, package download, tenant-sample loader |
| `tests/run.mjs` | Test harness: no dependencies beyond Node 18+ and the vendored js-yaml |
| `vendor/` | js-yaml 4.1.0 and JSZip 3.10.1, served locally |

## Paste rules the emitter enforces

These were field-tested in commercial and GCC High tenants:

- The top-level key is `Screens:`. Children are a dash list.
- Classic controls carry no version pins. The one exception is `Gallery@2.15.0` with `Variant: Vertical`.
- Gallery template controls sit under the gallery's `Children:`. There is no `GalleryTemplate` child.
- Any formula containing a newline, `: `, ` #`, `{` or `}` is written as a `|-` block scalar.
- Rectangles never get Radius properties. Rounded shapes are HtmlViewer "cards" or Buttons.
- Transparent tap buttons go after what they cover, because later controls render in front.
- Control names are unique across the whole app.
- Patch always rebases with `LookUp(List, ID = varSel.ID)` and uses the 3-argument `IfError` shape.
- `Navigate()` is never used in OnStart or OnVisible.

Gallery rows are emitted as passive content first, then the row tap button, then anything clickable on its own (inputs, buttons with actions, icons with actions). That way a delete icon in a row still gets its own click.

## Lint rules

Every finding has a stable code. To silence one rule on one control, put `lint-ok PS0xx` in any of its formulas or custom actions, for example as a Power Fx comment: `// lint-ok PS040`.

| Code | Level | Rule |
|---|---|---|
| PS001 | error | A name is used twice anywhere in the app (PA2110 on paste) |
| PS002, PS003 | error | Invalid screen or control name |
| PS004, PS010, PS027 | error | `ThisItem` outside a gallery, or a row value shown outside one |
| PS005 | error | `IfError(Patch(...), true)`: the fallback must be a record |
| PS006 | error | `Set(collection, Blank())` strips the collection's type |
| PS007 | warn | `x in Combo.Selected.Field` is a substring test |
| PS008, PS009 | error | Binding to a column that is not in the data model |
| PS011–PS014 | warn/error | Gallery has no data, missing list, empty template, or missing search box |
| PS015 | error | People picker inside a gallery |
| PS016 | warn | Large embedded image |
| PS017, PS018 | warn | Control off the screen, or past its gallery row |
| PS019–PS022 | error | Action has a missing target or list, or "Open this row" is used outside a gallery |
| PS023, PS024 | warn | Save has no bound inputs, or a required column has no input on the screen |
| PS025 | error | Invalid variable name |
| PS026, PS028, PS042 | error | `Navigate()` in OnVisible or App.OnStart |
| PS029–PS034 | error/warn | Data model: duplicate lists or columns, names SharePoint already uses, spaces in names, Choice column without choices, Lookup column without a target |
| PS035 | info | `CountRows` is not delegable |
| PS036 | warn | HTML text with `<style>` or CSS rules: only inline styles render |
| PS037 | warn | Input on a Save screen is not bound to a column |
| PS038, PS039 | warn | Row control or "Open this row" refers to a different list than its gallery |
| PS040 | warn | A later control sits in front of a clickable one and takes its clicks |
| PS041 | warn | Visible depends on a variable that nothing sets |
| PS043–PS046 | error | Flow without a list, a valid recipient or the site URL, or watching a column that is not Choice |
| PS047 | warn | `{Link}` without the app web link, or without a screen set to open items from email links |
| PS048 | warn | Unknown `{token}` in a flow email |
| PS049 | error | "Run a flow" action without a flow of the button kind |
| PS050 | error | Two flows with the same Power Apps name once spaces are removed |
| PS051 | warn | An HTML box (first fixed `height:`, plus its margins) is taller than the HtmlViewer's Height − 10, so Power Apps shows a scrollbar |
| PS052 | info | A theme text/background pair (Text on Bg or Surface, TextMuted on Surface, OnPrimary on Primary, OnHeader on HeaderBg) is below WCAG AA 4.5:1. Since round 4 every preset passes. |
| PS053 | error | Approval chain list missing or stage count outside 1–4 |
| PS054 | error | Group-approved stage without a valid group object ID |
| PS055 | error | Stage approver column is not a Person column |
| PS056 | error | A chain column (stage decision, date, comment, history, current stage) is missing from the list |
| PS057 | error | Approve/Reject action without a chain, or Reject without a comment box |
| PS058 | warn | A formula override on a property the tenant rejected (switched off in *Report a paste error*): the override is dropped from the YAML |
| PS059 | warn | On a list marked Large: `CountRows` without the "2000+" cap, `Search()`, or a `Filter` on the list itself using `<>`, `in`, `= Blank()`, `IsBlank()` or a function over a column |

| PS060 | info / warn | Attachments: reminder to insert the hand-added Edit form over each placeholder after pasting (info); a placeholder whose list has Attachments off, or inside a gallery, does nothing (warn) |
| PS061 | error | Approval chain with admin override but no admin group |

To add a rule, give it the next code in `LINT_RULES`, implement it in `lint()` in `js/emit.js`, and add a failing case to the "lint rules fire" test in `tests/run.mjs`.

## Tests

```
node powerapp-studio/tests/run.mjs
```

The harness covers:

- Every scenario in tablet and phone format: no lint errors or warnings, and the YAML parses.
- Paste rules: gallery pin and variant, nested templates, no version pins, no Rectangle radius, `ThisItem` only in galleries, unique names, and block scalars wherever `: `, ` #`, `{` or `}` appear.
- Exact round-trip of tricky literals.
- Gallery click order and tap-button visibility.
- The Power Fx builders.
- The lint rules, including `lint-ok` suppression.
- Unique names after copying a screen.
- `normalizeProject` against hostile project files.
- YAML import round trip.
- The three-column layout for tall forms.
- Play-mode save and validation.
- Provisioning-script ordering.
- Every scenario in tablet and phone format with every option on: lint-clean with no warnings, paste rules, every HtmlViewer within Height − 10, loading screen first and navigating home from `OnTimerEnd`, admin check wrapped in `IfError`, AppUsage provisioned, the Mine filter, red borders, confirmation page and two-click delete present.
- The same app with every option off.
- PS051 and quote-free `FontCss`.
- Play: "Still needed" names the blank field, the confirmation page appears, two-click delete, items created in Play count as Mine, and the loading timer reaches Home.
- A plan's `options` reach the app, a bad group ID is dropped and reported, old plans and projects get defaults.
- Paste assistant: the hash is stable, the deploy record survives save/load, editing one control flips only that screen, editing a flow flips that flow, removed screens are reported, and a hostile deploy record is cleaned.
- Finish setup: after Describe without a site URL, entering it clears every PS045. App link, email-link screen, admin GUID (bad values refused) and the AppUsage confirmation all tick their items.
- Customer guide: test cases match the app (one per flow, per distinct decision, per save screen, per delete list), are numbered, and the page makes no external requests.
- Demo: lint-clean with no warnings, setup complete, two flows, paste rules hold.
- Ask for a change: the schema stays inside structured-output limits (closed objects, every property required, no size or `oneOf` keywords). The request carries names and types only. A mocked response (`tests/fixture-change.json`) is reviewed (invalid edits flagged, destructive ones marked) and applied. Every untouched screen's YAML stays byte-identical and the original project is not modified, which is what Undo restores. Adding a field reflows the tablet form so it still fits above the Save bar. A batch of list, screen, option and flow edits leaves no stale names behind.
- CSV and Excel: quotes, doubled quotes, newlines in quotes, BOM, semicolon and tab files, write/read round trip. Type inference for every type. A generated `.xlsx` read in workbook sheet order, with rich-text strings and style-based dates. Imported lists get clean names, are capped at 200 rows, survive `normalizeProject`, feed Play, and export back to CSV.
- Accessibility: contrast ratios match WCAG reference values, and the fix reaches AA with the smallest nudge. PS052 is info only. `AccessibleLabel` is absent by default and present on inputs and tap buttons when opted in.
- Report a paste error: one case per error family (PA2108, PA2110, PA2109, PA1001, unknown names for lists, connectors, `appTheme`, `ThisItem`, screens, variables and flows, record/text type errors, column names, delegation warnings, flow connection, template and import errors). Switching a property off removes it only for that control type and changes nothing else; PS058; the findings log is capped, survives load and is stripped from share links; an empty `suppressProps` is byte-identical.
- App bundle: project → zip → project is identical; the schema digest shape; import without an embedded project rebuilds screens, lists and the start screen; a zip that is not a bundle is explained.
- Large lists: each PS059 pattern; every scenario marked Large stays lint-clean and passes the paste rules; Small is byte-identical; capped KPI counts; non-delegable extra filters move outside; the row-limit setup item.
- Version history: diff (added column, changed control, added flow, theme), restore, the cap drops automatic versions first, junk snapshots are dropped on load.
- Drift: a solution zip with a nested `.msapp` is read, Studio-only controls and screens are found, imported verbatim (formula positions kept) without touching other screens, and an `.msapp` without YAML source is explained.

Run the harness before every push. If you change the emitter on purpose, look at the diff of a generated screen before pasting it into Studio.

## Local testing

ES modules do not load from `file://`. Serve the folder instead, for example `python -m http.server`, then open `/powerapp-studio/`.

## Keyboard

When the canvas has focus (click it, or Tab to it):

- Enter selects the first control, and Tab / Shift+Tab move through the controls.
- Arrow keys move the selection by 1px, or 10px with Shift. Hold Alt while dragging to skip snapping.
- Delete removes the selection. Ctrl+C, Ctrl+V and Ctrl+D copy, paste and duplicate. Ctrl+Z and Ctrl+Y undo and redo.
- Escape deselects, after which Tab leaves the canvas.
- F5 opens Play.

## Flows

Each flow downloads as a `.zip` in the legacy package layout Power Automate exports:

```
manifest.json
Microsoft.Flow/flows/manifest.json
Microsoft.Flow/flows/<key>/definition.json
Microsoft.Flow/flows/<key>/apisMap.json
Microsoft.Flow/flows/<key>/connectionsMap.json
```

To install one: My flows → Import → **Import Package (Legacy)** → upload the zip → **Create as new** → pick your SharePoint and Office 365 Outlook connections → Import → turn the flow on. Packages contain no connection ids, so they import into any tenant.

What each template builds:

| Template | Trigger | Actions |
|---|---|---|
| When an item is created | SharePoint `GetOnNewItems` | Send an email (V2) |
| When a column changes | SharePoint `GetOnUpdatedItems` | `GetItemChanges` (needs version history), a Condition on `ColumnHasChanged` and the new value, then Send an email (V2) |
| Daily reminder | Recurrence at a set hour and time zone | Get items with an OData `$filter` (value and age), then Apply to each → Send an email (V2) |
| Run from a button | Power Apps (V2) with an `ItemID` number | Get item, Send an email (V2), Respond to a PowerApp |

Flows read columns by internal name, which is the column name without spaces, the same as the provisioning script creates. In the app, "Run a flow" emits `FlowName.Run(varSelList.ID)` inside the 3-argument `IfError`. The name is the flow's display name without spaces or symbols, and the button flow must be added in Studio's Power Automate pane.

Email links: set the app web link and the screen that opens items (App tab). `App.StartScreen` becomes `If(IsBlank(Param("ID")), <start>, <screen>)`. OnStart and that screen's OnVisible both load the item from `Param("ID")`, so the non-blocking OnStart cannot overwrite it.

The flow packages have not yet been import-tested in a real tenant. To match a tenant's export exactly, load one of its own exported flow packages under Flows → Tenant sample; connector display names, icons and connection names are copied from it.

## Describe your app (Claude)

The browser calls `https://api.anthropic.com/v1/messages` directly, sending the `anthropic-dangerous-direct-browser-access: true` header and `anthropic-version: 2026-10-21`. It asks for JSON that matches `PLAN_SCHEMA` through structured outputs (`output_config.format`). No server is involved:

- The API key is typed by the visitor. It is kept in `sessionStorage`, or `localStorage` if they tick Remember, and never written to project files or exports.
- The description goes to Anthropic and is billed to the visitor's key. The dialog says so before sending.
- `specToProject()` never trusts the plan:
  - Names are made into identifiers, and names SharePoint already uses are renamed.
  - Choice columns without choices and Lookups to missing lists become text.
  - Recipients must be Person columns or real addresses.
  - Status values must exist, and hours and days are clamped.
  - Email tokens follow renamed columns.
  - Every fix is shown to the visitor.
- Refining sends the requirements, the previous plan and the requested change as a three-turn conversation.
- `tests/fixture-plan.json` is a deliberately messy plan. The tests check that it builds a lint-clean project, and they mock `fetch` to check the request headers and error handling.

Corporate and government networks often block `api.anthropic.com`; the dialog reports that case clearly. The Claude artifact preview also blocks outside requests, so this feature only works on the live site.

## Ask for a change (Claude)

**Ask for a change** sits in the top bar, the Project menu and the Describe dialog (when an app already exists). Claude gets a compact summary of the project and the request: app options, lists with column names and types, screens with their kind, form and shown columns, galleries, and flows. It never gets sample rows, the site URL or the app link. Claude answers with `{summary, ops[], notes[]}` through structured outputs. The edit operations are `renameApp`, `setTheme`, `setOption`, `addList`, `renameList`, `addColumn`, `renameColumn`, `removeColumn`, `setChoices`, `addFieldToForm`, `addFieldToDetail`, `addFieldToGallery`, `addScreenFromTemplate`, `addFlow`, `editFlow` and `removeFlow`.

- `reviewOps()` checks each operation against the project, in order, so later edits can use a column added by an earlier one. Edits that cannot run are shown unticked with the reason. Renames, removals and choice changes are marked, because formulas you wrote by hand may still use the old values.
- The review shows a dry run of the ticked edits on a copy: which screens are new, which change, and anything worth knowing.
- Applying swaps in the edited copy as **one undo step**, then lists any new lint findings.
- Edits go through the same helpers as the designer (`js/edit.js`). Screens an edit does not touch keep byte-identical YAML, so the Paste assistant offers only the changed screens.
- Renaming a list or column, here or on the Data tab, now also rewrites the generator's baked formulas: table references, `varSel…`/`varScope…` names, `.Column` member access, Filter predicates, flow tokens and captions.

## Import from CSV or Excel

Data tab → **Import CSV or Excel**. Choose a `.csv`, `.tsv` or `.xlsx` file, or paste CSV. The file is read in the browser and never uploaded.

- **CSV**: quoted fields, doubled quotes, commas and line breaks inside quotes, CRLF, a BOM, and semicolon- or tab-separated files.
- **Excel**: the first sheet in workbook order, read with the vendored JSZip. Dates are recognized from the cell's number format. Save older `.xls` files as `.xlsx` first.
- **Types**:
  - Yes/No when every value is yes/no/true/false.
  - Person when every value is an email address.
  - Currency when values carry `$`, or when a money-sounding header (cost, price, amount…) has plain numbers.
  - Number otherwise for numbers. Values with leading zeros, such as ZIP codes, stay Text.
  - Date for ISO or US dates.
  - Choice for up to 12 repeated values.
  - Multiple lines of text for long or multi-line values.
  - Text for everything else.
- The preview lets you rename, retype or skip each column and pick the Title column. Names become identifiers, and names SharePoint already uses (Created, Location…) get the list's name in front.
- Up to 200 rows are kept in the project as Play sample data; the list card shows the count. **Export CSV** writes the rows with the internal column names as headers. Use them in SharePoint with *New → List → From CSV*, or paste them into the list's grid view (*Edit in grid view*) with the columns in the same order. Larger files belong straight in SharePoint.

## Accessibility

- **Theme → Contrast** shows the ratio of each text/background pair the screens use, with a **Fix** button for any pair below 4.5:1. The fix moves the text color's lightness (same hue) just far enough. When the text is already white or black, it moves the background instead. Lint PS052 (info) reports the same pairs.
- **App → Add accessibility labels (verify in your tenant)** is off by default. When on, it adds `AccessibleLabel` to text inputs, dropdowns, date pickers, toggles and people pickers (the bound column's name), and to the invisible tap buttons over icons, images and cards (taken from the control name, so `icoBack…` becomes "Back").

## Check against a Studio export

Project → **Check against a Studio export**. Choose an `.msapp` saved from Studio, a solution `.zip` (nested `.msapp` files are opened), or the screen `.pa.yaml` files from `Src`. The studio compares each screen with what it would emit:

- **Only in Studio**: top-level controls the project does not have. A container's children come with it.
- **Screens only in Studio.**
- **In your project, not in the export**: not pasted yet, or deleted in Studio.

Ticked items are imported as raw controls with every property exactly as Studio wrote them, including formula positions. They land in front of the screen's other controls, so you can paste the screen again without losing them. An `.msapp` without YAML source is reported, with what to do instead.

## Paste assistant, Finish setup, guide, demo

**Paste assistant** (Get YAML → Paste assistant, the default tab):
- **Steps, in order:** SharePoint lists (provisioning script), blank canvas app, data sources, App.Formulas, App.OnStart, flows (before the screens that run them), each screen, App.StartScreen, then test with Play.
- **Marking a step done** stores `{hash, at, name}` under `project.deployed[key]`. Keys are `formulas`, `scr:<screenId>`, `flow:<flowId>` and so on.
- **Hashes:** screens hash their full YAML. Flows hash their definition, not the package, since package keys are random per download. Lists hash their schema.
- **Next deploy:** "Only what changed" lists just the steps whose hash differs. It also names screens and flows deleted since the last deploy.
- **Re-pasted screens** show the warning about copying hand-added controls out to Notepad first.

**Finish setup** (top-bar button when anything is open, the Project menu, a card in the Paste assistant, and automatically after Generate or Describe → Build):

| Item | Shown when | Clears |
|---|---|---|
| SharePoint site URL | any enabled flow | PS045 |
| App web link, email-link screen and list | a flow uses `{Link}` | PS047 |
| Admin group ID | screens use `varIsAdmin`, or a group is set | the admin gate |
| AppUsage list exists | usage tracking is on | (confirmation only) |

**Customer guide** (Get YAML → Export guide, or the Project menu):
- Downloads `<App>-guide.html`.
- Screen thumbnails are the designer's own render plus the studio's `.pa-*` CSS rules. External `src`/`href`/`url()` references are stripped, so the page works offline.
- Pass/Fail marks and notes are kept in the reader's `localStorage` (wrapped in try/catch). Copy results writes a plain-text summary.

**Demo**: `buildDemo()` in `js/templates.js` builds a Request Tracker with every option and two flows, fully set up. The tour (`startTour()`) highlights Insert, the canvas, Play and Get YAML. Finishing or skipping it sets `pas.tour.done` in localStorage. Project → Take the tour runs it again.

**Designer polish**: on the Data tab the left panel widens to 380px, and column names get their own row, so long names are never cut off. The per-control orange "conditional" dots now show only on hover or selection; the "N conditional" chip in the stage bar marks them all.

None of this changes screen YAML: generated screens, formulas and flow definitions are byte-identical to before for every scenario and the fixture plan.

## Round 4: approval chains, My blocks, preview links, contrast

**Approval chains.** Turn the chain on in the wizard (step 3), in a Claude plan, or with *Ask for a change*. Each stage `K` adds its own columns to the list:

| Column | Type |
|---|---|
| `KDecision` | Choice: Pending, Approved, Rejected |
| `KDecidedOn` | Date |
| `KDecidedBy` | Person |
| `KComments` | Note |
| `KApprover` | Person, for column-approved stages only |

The chain itself adds `CurrentStage` (Number) and `ApprovalHistory` (Note, appended with one line per decision). Patterns follow the WAL lessons in the skill:

- Each stepper step is colored from that stage's own decision column, never from the overall status.
- Decisions rebase with `LookUp(List, ID = …)` inside the 3-argument `IfError`.
- Requesters never decide their own item.
- Group stages check membership with `Office365Groups.ListGroupMembers` wrapped in `IfError(…, false)`.
- Reject needs a comment.
- A rejected item can be resubmitted, which resets every stage to Pending and logs the reset in the history.

Flows email the next stage's approver whenever `CurrentStage` changes, and email the requester the final outcome. Projects without a chain emit byte-identical YAML.

**My blocks.** Select controls, then choose *Save as block* (inspector or multi-select). Blocks live in this browser's storage and can be exported or imported as `my-blocks.json`. Inserting one renames every control and remaps references in formulas and actions. Imported files are sanitized like any untrusted project.

**Share a Play preview.** Project → *Share a Play preview link*. The project is trimmed and sanitized:

- Removed: API keys, flow samples and setup drafts.
- Replaced: group IDs and fixed addresses.
- Left out unless ticked: the site URL and app link.
- Capped: imported sample rows, at 25.

It is then compressed with `deflate-raw` and base64url-encoded into `#play=…`. Browsers never send the fragment to a server. Links over about 16 KB get a warning, and links over about 100 KB are refused in favour of the project file. Opening a link shows a Play-only viewer with *Open in Studio as a copy*.

**Contrast.** The SaTech Cyan primary is now `#0E7490` and the Sunset primary `#C4451C`, so every preset passes WCAG AA. This changes App.Formulas for new projects that use those presets. Saved projects keep their stored colors.

**Deferred: Teams notifications in flows.** Whether the Microsoft Teams connector is available and allowed in GCC High tenants is not verified, and an untested connector in an imported package fails at import time. The Outlook email flows already cover notifications.

**Test in a tenant before relying on these:**
- A chain with a group stage: `Office365Groups` must be added as a data source.
- The decision Patch on a list with the generated chain columns.
- The stepper HtmlText.
- A resubmit after a rejection.

## Round 5: paste errors, app bundles, large lists, version history

**Report a paste error.** Top bar → *Paste error*, Project → *Report a paste error*, *Got an error?* in each Paste assistant step, or *Studio showed an error?* on the lint report. Paste what Studio (red banner, formula bar, App checker) or Power Automate's import page said; several messages at once are fine. Each one is mapped to the designer control and screen where the names allow it, with the likely cause and the fix, and buttons for the next step: select the control, open the Data tab, mark a list large, or, for PA2108 (unknown property), **switch that property off for that control type**. Switched-off properties live in the project as `app.suppressProps` (`{"Classic/DatePicker": ["IconBackground"]}`) and the emitter leaves them out of every screen. Every explained error goes into the **tenant log** (Tenant log tab), which you can export as text and send along when a paste keeps failing. With no switched-off properties the YAML is byte-identical to before.

**App bundle.** Project → *Export app bundle (.zip)* writes the powerapps-canvas-yaml skill's layout:

```
<App>_bundle_<date>/
  canonical/<screen>.pa.yaml, canonical/App.pa.yaml
  schema/schema.json            lists, internal and display names, types, choices
  flows/<flow>/definition.json
  reports/lint.txt, reports/drift.txt
  CHANGELOG.md                  from the version history and deploy records
  bundle.json                   format, generator, file list and the whole project
```

*Import app bundle…* restores the embedded project when `bundle.json` has one. Bundles made by `screen_tool.py` have none: the app is rebuilt from `canonical/*.pa.yaml` and `schema/schema.json` (App.Formulas and OnStart are regenerated, and the dialog says so). The current app is saved as a version first.

**Large lists.** Data tab → *Expected size* per list. Not set and Small change nothing. Medium and Large add a Finish-setup item to raise the data row limit to 2000 (Settings → General). Large also:

- caps KPI counts as `With({n: CountRows(...)}, If(n >= 2000, "2000+", Text(n)))`, because `CountRows` is never delegated to SharePoint;
- moves a gallery's extra filter outside when SharePoint cannot delegate it: `Filter(Filter(List, delegable), local)`;
- turns on PS059 for formulas you write.

The generated Mine, status and search filters were already delegable (`'Created By'.Email =`, Choice `.Value =`, `StartsWith`), so Large leaves them as they are.

**Version history.** Project → *Version history*: save a named version, compare any version with now (lists, columns, screens, controls, flows, theme, app settings), restore, delete. A version is saved automatically before Describe → Build, Ask for a change, a Studio import, a bundle import and a restore. Up to 15 are kept in the project (the oldest automatic ones go first), with a warning past about 1.5 MB because the browser keeps about 5 MB per site. Versions, the tenant log and deploy records are not part of Undo and are left out of share links.

**Test in a tenant before relying on these:**
- A PA2108 copied from Studio, switched off, and the screen pasted again.
- The "2000+" KPI labels and a reshaped extra filter on a list with more than 2,000 items (no delegation warning in Studio).
- A bundle made here opened by `screen_tool.py`, and one made by `screen_tool.py` imported here.

## Round 6: scenario gallery, attachments, approval-chain extras, version compare

**Scenario gallery.** *New app* opens on a gallery of twelve scenarios, each card with a rendered home screen and what it includes (lists, screens, flows, approvals, attachments). *Preview in Play* runs the scenario without touching your app; *Use this* goes to the data step. New scenarios: Work orders, Equipment checkout (lookup between two lists), Incident reports, Employee onboarding (lookup), Leave requests (two-stage chain: Supervisor column, then HR group) and Policy review (attachments on). Each ships with its flows (turn them off in the wizard) and builds lint-clean on tablet and phone with the options on or off.

**Attachments** (Data tab → a list's *Attachments* toggle, the skill's pattern). Attachments cannot be pasted, so the edit and details screens get a dashed placeholder naming the form to insert by hand: `frmAttach<List>` on the edit screen (an Edit form: Item `=varSel<List>`, DefaultMode Edit, only the Attachments card, unlocked with DisplayMode Edit) and `frmAttachView<List>` on details (View, referenced by no formula). With the toggle on, the edit screen's OnVisible calls `If(IsBlank(varSel.ID), NewForm(frm), EditForm(frm))`, Save passes `frm.Updates` as an extra `Patch` argument, and "Add another" calls `ResetForm`. The Paste assistant gets an *Attachments on <screen>* step with the exact steps and a Copy button; it comes back whenever the screen is re-pasted (a re-paste deletes hand-added controls, so the screen step says to keep a Notepad copy). *Report a paste error* recognises "frmAttach… isn't recognized". The provisioning script adds `Set-PnPList -EnableAttachments $true`. Toggling on in an existing app adds the box below the form; when there is no room it offers to rebuild that screen in place (same id and name). Play shows the files and lets you attach and remove them. The form names carry the list (`frmAttachPolicies`) because names must be unique app-wide.

**Approval-chain extras** (wizard chain editor, Claude plans and *Ask for a change* `setApprovalStages`; each one opt-in):
- *Deadline* per stage (days). Adds a `StageSince` date the app sets when an item reaches a stage; the queue shows an Overdue badge, details say "overdue", and a daily reminder flow per stage emails whoever the stage waits on (`CurrentStage eq n and StageSince lt today − days`).
- *Approvers can reassign.* Adds a `<Stage>Delegate` Person column per stage. The approver (or an admin) picks someone in the details screen's *Options* overlay; that person can then decide the stage (queue and permission checks include the stand-in; the approver column is kept) and the history records "Reassigned to …".
- *Admins can override* (needs the admin group). Admins approve or reject the waiting stage from the same overlay with a required reason; the history line says "(admin override)".
With none of them on, chain output is unchanged.

**Polish.** The lint report has one-click size actions: *Mark List large* on PS035 and *List stays small* on PS059. Version history compares any two versions (or one with the app now) and shows a line diff of each changed screen's YAML. Versions now live in IndexedDB (or their own localStorage key where IndexedDB is unavailable), not inside the saved project; old projects migrate on first load, and the project only drops its in-project copy after the versions are stored. Project files and bundles still carry them. Opening a bundle offers *Add to my app*: its lists, screens and flows join the open app, a same-named list with different columns and every clashing screen or control name get a number, and the formulas that name them follow.

**Deliberate output changes.** Details-screen section titles (and the chain's History title) now hide with their cards; before, they floated over the "nothing selected" state. Edit forms on phones too long for the screen shorten their multi-line boxes instead of running under the Save bar (no shipped scenario before round 6 was affected). Everything else is byte-identical to round 5 when the new features are off.

**Needs a tenant test:** the attachments form pattern end to end with these names, the reassign Patch of a Person column from a people-picker variable, the `With({sla: …})` overdue formulas in a gallery, and the reminder flow's OData filter on `StageSince`.

## Provisioning script: PnP or GCC High / DoD

Get YAML → *Provisioning script* has two flavours. They create the same lists and columns, in the same order (lookup targets first), and both are safe to re-run (existing lists and columns print `[exists]` and are left alone).

- **PnP PowerShell** (`Provision-Lists.ps1`): `Connect-PnPOnline -Interactive` and `Add-PnPFieldFromXml`. For commercial and GCC tenants.
- **GCC High / DoD** (`Provision-Lists-GCCHigh.txt`): no PnP. It's a self-contained script you open in Notepad and paste into a PowerShell window:
  - It loads MSAL.PS from a local path (`$MsalModulePath`, default `C:\eclipse\msal.ps`), which avoids the execution-policy error.
  - `Connect-DoD365Interactive` signs in with `Get-MsalToken -Interactive` (CAC/PIV) against `https://login.microsoftonline.us`, with redirect URI `urn:ietf:wg:oauth:2.0:oob`.
  - It then calls SharePoint REST with a request digest. Lists are created through `/_api/web/lists` (BaseTemplate 100), columns through `fields/createfieldasxml` with the same field XML as the PnP flavour (Options 8), and attachments and optional-Title settings through MERGE calls.
  - The EDIT ME block holds the tenant ID, client ID, SharePoint resource and site URL. The resource and site are filled in from *Finish setup → Site URL* when it's set.

Which flavour you get by default depends on the site URL: one on `sharepoint.us` or `sharepoint-mil.us` gives GCC High, anything else gives PnP. Clicking a flavour button saves the choice with the project (`app.scriptKind`). The paste assistant's *SharePoint lists* step and the zip use the same choice.
