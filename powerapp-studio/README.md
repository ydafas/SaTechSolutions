# Power App Studio

A browser-only Power Apps canvas app designer for satechitsolutions.com. There is no build step and no server: GitHub Pages serves these files as they are.

## What it does

1. **Requirements wizard**: pick a scenario, shape the SharePoint lists, then choose the screens and a theme. It generates wired screens: home, browse, details, edit and approvals.
2. **Designer**: drag-and-drop canvas, gallery templates, inputs bound to columns, an actions builder, Visible rules, Power Fx overrides, layers, undo, copy and paste, alignment tools.
3. **Theme**: presets, 16 color tokens, font, corner radius. These are emitted as one `appTheme` named formula.
4. **Play**: runs the app against sample data, including search, open, edit, save, delete, approve/reject, variables and people pickers.
5. **Get YAML**: App.Formulas, App.OnStart, App.StartScreen and one paste-ready YAML block per screen. It also produces a lint report, a PnP provisioning script, the project JSON and a zip of everything.
6. **Import**: paste YAML copied from Power Apps Studio (or a `Src/*.pa.yaml` file) to keep editing it here.
8. **Describe your app** (uses Claude): write requirements in plain English. Claude returns a plan of lists, screens, approvals and flows, which you review and refine before the studio builds it. It uses the visitor's own Anthropic API key.
7. **Flows**: Power Automate flows for the app's lists, exported as legacy flow packages. There are four templates: email on a new item, email when a Choice column changes, a daily reminder, and a flow a button runs. The email editor uses `{Column}` tokens and `{Link}`, which deep-links into the app. Play mode shows which flows would fire.

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

These templates passed the harness and the skill's `screen_tool.py check`, but have not yet been pasted into a tenant: the stepper HTML built with `With()`, the confirmation overlay on a multi-screen app, the Mine/All `If` branches over `SortByColumns(Filter(...))`, the loading-screen emblem data URI with the timer step machine, and `appTheme.FontCss` inside HtmlViewer styles. The admin check (`mail` field) and the AppUsage block are copied from verified skill patterns.

## Files

| File | Role |
|---|---|
| `index.html`, `studio.css` | Shell and styles (SaTech ink and cyan) |
| `js/core.js` | Project model and JSDoc types, control catalog, column types, theme presets, sample data, `renameCopy` (unique names for copies), `normalizeProject` (makes loaded and shared projects safe) |
| `js/fx.js` | Power Fx builders for the field-tested shapes: `ifErrorFx`, `upsertBaseFx`, `rebaseFx`, `spUserFx` / `meUserFx`, `recordFx`, `equalsFx`, `notifyFx`, `navigateFx` |
| `js/emit.js` | YAML and Power Fx generation, deploy checklist, PnP script, lint (`LINT_RULES`) |
| `js/render.js` | Control rendering, HTML sanitizer, mini Power Fx evaluator, Play runtime |
| `js/templates.js` | Scenarios, screen templates, insertable blocks |
| `js/import.js` | YAML import with name-clash renaming |
| `js/ui.js` | Designer UI: canvas, panels, inspector, dialogs, Play |
| `js/flows.js` | Flow model and templates, Logic Apps definitions, legacy package files (`flowPackageFiles`), tokens to expressions, deep link (`deepLinkInfo`), flow lint |
| `js/ai.js` | Describe-your-app: plan JSON schema, system prompt, browser call to the Claude API, `specToProject` (cleans the plan and builds the project, wiring button flows and deep links) |
| `js/describeui.js` | The Describe dialog: requirements, key, model, review, refine, build |
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
