# Power App Studio

A browser-only Power Apps canvas app designer for satechitsolutions.com. There is no build step and no server: GitHub Pages serves these files as they are.

## What it does

1. **Requirements wizard**: pick a scenario, shape the SharePoint lists, then choose the screens and a theme. It generates wired screens: home, browse, details, edit and approvals.
2. **Designer**: drag-and-drop canvas, gallery templates, inputs bound to columns, an actions builder, Visible rules, Power Fx overrides, layers, undo, copy and paste, alignment tools.
3. **Theme**: presets, 16 color tokens, font, corner radius. These are emitted as one `appTheme` named formula.
4. **Play**: runs the app against sample data, including search, open, edit, save, delete, approve/reject, variables and people pickers.
5. **Get YAML**: App.Formulas, App.OnStart, App.StartScreen and one paste-ready YAML block per screen. It also produces a lint report, a PnP provisioning script, the project JSON and a zip of everything.
6. **Import**: paste YAML copied from Power Apps Studio (or a `Src/*.pa.yaml` file) to keep editing it here.

## Files

| File | Role |
|---|---|
| `index.html`, `studio.css` | Shell and styles (SaTech ink and cyan) |
| `js/core.js` | Project model, control catalog, column types, theme presets, sample data |
| `js/emit.js` | YAML and Power Fx generation, deploy checklist, PnP script, lint |
| `js/render.js` | Control rendering, mini Power Fx evaluator, Play runtime |
| `js/templates.js` | Scenarios, screen templates, insertable blocks |
| `js/import.js` | YAML import with name-clash renaming |
| `js/ui.js` | Designer UI: canvas, panels, inspector, dialogs, Play |
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

To add a lint rule, add a check in `lint()` in `js/emit.js`.

## Local testing

ES modules do not load from `file://`. Serve the folder instead, for example `python -m http.server`, then open `/powerapp-studio/`.
