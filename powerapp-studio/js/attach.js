// attach.js — list attachments, the skill's proven pattern (SKILL.md, "Rebuilding a SharePoint list form as a
// standalone app", step 2). Attachments cannot be pasted, so the generated screen leaves a labelled placeholder
// where the user inserts an Edit form by hand. On edit screens the formulas then use it: the loader calls
// NewForm/EditForm and Save merges its .Updates as an extra Patch argument. On detail screens the hand-added
// form is View only and no formula refers to it. Formulas only name a form while the list's Attachments toggle
// is on AND its placeholder is on the screen, so an app with no attachments is unchanged. Pure; no DOM.
import { slug, walk, varSel } from './core.js';

/** The hand-added form's name: frmAttach<List> on edit screens, frmAttachView<List> on detail screens. */
export const attachFormName = (list, mode) => (mode === 'view' ? 'frmAttachView' : 'frmAttach') + slug(list);
export const FORM_RE = /\bfrmAttach(?:View)?[A-Za-z0-9_]*/;
export const attachOn = (project, list) => !!(project.lists || []).find((l) => l.name === list)?.attachments;

/** Placeholders on a screen whose list has Attachments on: [{ctl, list, mode, form}]. `mode` filters. */
export function attachSlots(project, screen, mode = null) {
  const out = [];
  for (const [c, parent] of walk(screen.controls || [])) {
    const a = c.attachSlot;
    if (parent || !a || !attachOn(project, a.list) || (mode && a.mode !== mode)) continue;
    out.push({ ctl: c, list: a.list, mode: a.mode, form: attachFormName(a.list, a.mode) });
  }
  return out;
}
/** Every active placeholder in the app (for the Paste assistant and lint). */
export const allAttachSlots = (project) => project.screens.flatMap((scr) => attachSlots(project, scr).map((x) => ({ ...x, screen: scr })));

/** Static placeholder text: a dashed box that says what to insert. Fits HtmlViewer's 5px padding. */
export function placeholderHtml(list, mode, h) {
  const form = attachFormName(list, mode);
  const inner = Math.max(40, h - 10);
  return `<div style='box-sizing:border-box;height:${inner}px;border:2px dashed #94A3B8;border-radius:8px;padding:10px 14px;font-size:12px;color:#475569'>`
    + `<b style='font-size:13px;color:#1E293B'>Attachments</b><br>Insert an Edit form named <b>${form}</b> here after pasting `
    + `(Paste assistant, Attachments step).${mode === 'view' ? ' It only shows files.' : ' Save keeps the files it holds.'}</div>`;
}

/** Edit-screen loader line (in OnVisible): a new item gets an empty form, an existing one its files. */
export const attachLoaderFx = (slot) => `If(IsBlank(${varSel(slot.list)}.ID), NewForm(${slot.form}), EditForm(${slot.form}))`;

/** Paste assistant instructions for one screen's placeholders (exact names and property values). */
export function attachHow(screenName, slots) {
  const lines = [`Attachments cannot be pasted, so add them by hand on ${screenName}. Do this after the screen is pasted, every time you paste it again.`];
  slots.forEach((x, i) => {
    const V = varSel(x.list);
    lines.push('', `${slots.length > 1 ? `${i + 1}. ` : ''}${x.mode === 'view' ? 'View-only form' : 'Edit form'} where the dashed "Attachments" box (${x.ctl.name}) is:`,
      '• Insert → Edit form. Rename it to ' + x.form + ' (exactly; ' + (x.mode === 'view' ? 'nothing refers to it, but keep the name for re-pastes).' : 'Save and the screen loader refer to it).'),
      `• Data source: ${x.list}. Edit fields: remove every field except Attachments.`,
      `• Item: =${V}`,
      `• DefaultMode: =${x.mode === 'view' ? 'FormMode.View' : 'FormMode.Edit'}`,
      `• Select the Attachments card → Advanced → Unlock to change properties → DisplayMode: =${x.mode === 'view' ? 'DisplayMode.View' : 'DisplayMode.Edit'}`,
      `• Move and size it over the dashed box (X ${x.ctl.x}, Y ${x.ctl.y}, Width ${x.ctl.w}, Height ${x.ctl.h}), then bring it to the front.`);
  });
  lines.push('', 'Before you delete this screen to paste a new version: select each form, Ctrl+C, and keep it in Notepad. Paste it back onto the new screen afterwards (or redo these steps).');
  return lines.join('\n');
}
/** Property values to copy (the Copy button of the Paste assistant step). */
export const attachCopyText = (slots) => slots.map((x) => [`${x.form}`, `  Item: =${varSel(x.list)}`, `  DefaultMode: =${x.mode === 'view' ? 'FormMode.View' : 'FormMode.Edit'}`, `  Attachments card DisplayMode: =${x.mode === 'view' ? 'DisplayMode.View' : 'DisplayMode.Edit'}`].join('\n')).join('\n\n');

