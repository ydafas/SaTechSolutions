// fx.js — small Power Fx builders shared by the emitter.
// Each shape here is one that pasted and ran cleanly in Power Apps Studio (commercial and GCC High).
// Keep them pure string builders: no project lookups, so they are easy to test and reuse.
import { s } from './core.js';

/** Notify("text", NotificationType.Kind) */
export const notifyFx = (text, kind = 'Information') => `Notify(${s(text)}, NotificationType.${kind})`;

/** Navigate(screen, ScreenTransition.X) */
export const navigateFx = (screenName, transition = 'Fade') => `Navigate(${screenName}, ScreenTransition.${transition || 'Fade'})`;

/** Fresh copy of the selected row. Never Patch a cached gallery/variable record (flows race it). */
export const rebaseFx = (list, selVar) => `LookUp(${list}, ID = ${selVar}.ID)`;

/** Existing row if one is selected, otherwise a new one: the create-or-update base for Patch. */
export const upsertBaseFx = (list, selVar) => `Coalesce(${rebaseFx(list, selVar)}, Defaults(${list}))`;

/**
 * 3-argument IfError: the last argument runs only when nothing errored.
 * IfError(value, Notify("<verb> failed. " & FirstError.Message, Error), success1; success2; ...)
 */
export function ifErrorFx(value, verb, onSuccess) {
  return `IfError(\n    ${value},\n    Notify("${verb} failed. " & FirstError.Message, NotificationType.Error),\n    ${onSuccess.join(';\n    ')}\n)`;
}

/** SharePoint Person column value for Patch. emailExpr / nameExpr are Power Fx expressions. */
export const spUserFx = (emailExpr, nameExpr) =>
  `{'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser", Claims: "i:0#.f|membership|" & Lower(${emailExpr}), DisplayName: ${nameExpr}, Email: ${emailExpr}, Department: "", JobTitle: "", Picture: ""}`;

/** Person value for the signed-in user. */
export const meUserFx = () => spUserFx('User().Email', 'User().FullName');

/**
 * Multi-line record literal: [[key, valueFx], ...] → "{\n        Key: value,\n    }".
 * Keys must already be quoted with q() where needed.
 */
export function recordFx(pairs, indent = '    ') {
  if (!pairs.length) return '{}';
  return '{\n' + pairs.map(([k, v]) => `${indent}    ${k}: ${v}`).join(',\n') + `\n${indent}}`;
}

/** Turn a column-equals test into a filter predicate for a column type. */
export function equalsFx(colRef, colType, value) {
  switch (colType) {
    case 'Choice': return `${colRef}.Value = ${s(value)}`;
    case 'YesNo': return `${colRef} = ${/^(true|yes|1)$/i.test(String(value)) ? 'true' : 'false'}`;
    case 'Number': case 'Currency': return `${colRef} = ${Number(value) || 0}`;
    case 'Person': return `${colRef}.Email = User().Email`;
    default: return `${colRef} = ${s(value)}`;
  }
}
