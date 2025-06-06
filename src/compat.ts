/* global Office */

export function isOutlook2016(): boolean {
  return !("closeContainer" in Office.context.ui);
}

export function isMacOS(): boolean {
  return Office.context.diagnostics !== undefined && Office.context.diagnostics.platform === Office.PlatformType.Mac;
}

export function isOWA(): boolean {
  return Office.context.mailbox.diagnostics.hostName === "OutlookWebApp";
}
