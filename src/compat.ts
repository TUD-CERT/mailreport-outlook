/* global Office */

export function outlook2016CompatMode(): boolean {
  return !("closeContainer" in Office.context.ui);
}
