/* global document, Office, setTimeout */
/**
 * Encodes various characters to their safe HTML counterparts. Used to prevent HTML interpretation of
 * E-Mail headers such as "Name <name@example.com>".
 */
export function encodeHTML(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * In case we're running inside Outlook Web App (OWA), adds some
 * padding to the task pane content to match the padding of the desktop version.
 */
export function fixOWAPadding() {
  if (Office.context.mailbox.diagnostics.hostName === "OutlookWebApp")
    document.documentElement.style.marginLeft = "8px";
}

/**
 * Pauses execution for a set amout of time.
 */
export async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Returns a string representation of the given object as "key1: val1<delimiter>key2: val2<delimiter>...".
 * The result doesn't end with a delimiter.
 */
export function objToStr(source: { [key: string]: string }, delimiter: string) {
  let result = "";
  for (let key in source) result += `${key}: ${source[key]}${delimiter}`;
  return result;
}
