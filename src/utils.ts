/* global document, Office */
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
  if (Office.context.mailbox.diagnostics.hostName === "OutlookWebApp") document.body.style.marginLeft = "8px";
}
