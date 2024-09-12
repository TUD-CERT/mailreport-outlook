/**
 * Encodes various characters to their safe HTML counterparts. Used to prevent HTML interpretation of
 * E-Mail headers such as "Name <name@example.com>".
 */
export function encodeHTML(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
