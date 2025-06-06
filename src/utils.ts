/* global document, localStorage, Office, setTimeout */
import { OfficeThemeId } from "./models";

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

/**
 * Shows a div.view defined by selector and hides all other div.view elements.
 * Used to switch between multiple "views" within a single document.
 */
export function showView(selector: string) {
  const $unselected = document.querySelectorAll(`div.view:not(${selector})`);
  document.querySelector(selector).classList.remove("hide");
  for (const e of $unselected) e.classList.add("hide");
}

/**
 * If the Office.OfficeTheme interface is supported,
 * update CSS styles with the currently selected theme.
 */
export function applyTheme() {
  // Use localStorage as cache to pass the currently selected theme to dialogs
  const cachedTheme = localStorage.getItem("mailreport-theme");
  let theme = Office.context.officeTheme;
  if (theme === undefined) {
    if (cachedTheme === null) return; // No OfficeTheme interface support
    theme = JSON.parse(cachedTheme);
  }
  localStorage.setItem("mailreport-theme", JSON.stringify(theme)); // Update cache

  // Contrasting official docs, themeId is a string and Office.ThemeId is undefined on Outlook 2021 LTSC and 2024 LTSC.
  const selectedTheme = theme.themeId as unknown as OfficeThemeId;

  // Always set a solid background color to fix rendering issues on Outlook 2024 LTSC, which has a transparent background by default.
  document.querySelector("body").style.backgroundColor = theme.bodyBackgroundColor;

  if (selectedTheme === OfficeThemeId.Colorful || selectedTheme === OfficeThemeId.White) return;

  document.querySelectorAll("label, label > span, p").forEach((e) => {
    e.style.color = theme.bodyForegroundColor;
  });
  document.querySelectorAll("div.view, h1, span").forEach((e) => {
    switch (selectedTheme) {
      case OfficeThemeId.DarkGray:
        e.classList.replace("ms-fontColor-themeDarker", "ms-fontColor-themeLight");
        e.classList.replace("ms-fontColor-greenDark", "ms-fontColor-greenLight");
        e.classList.replace("ms-fontColor-redDark", "ms-fontColor-orangeLighter");
        e.classList.replace("ms-fontColor-black", "ms-fontColor-white");
        break;
      case OfficeThemeId.Black:
        e.classList.replace("ms-fontColor-themeDarker", "ms-fontColor-themeTertiary");
        e.classList.replace("ms-fontColor-greenDark", "ms-fontColor-green");
        e.classList.replace("ms-fontColor-redDark", "ms-fontColor-red");
        e.classList.replace("ms-fontColor-black", "ms-fontColor-white");
        break;
    }
  });
}
