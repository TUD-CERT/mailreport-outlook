/* global document, localStorage, Office, setTimeout, window */
import { OfficeThemeId, Settings } from "./models";
import { isMacOS, isOWA } from "./compat";

/**
 * Encodes various characters to their safe HTML counterparts. Used to prevent HTML interpretation of
 * E-Mail headers such as "Name <name@example.com>".
 */
export function encodeHTML(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * In case we're running inside Outlook Web App (OWA) or on Mac, add some
 * padding to the task pane content to match the padding of the Windows desktop versions.
 */
export function fixTaskPanePadding() {
  if (isOWA() || isMacOS()) document.documentElement.style.marginLeft = "8px";
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
 * Depending on availability of the Office.OfficeTheme interface and running OS,
 * update DOM styles based on the currently selected theme.
 */
export function applyTheme() {
  const $body = document.querySelector("body");

  let theme = Office.context.officeTheme;
  // Outlook on MacOS sets theme to '{}' or undefined (in dialogs) and
  // renders in a light/dark style based on OS configuration.
  // Therefore, build matching themes manually.
  // Since fabric dropdowns on MacOS improperly overwrite styles on <body>,
  // set the background via a CSS class. On all other platforms, set body.background-color
  // dynamically to a value provided by the Office.OfficeTheme interface.
  if (isMacOS()) {
    if (window.matchMedia("(prefers-color-scheme: dark)").matches) {
      theme = {
        themeId: OfficeThemeId.Black as unknown as Office.ThemeId,
        bodyForegroundColor: "#acacac",
        bodyBackgroundColor: null,
        controlBackgroundColor: null,
        controlForegroundColor: null,
        isDarkTheme: true,
      };
      $body.classList.add("macos-dark");
    } else {
      theme = {
        themeId: OfficeThemeId.White as unknown as Office.ThemeId,
        bodyForegroundColor: null,
        bodyBackgroundColor: null,
        controlBackgroundColor: null,
        controlForegroundColor: null,
        isDarkTheme: false,
      };
      $body.classList.add("macos-light");
    }
  } else {
    // Use localStorage as cache to pass the currently selected app theme to dialogs.
    const cachedTheme = localStorage.getItem("mailreport-theme");
    if (theme === undefined) {
      if (cachedTheme === null) return; // No OfficeTheme interface support
      theme = JSON.parse(cachedTheme); // We are rendering a dialog
    }
    localStorage.setItem("mailreport-theme", JSON.stringify(theme)); // Update cache
  }

  // Contrasting official docs, themeId is a string and Office.ThemeId is undefined on Outlook 2021 LTSC and 2024 LTSC.
  const selectedTheme = theme.themeId as unknown as OfficeThemeId;

  // Always set a solid background color to fix rendering issues on Outlook 2024 LTSC, which has a transparent background by default.
  $body.style.backgroundColor = theme.bodyBackgroundColor;

  // Colorful and White themes don't require adjustments
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

/**
 * Returns an object with telemetry headers to send with requests.
 * Whether any headers are returned depends on the current plugin settings.
 */
export function generateTelemetryHeaders(settings: Settings) {
  const headers: { [key: string]: any } = {};
  if (settings.send_telemetry) {
    const platform =
      Office.context.diagnostics === undefined
        ? ""
        : ` @ ${Office.context.diagnostics.platform === Office.PlatformType.PC ? "Windows" : Office.context.diagnostics.platform}`;
    headers["Reporting-Agent"] =
      `${Office.context.mailbox.diagnostics.hostName}/${Office.context.mailbox.diagnostics.hostVersion}${platform}`;
    headers["Reporting-Plugin"] = settings.plugin_id;
  }
  return headers;
}
