/* global console, Office */
import defaultSettings from "./defaults.json";
import { ReportAction } from "./models";

/**
 * Shared access to the add-in's settings, persisted via Office.RoamingSettings
 */

export class Settings {
  report_action: ReportAction = ReportAction.JUNK;
}

/**
 * Attempts to retrieve settings from Office.RoamingSettings, otherwise returns static defaults.
 */
export function getSettings(): Settings {
  const settings = new Settings();
  for (const key in settings) {
    settings[key] = Office.context.roamingSettings.get(key) ?? defaultSettings[key];
  }
  return settings;
}

/**
 * Persists the given settings in Office.Roamingsettings.
 */
export function setSettings(settings: Settings): void {
  console.log("Saving settings: ", settings);
  for (const key in settings) {
    Office.context.roamingSettings.set(key, settings[key]);
  }
  Office.context.roamingSettings.saveAsync();
}

/**
 * Returns all static default settings.
 */
export function getDefaults(): Settings {
  const settings = new Settings();
  for (const key in settings) {
    settings[key] = defaultSettings[key];
  }
  return settings;
}

/**
 * Returns true if the given settings are equal to the ones set in Office.Roamingsettings.
 * Ignores keys that are only set in Office.RoamingSettings to support default values
 * that shouldn't be modified by users.
 */
export function isEqualToSettings(settings: Settings): boolean {
  const currentSettings = getSettings();
  return Object.entries(settings)
    .map(([k, v]) => {
      return k in currentSettings && currentSettings[k] === v;
    })
    .every(Boolean);
}
