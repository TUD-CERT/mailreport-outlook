/* global console, Office */
import defaultSettings from "./defaults.json";
import { Settings } from "./models";

/**
 * Attempts to retrieve settings from Office.RoamingSettings, otherwise returns static defaults.
 */
export function getSettings(): Settings {
  const settings = new Settings();
  for (const key in defaultSettings) {
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
  for (const key in defaultSettings) {
    settings[key] = defaultSettings[key];
  }
  return settings;
}
