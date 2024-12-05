/* global console, Office */
import defaultSettings from "./defaults.json";
import { Settings } from "./models";

/**
 * Attempts to retrieve settings from Office.RoamingSettings, otherwise returns static defaults.
 * Caution: Due to roamingSettings not offering an API to verify the existence of keys, non-existing
 * keys can't be differentiated from existing keys with a value of null. Therefore, keys set
 * to null in roamingSettings will always be replaced by their default value from defaultSettings.
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
