/* global console, document, fabric, HTMLElement, HTMLInputElement, HTMLSelectElement, Office */
import { localizeDocument } from "../i18n";
import { ReportAction } from "../models";
import { getDefaults, getSettings, setSettings, Settings } from "../settings";
import { fixOWAPadding } from "../utils";

/**
 * Updates the currently selected value of a fabric <select> element.
 * Based on https://github.com/OfficeDev/office-ui-fabric-js/issues/331
 */
function updateSelect(selectElement: HTMLElement, newValue: string) {
  const text = selectElement.querySelector(`option[value="${newValue}"]`).textContent;
  selectElement.querySelectorAll("li").forEach((e) => {
    if (e.textContent === text) e.classList.add("is-selected");
    else e.classList.remove("is-selected");
  });
  selectElement.querySelector(".ms-Dropdown-title").textContent = text;
  selectElement.querySelector("select").value = newValue;
}

/**
 * Returns a settings object created from the currently selected form values.
 * Takes into account the current permission configuration: If advanced
 * configuration is disabled, only basic config keys/values are returned.
 */
function getFormSettings(currentSettings: Settings): Settings {
  const settings = new Settings();
  settings.report_action = (<HTMLSelectElement>document.getElementById("mailreport-report_action_select"))
    .value as ReportAction;
  if (currentSettings.permit_advanced_config) {
    settings.smtp_to = (<HTMLInputElement>document.getElementById("mailreport-smtp_to")).value;
  }
  return settings;
}

/**
 * Restores all form fields from the given settings object.
 */
function restoreFormSettings(settings: Settings) {
  updateSelect(document.getElementById("mailreport-report_action"), settings.report_action);
  (<HTMLInputElement>document.getElementById("mailreport-smtp_to")).value = settings.smtp_to;
  updateFormFields();
}

/**
 * Shows or hides form fields depending on the currently selected settings.
 * Also adds or removes 'required' attributes depending on the selected fields.
 */
function updateFormFields() {
  // Advanced settings
  let showAdvancedSettings = (<HTMLInputElement>document.querySelector('input[name="mailreport-advanced"]')).checked;
  const advancedElements = document.querySelectorAll(".mailreport-advanced");
  for (let i = 0; i < advancedElements.length; i++) {
    const $element = advancedElements[i];
    if (showAdvancedSettings) $element.classList.remove("hide");
    else $element.classList.add("hide");
  }
}

/**
 * Updates visibility of various options according to permission configuration.
 */
function showPermittedElements(settings: Settings) {
  let $showAdvancedCheckbox = document.getElementById("mailreport-show_advanced");
  if (settings.permit_advanced_config) $showAdvancedCheckbox.classList.remove("hide");
  else $showAdvancedCheckbox.classList.add("hide");
}

Office.onReady(() => {
  localizeDocument();
  fixOWAPadding();
  const dropdownHTMLElements = document.querySelectorAll(".ms-Dropdown"),
    checkboxHTMLElements = document.querySelectorAll(".ms-CheckBox"),
    $resetButton = document.getElementById("mailreport-options-reset"),
    $form = document.querySelector("#mailreport-options form"),
    visibilityChangingHTMLElements = document.querySelectorAll('input[type="checkbox"]');

  for (var i = 0; i < dropdownHTMLElements.length; ++i) {
    new fabric["Dropdown"](dropdownHTMLElements[i]);
  }
  for (i = 0; i < checkboxHTMLElements.length; ++i) {
    new fabric["CheckBox"](checkboxHTMLElements[i]);
  }
  new fabric["Button"]($resetButton, () => {
    const defaultSettings = getDefaults();
    restoreFormSettings(defaultSettings);
    console.log("restored default settings", defaultSettings);
  });
  $form.addEventListener("submit", (e) => {
    e.preventDefault();
    const settings = getFormSettings(getSettings());
    setSettings(settings);
    console.log("saved settings", settings);
    Office.context.ui.closeContainer();
  });
  for (let i = 0; i < visibilityChangingHTMLElements.length; i++) {
    visibilityChangingHTMLElements[i].addEventListener("change", updateFormFields);
  }

  const settings = getSettings();
  console.log("loaded settings", settings);
  restoreFormSettings(settings);
  showPermittedElements(settings);
});
