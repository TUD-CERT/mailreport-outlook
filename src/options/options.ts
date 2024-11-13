/* global console, document, fabric, Element, HTMLButtonElement, HTMLElement, HTMLFormElement, HTMLInputElement, HTMLSelectElement, NodeListOf, Office */
import { localizeDocument } from "../i18n";
import { ReportAction, Settings } from "../models";
import { getDefaults, getSettings, setSettings } from "../settings";
import { fixOWAPadding } from "../utils";

class OptionsForm {
  advancedElements: NodeListOf<Element>;
  expressiveSubjectCheckbox: any; // fabric CheckBox component
  form: HTMLFormElement;
  reportActionDropdown: HTMLElement;
  resetButton: HTMLButtonElement;
  smtpToInput: HTMLInputElement;
  toggleAdvancedCheckbox: any; // fabric CheckBox component

  constructor() {
    this.advancedElements = document.querySelectorAll(".mailreport-advanced");
    this.form = document.querySelector("#mailreport-options form");
    this.reportActionDropdown = <HTMLSelectElement>document.getElementById("mailreport-report_action");
    this.resetButton = <HTMLButtonElement>document.getElementById("mailreport-options-reset");
    this.smtpToInput = <HTMLInputElement>document.getElementById("mailreport-smtp_to");
  }

  initialize() {
    const dropdownElements = document.querySelectorAll(".ms-Dropdown"),
      $toggleAdvancedCheckbox = document.getElementById("mailreport-show_advanced"),
      $toggleExpressiveSubjectCheckbox = document.getElementById("mailreport-toggle_expressive_subject");
    // Initialize fabric components
    this.expressiveSubjectCheckbox = new fabric["CheckBox"]($toggleExpressiveSubjectCheckbox);
    dropdownElements.forEach((e) => new fabric["Dropdown"](e));
    // Update form field visibility when toggling the advanced options checkbox
    this.toggleAdvancedCheckbox = new fabric["CheckBox"]($toggleAdvancedCheckbox);
    this.toggleAdvancedCheckbox._choiceInput.addEventListener("change", () => updateFormFields(this));
    // Set reset button handler
    new fabric["Button"](this.resetButton, () => {
      const defaultSettings = getDefaults();
      restoreFormSettings(this, defaultSettings);
      console.log("restored default settings", defaultSettings);
    });
    // Set form submission handler
    this.form.addEventListener("submit", (e) => {
      e.preventDefault();
      const settings = getFormSettings(this, getSettings());
      setSettings(settings);
      console.log("saved settings", settings);
      Office.context.ui.closeContainer();
    });
  }
}

/**
 * Updates the currently selected value of a fabric Dropdown component.
 * Based on https://github.com/OfficeDev/office-ui-fabric-js/issues/331
 */
function updateDropdown(dropdownElement: HTMLElement, newValue: string) {
  const text = dropdownElement.querySelector(`option[value="${newValue}"]`).textContent;
  dropdownElement.querySelectorAll("li").forEach((e) => {
    if (e.textContent === text) e.classList.add("is-selected");
    else e.classList.remove("is-selected");
  });
  dropdownElement.querySelector(".ms-Dropdown-title").textContent = text;
  dropdownElement.querySelector("select").value = newValue;
}

/**
 * Returns the value of a Dropdown component from its internal
 * <select> value directly.
 */
function getDropdownValue(dropdownElement: HTMLElement) {
  return (<HTMLSelectElement>dropdownElement.querySelector("select")).value;
}

/**
 * Returns a settings object created from the currently selected form values.
 * Takes into account the current permission configuration: If advanced
 * configuration is disabled, only basic config keys/values are returned.
 */
function getFormSettings(form: OptionsForm, currentSettings: Settings): Settings {
  const settings = new Settings();
  settings.report_action = getDropdownValue(form.reportActionDropdown) as ReportAction;
  if (currentSettings.permit_advanced_config) {
    settings.smtp_to = form.smtpToInput.value;
    settings.smtp_use_expressive_subject = form.expressiveSubjectCheckbox.getValue();
  }
  return settings;
}

/**
 * Restores all form fields from the given settings object.
 */
function restoreFormSettings(form: OptionsForm, settings: Settings) {
  updateDropdown(form.reportActionDropdown, settings.report_action);
  form.smtpToInput.value = settings.smtp_to;
  if (settings.smtp_use_expressive_subject) form.expressiveSubjectCheckbox.check();
  else form.expressiveSubjectCheckbox.unCheck();
  updateFormFields(form);
}

/**
 * Shows or hides form fields depending on the currently selected settings.
 * Also adds or removes 'required' attributes depending on the selected fields.
 */
function updateFormFields(form: OptionsForm) {
  // Advanced settings
  const advancedElements = form.advancedElements;
  for (let i = 0; i < advancedElements.length; i++) {
    const $element = advancedElements[i];
    if (form.toggleAdvancedCheckbox.getValue()) $element.classList.remove("hide");
    else $element.classList.add("hide");
  }
}

/**
 * Updates visibility of various options according to permission configuration.
 */
function showPermittedElements(form: OptionsForm, settings: Settings) {
  let $showAdvancedCheckbox = form.toggleAdvancedCheckbox._container;
  if (settings.permit_advanced_config) $showAdvancedCheckbox.classList.remove("hide");
  else $showAdvancedCheckbox.classList.add("hide");
}

Office.onReady(() => {
  localizeDocument();
  fixOWAPadding();
  const form = new OptionsForm();
  form.initialize();
  console.log(form);
  const settings = getSettings();
  console.log("loaded settings", settings);
  restoreFormSettings(form, settings);
  showPermittedElements(form, settings);
});
