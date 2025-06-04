/* global console, document, fabric, Element, HTMLButtonElement, HTMLElement, HTMLFormElement, HTMLInputElement, HTMLSelectElement, NodeListOf, Office */
import { outlook2016CompatMode } from "../compat";
import { localizeDocument } from "../i18n";
import { ReportAction, Settings, Transport } from "../models";
import { getDefaults, getSettings, setSettings } from "../settings";
import { fixOWAPadding, showView } from "../utils";

class OptionsForm {
  advancedElements: NodeListOf<Element>;
  expressiveSubjectCheckbox: any; // fabric CheckBox component
  form: HTMLFormElement;
  httpElements: NodeListOf<Element>;
  lucyClientIDInput: HTMLInputElement;
  lucyServerInput: HTMLInputElement;
  phishingTransportDropdown: HTMLElement;
  simulationTransportDropdown: HTMLElement;
  reportActionDropdown: HTMLElement;
  resetButton: HTMLButtonElement;
  smtpElements: NodeListOf<Element>;
  smtpToInput: HTMLInputElement;
  toggleAdvancedCheckbox: any; // fabric CheckBox component

  constructor() {
    this.advancedElements = document.querySelectorAll(".mailreport-advanced");
    this.form = document.querySelector("#mailreport-options form");
    this.httpElements = document.querySelectorAll(".mailreport-http");
    this.lucyClientIDInput = <HTMLInputElement>document.getElementById("mailreport-lucy_client_id");
    this.lucyServerInput = <HTMLInputElement>document.getElementById("mailreport-http_lucy_server");
    this.phishingTransportDropdown = <HTMLSelectElement>document.getElementById("mailreport-phishing_transport");
    this.simulationTransportDropdown = <HTMLSelectElement>document.getElementById("mailreport-simulation_transport");
    this.reportActionDropdown = <HTMLSelectElement>document.getElementById("mailreport-report_action");
    this.resetButton = <HTMLButtonElement>document.getElementById("mailreport-options-reset");
    this.smtpElements = document.querySelectorAll(".mailreport-smtp");
    this.smtpToInput = <HTMLInputElement>document.getElementById("mailreport-smtp_to");
  }

  initialize() {
    const dropdownElements = document.querySelectorAll(".ms-Dropdown"),
      $toggleAdvancedCheckbox = document.getElementById("mailreport-show_advanced"),
      $toggleExpressiveSubjectCheckbox = document.getElementById("mailreport-toggle_expressive_subject");
    // Initialize fabric components
    this.expressiveSubjectCheckbox = new fabric["CheckBox"]($toggleExpressiveSubjectCheckbox);
    this.toggleAdvancedCheckbox = new fabric["CheckBox"]($toggleAdvancedCheckbox);
    dropdownElements.forEach((e) => new fabric["Dropdown"](e));
    // Update form field visibility when toggling checkbox that show or hide elements
    const visibilityChangingElements = [
      this.phishingTransportDropdown.querySelector("select"),
      this.simulationTransportDropdown.querySelector("select"),
      this.toggleAdvancedCheckbox._choiceInput,
    ];
    visibilityChangingElements.forEach((e) => {
      e.addEventListener("change", () => updateFormFields(this));
    });
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
      if (outlook2016CompatMode()) {
        showView("#mailreport-options-close");
        return;
      }
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
    settings.lucy_client_id = form.lucyClientIDInput.value !== "" ? parseInt(form.lucyClientIDInput.value) : null;
    settings.lucy_server = form.lucyServerInput.value;
    settings.phishing_transport = getDropdownValue(form.phishingTransportDropdown) as Transport;
    settings.simulation_transport = getDropdownValue(form.simulationTransportDropdown) as Transport;
    settings.smtp_to = form.smtpToInput.value;
    settings.smtp_use_expressive_subject = form.expressiveSubjectCheckbox.getValue();
  }
  return settings;
}

/**
 * Restores all form fields from the given settings object.
 */
function restoreFormSettings(form: OptionsForm, settings: Settings) {
  form.lucyClientIDInput.value = settings.lucy_client_id !== null ? settings.lucy_client_id.toString() : "";
  form.lucyServerInput.value = settings.lucy_server;
  updateDropdown(form.phishingTransportDropdown, settings.phishing_transport);
  updateDropdown(form.simulationTransportDropdown, settings.simulation_transport);
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
  form.advancedElements.forEach((e) => {
    if (form.toggleAdvancedCheckbox.getValue()) e.classList.remove("hide");
    else e.classList.add("hide");
  });
  // HTTP(S)+SMTP
  let httpEnabled = false,
    smtpEnabled = false;
  [getDropdownValue(form.phishingTransportDropdown), getDropdownValue(form.simulationTransportDropdown)].forEach(
    (t) => {
      httpEnabled = httpEnabled || t === Transport.HTTP || t === Transport.HTTPSMTP;
      smtpEnabled = smtpEnabled || t === Transport.SMTP || t === Transport.HTTPSMTP;
    }
  );
  if (httpEnabled) {
    form.httpElements.forEach((e) => {
      e.classList.remove("hide");
    });
    form.lucyServerInput.required = true;
  } else {
    form.httpElements.forEach((e) => {
      e.classList.add("hide");
    });
    form.lucyServerInput.required = false;
  }
  if (smtpEnabled) {
    form.smtpElements.forEach((e) => {
      e.classList.remove("hide");
    });
    form.smtpToInput.required = true;
  } else {
    form.smtpElements.forEach((e) => {
      e.classList.add("hide");
    });
    form.smtpToInput.required = false;
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
  const settings = getSettings();
  console.log("loaded settings", settings);
  restoreFormSettings(form, settings);
  showPermittedElements(form, settings);
});
