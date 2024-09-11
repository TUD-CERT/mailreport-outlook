/* global console, document, fabric, Office */
import { getSettings, ReportAction, setSettings, Settings } from "../settings";

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
 * 
 * Returns a settings object created from the currently selected form values.
 */
function getFormSettings(): Settings {
  const settings = new Settings();
  settings.report_action = (<HTMLSelectElement>document.querySelector(`#mailreport-report_action select`))
    .value as ReportAction;
  return settings;
}

/**
 * Restores all form fields from the given settings object.
 */
function restoreFormSettings(settings: Settings) {
  updateSelect(document.getElementById("mailreport-report_action"), settings.report_action);
}

Office.onReady(() => {
  const dropdownHTMLElements = document.querySelectorAll(".ms-Dropdown"),
    submitButon = document.getElementById("mailreport-options-save");

  for (var i = 0; i < dropdownHTMLElements.length; ++i) {
    new fabric["Dropdown"](dropdownHTMLElements[i]);
  }
  new fabric["Button"](submitButon, () => {
    const settings = getFormSettings();
    setSettings(settings);
    console.log("saved settings", settings);
    Office.context.ui.closeContainer();
  });

  const settings = getSettings();
  console.log("loaded settings", settings);
  restoreFormSettings(settings);
});
