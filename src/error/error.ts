/* global document, HTMLParagraphElement, Office */
import { localizeDocument } from "../i18n";
import { applyTheme } from "../utils";
import URI from "urijs";

Office.onReady(() => {
  localizeDocument();
  applyTheme();

  (<HTMLParagraphElement>document.querySelector("#mailreport-error-diag")).textContent = new URI().search(true).diag;
  document.querySelector("button").addEventListener("click", () => {
    Office.context.ui.messageParent("");
  });
});
