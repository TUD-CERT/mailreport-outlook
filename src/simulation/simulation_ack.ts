/* global document, Office */
import { localizeDocument } from "../i18n";
import { applyTheme } from "../utils";

Office.onReady(() => {
  localizeDocument();
  applyTheme();

  document.querySelector("button").addEventListener("click", () => {
    Office.context.ui.messageParent("");
  });
});
