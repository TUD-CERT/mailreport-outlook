/* global document, Office */
import { localizeDocument } from "../i18n";

Office.onReady(() => {
  localizeDocument();

  document.querySelector("button").addEventListener("click", () => {
    Office.context.ui.messageParent("");
  });
});
