/* global document, HTMLTextAreaElement, Office */
import { localizeToken, localizeDocument } from "../i18n";
import { ReportAction } from "../models";
import { reportFraud } from "../reporting";
import { getSettings } from "../settings";
import { fixOWAPadding, sleep } from "../utils";

function showView(selector: string) {
  const $unselected = document.querySelectorAll(`div.view:not(${selector})`);
  document.querySelector(selector).classList.remove("hide");
  for (const e of $unselected) e.classList.add("hide");
}

async function handleFraudReport() {
  showView("#mailreport-fraud-pending");
  const comment = (<HTMLTextAreaElement>document.getElementById("reportComment")).value;
  const success = await reportFraud(Office.context.mailbox.item, comment);
  showView(success ? "#mailreport-fraud-success" : "#mailreport-fraud-error");
  await sleep(success ? 2000 : 5000);
  Office.context.ui.closeContainer();
}

Office.onReady((info) => {
  localizeDocument();
  // Display the reporting action depending on current settings
  const $reportAction = document.getElementById("mailreport-fraud-action");
  switch (getSettings().report_action) {
    case ReportAction.JUNK:
      $reportAction.textContent = localizeToken("__MSG_reportCommentJunk__");
      break;
    case ReportAction.TRASH:
      $reportAction.textContent = localizeToken("__MSG_reportCommentTrash__");
      break;
  }
  fixOWAPadding();
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sendFraudReport").onclick = handleFraudReport;
  }
});
