/* global document, HTMLTextAreaElement, Office */
import { localizeToken, localizeDocument } from "../i18n";
import { ReportAction, ReportResult } from "../models";
import { reportFraud } from "../reporting";
import { getSettings } from "../settings";
import { showSimulationAcknowledgement } from "../simulation";
import { fixOWAPadding, sleep } from "../utils";

function showView(selector: string) {
  const $unselected = document.querySelectorAll(`div.view:not(${selector})`);
  document.querySelector(selector).classList.remove("hide");
  for (const e of $unselected) e.classList.add("hide");
}

async function handleFraudReport() {
  showView("#mailreport-fraud-pending");
  const comment = (<HTMLTextAreaElement>document.getElementById("reportComment")).value;
  const reportResult = await reportFraud(Office.context.mailbox.item, comment);
  switch (reportResult) {
    case ReportResult.SUCCESS:
      showView("#mailreport-fraud-success");
      await sleep(2000);
      break;
    case ReportResult.SIMULATION:
      await showSimulationAcknowledgement();
      break;
    case ReportResult.ERROR:
      showView("#mailreport-fraud-error");
      await sleep(5000);
      break;
  }
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
