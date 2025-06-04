/* global document, HTMLTextAreaElement, Office */
import { outlook2016CompatMode } from "../compat";
import { localizeToken, localizeDocument } from "../i18n";
import { ReportAction, ReportResult } from "../models";
import { reportFraud } from "../reporting";
import { getSettings } from "../settings";
import { showSimulationAcknowledgement } from "../simulation";
import { fixOWAPadding, showView, sleep } from "../utils";

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

  if (outlook2016CompatMode()) {
    showView("#mailreport-fraud-close");
    return;
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
