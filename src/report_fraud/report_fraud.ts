/* global document, HTMLParagraphElement, HTMLTextAreaElement, Office */
import { isOutlook2016 } from "../compat";
import { moveMessageTo } from "../ews";
import { localizeToken, localizeDocument } from "../i18n";
import { MoveMessageStatus, ReportAction, ReportResultStatus } from "../models";
import { reportFraud } from "../reporting";
import { getSettings } from "../settings";
import { showSimulationAcknowledgement } from "../simulation";
import { applyTheme, fixTaskPanePadding, showView, sleep } from "../utils";

async function handleFraudReport() {
  showView("#mailreport-fraud-pending");
  const comment = (<HTMLTextAreaElement>document.getElementById("reportComment")).value,
    mail = Office.context.mailbox.item,
    reportResult = await reportFraud(mail, comment);
  switch (reportResult.reportStatus) {
    case ReportResultStatus.SUCCESS:
      showView("#mailreport-fraud-success");
      await sleep(2000);
      break;
    case ReportResultStatus.SIMULATION:
      await showSimulationAcknowledgement();
      break;
    case ReportResultStatus.ERROR:
      showView("#mailreport-fraud-error");
      (<HTMLParagraphElement>document.querySelector("#mailreport-fraud-error-diag")).textContent =
        reportResult.diagnosis;
      return; // Do not close this view automatically
  }
  // Move message after we have slept or the user has closed the sim ack dialog.
  // If we had moved it earlier, the task pane and dialog would close immediately.
  if (reportResult.moveMessageStatus === MoveMessageStatus.PENDING)
    reportResult.moveMessageStatus = await moveMessageTo(mail, reportResult.moveMessageTarget);
  if (isOutlook2016()) {
    showView("#mailreport-fraud-close");
    return;
  }
  Office.context.ui.closeContainer();
}

Office.onReady((info) => {
  localizeDocument();
  applyTheme();

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
  fixTaskPanePadding();
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sendFraudReport").onclick = handleFraudReport;
  }
});
