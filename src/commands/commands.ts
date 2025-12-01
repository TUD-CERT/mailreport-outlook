/* global Office, window */
import { moveMessageTo } from "../ews";
import { MoveMessageStatus, ReportResultStatus } from "../models";
import { reportSpam } from "../reporting";
import { showSimulationAcknowledgement } from "../simulation";
import URI from "urijs";

// Must be run each time a new page is loaded.
Office.onReady();

async function showErrorDialog(diagnosis: string) {
  const url = new URI("error.html").addQuery("diag", diagnosis).absoluteTo(window.location).toString();
  const dialogOptions = { width: 60, height: 26, displayInIframe: true };
  return await new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(url, dialogOptions, (asyncResult: Office.AsyncResult<Office.Dialog>) => {
      const dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        dialog.close();
        resolve(null);
      });
    });
  });
}

async function handleSpamReport(event: Office.AddinCommands.Event) {
  const mail = Office.context.mailbox.item,
    reportResult = await reportSpam(mail);
  switch (reportResult.reportStatus) {
    case ReportResultStatus.SIMULATION:
      await showSimulationAcknowledgement();
      break;
    case ReportResultStatus.ERROR:
      await showErrorDialog(reportResult.diagnosis);
      break;
  }
  // Move message after the user has closed any potential dialogs.
  // If we had moved it earlier, any spawned dialog would close immediately.
  if (reportResult.moveMessageStatus === MoveMessageStatus.PENDING)
    reportResult.moveMessageStatus = await moveMessageTo(mail, reportResult.moveMessageTarget);
  event.completed();
}

Office.actions.associate("reportSpam", handleSpamReport);
