/* global Office, window */
import { ReportResultStatus } from "../models";
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
  const reportResult = await reportSpam(Office.context.mailbox.item);

  switch (reportResult.status) {
    case ReportResultStatus.SIMULATION:
      await showSimulationAcknowledgement();
      break;
    case ReportResultStatus.ERROR:
      await showErrorDialog(reportResult.diagnosis);
      break;
  }
  event.completed();
}

Office.actions.associate("reportSpam", handleSpamReport);
