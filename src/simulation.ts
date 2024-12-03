/* global Office, window */
import URI from "urijs";

export async function showSimulationAcknowledgement() {
  const url = new URI("simulation_ack.html").absoluteTo(window.location).toString();
  const dialogOptions = { width: 60, height: 20, displayInIframe: true };
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
