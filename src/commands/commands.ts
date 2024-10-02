/* global Office */
import { reportSpam } from "../reporting";

// Must be run each time a new page is loaded.
Office.onReady();

async function handleSpamReport(event: Office.AddinCommands.Event) {
  await reportSpam(Office.context.mailbox.item);
  event.completed();
}

Office.actions.associate("reportSpam", handleSpamReport);
