/* global Office */
import { ReportResult } from "../models";
import { reportSpam } from "../reporting";
import { showSimulationAcknowledgement } from "../simulation";

// Must be run each time a new page is loaded.
Office.onReady();

async function handleSpamReport(event: Office.AddinCommands.Event) {
  const reportResult = await reportSpam(Office.context.mailbox.item);
  if (reportResult === ReportResult.SIMULATION) await showSimulationAcknowledgement();
  event.completed();
}

Office.actions.associate("reportSpam", handleSpamReport);
