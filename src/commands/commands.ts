/* global Office */
import { moveMessageTo, sendSMTPReport } from "../ews";
import { ReportAction } from "../models";
import { parseMessage } from "../reporting";

// Must be run each time a new page is loaded.
Office.onReady();

async function reportSpam(event: Office.AddinCommands.Event) {
  const mail = Office.context.mailbox.item;
  const message = await parseMessage(mail);
  await sendSMTPReport("cert@exchg.cert", "Spam Report", 2, message, null);
  await moveMessageTo(mail, ReportAction.JUNK);
  event.completed();
}

Office.actions.associate("reportSpam", reportSpam);
