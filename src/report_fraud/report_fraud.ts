/* global document, HTMLTextAreaElement, Office */
import { moveMessageTo, sendSMTPReport } from "../ews";
import { localizeDocument } from "../i18n";
import { parseMessage } from "../reporting";
import { getSettings } from "../settings";

async function reportFraud() {
  const mail = Office.context.mailbox.item;
  const message = await parseMessage(mail);
  const comment = (<HTMLTextAreaElement>document.getElementById("reportComment")).value;
  await sendSMTPReport("cert@exchg.cert", "Phishing Report", 2, message, comment.length > 0 ? comment : null);
  await moveMessageTo(mail, getSettings().report_action);
  Office.context.ui.closeContainer();
}

Office.onReady((info) => {
  localizeDocument();
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sendFraudReport").onclick = reportFraud;
  }
});
