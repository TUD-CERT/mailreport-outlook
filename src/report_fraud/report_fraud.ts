/* global document, HTMLTextAreaElement, Office, setTimeout */
import { moveMessageTo, sendSMTPReport } from "../ews";
import { localizeDocument } from "../i18n";
import { parseMessage } from "../reporting";
import { getSettings } from "../settings";
import { fixOWAPadding } from "../utils";

function showView(selector: string) {
  const $unselected = document.querySelectorAll(`div.view:not(${selector})`);
  document.querySelector(selector).classList.remove("hide");
  for (const e of $unselected) e.classList.add("hide");
}

async function reportFraud() {
  showView("#mailreport-fraud-pending");
  const mail = Office.context.mailbox.item;
  const message = await parseMessage(mail);
  const comment = (<HTMLTextAreaElement>document.getElementById("reportComment")).value;
  const successReport = await sendSMTPReport(
    "cert@exchg.cert",
    "Phishing Report",
    2,
    message,
    comment.length > 0 ? comment : null
  );
  const successMove = await moveMessageTo(mail, getSettings().report_action);
  showView(successReport && successMove ? "#mailreport-fraud-success" : "#mailreport-fraud-error");
  await new Promise((r) => setTimeout(r, successReport && successMove ? 2000 : 5000));
  Office.context.ui.closeContainer();
}

Office.onReady((info) => {
  localizeDocument();
  fixOWAPadding();
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sendFraudReport").onclick = reportFraud;
  }
});
