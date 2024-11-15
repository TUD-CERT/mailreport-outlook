/* global console, fetch, Office */
import { BodyType, Message, Transport } from "./models";
import { getRawEmail, moveMessageTo, sendSMTPReport } from "./ews";
import { ReportAction } from "./models";
import { getSettings } from "./settings";

async function parseMessage(email: Office.MessageRead): Promise<Message> {
  // The MailBox API doesn't permit us to programatically determine the original MIME content structure.
  // Therefore, we always coalesce content to HTML.

  // Retrieve HTML version of body
  const htmlContent = await new Promise<string>((resolve, reject) => {
    email.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject();
      }
    });
  });
  const result: Message = new Message();
  // Parse message sender: if email.from and email.sender differ, the message was sent by a delegate
  const messageFrom =
    email.from.displayName.length > 0
      ? `${email.from.displayName} <${email.from.emailAddress}>`
      : email.from.emailAddress;
  const messageSender =
    email.sender.displayName.length > 0
      ? `${email.sender.displayName} <${email.sender.emailAddress}>`
      : email.sender.emailAddress;
  result.from =
    email.from.emailAddress !== email.sender.emailAddress ? `${messageFrom} via ${messageSender}` : messageFrom;
  result.to = email.to
    .map((v) => (v.displayName.length > 0 ? `${v.displayName} <${v.emailAddress}>` : v.emailAddress))
    .join(", ");
  result.date = email.dateTimeCreated;
  result.subject = email.subject;
  result.preview = htmlContent;
  result.previewType = BodyType.HTML;
  result.raw = await getRawEmail(email.itemId);
  return result;
}

/**
 * Tries to send HTTP messages to the given Lucy URLs until one succeeds.
 * Returns a boolean to indicate whether the report was successful.
 */
async function sendHTTPReport(
  urls: string[],
  reporterAddress: string,
  message: Message,
  comment: string | null = null
) {
  const lucyReport = {
    email: reporterAddress,
    mail_content: message.raw,
    more_analysis: comment !== null,
    disable_incident_autoresponder: false,
    enable_comment_to_deeper_analysis_request: comment === null ? "" : comment,
  };
  let success = false;
  for (let url of urls) {
    console.log("Sending report as ", reporterAddress, " via HTTP(S) to ", url);
    // Send report
    try {
      await fetch(url, {
        method: "POST",
        mode: "no-cors", // Lucy server sets multiple CORS header, which Chrome/Edge doesn't like
        headers: { "Content-Type": "text/plain; Charset=UTF-8" }, // Content-Type taken from the Lucy Outlook AddIn
        body: JSON.stringify(lucyReport),
      });
      success = true;
      break;
    } catch (err) {
      console.log("Could not send report via HTTP(S)", err);
    }
  }
  return success;
}

export async function reportFraud(mail: Office.MessageRead, comment: string) {
  const message = await parseMessage(mail),
    settings = getSettings(),
    transport = settings.phishing_transport,
    parsedComment = comment.length > 0 ? comment : null;
  let success = true;

  if (transport === Transport.HTTP || transport === Transport.HTTPSMTP) {
    let lucyReportURL = `https://${settings.lucy_server}/phishing-report`;
    if (settings.lucy_client_id !== null) lucyReportURL += `/${settings.lucy_client_id}`;
    success =
      success &&
      (await sendHTTPReport([lucyReportURL], Office.context.mailbox.userProfile.emailAddress, message, parsedComment));
  }
  if (transport === Transport.SMTP || transport === Transport.HTTPSMTP) {
    let subject = "Phishing Report";
    if (settings.smtp_use_expressive_subject) subject += `: ${message.subject}`;
    success =
      success && (await sendSMTPReport(settings.smtp_to, subject, settings.lucy_client_id, message, parsedComment));
  }

  if (success) await moveMessageTo(mail, settings.report_action);
  return success;
}

export async function reportSpam(mail: Office.MessageRead) {
  const message = await parseMessage(mail),
    settings = getSettings(),
    transport = settings.phishing_transport;
  if (transport === Transport.HTTP) return; // HTTP endpoint does not support spam reports
  let subject = "Spam Report";
  if (settings.smtp_use_expressive_subject) subject += `: ${message.subject}`;
  await sendSMTPReport(settings.smtp_to, subject, settings.lucy_client_id, message, null);
  await moveMessageTo(mail, ReportAction.JUNK);
}
