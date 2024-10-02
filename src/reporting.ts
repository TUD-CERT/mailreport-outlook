/* global Office */
import { BodyType, Message } from "./models";
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

export async function reportFraud(mail: Office.MessageRead, comment: string) {
  const message = await parseMessage(mail);
  const successReport = await sendSMTPReport(
    "cert@exchg.cert",
    "Phishing Report",
    2,
    message,
    comment.length > 0 ? comment : null
  );
  const successMove = await moveMessageTo(mail, getSettings().report_action);
  return successReport && successMove;
}

export async function reportSpam(mail: Office.MessageRead) {
  const message = await parseMessage(mail);
  await sendSMTPReport("cert@exchg.cert", "Spam Report", 2, message, null);
  await moveMessageTo(mail, ReportAction.JUNK);
}