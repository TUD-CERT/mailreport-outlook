/* global console, document, DOMParser, fetch, Office */
import { getSettings, ReportAction } from "../settings";

enum BodyType {
  PLAIN,
  HTML,
}

class Message {
  from: string;
  to: string;
  date: Date;
  subject: string;
  preview: string;
  previewType: BodyType;
  raw: string;
}

/**
 * Encodes various characters to their safe HTML counterparts. Used to prevent HTML interpretation of
 * E-Mail headers such as "Name <name@example.com>".
 */
function encodeHTML(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

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
  // Retrieve raw email via EWS
  const request =
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header><t:RequestServerVersion Version="Exchange2013" /></soap:Header>' +
    "  <soap:Body>" +
    "    <m:GetItem>" +
    "      <m:ItemShape>" +
    "        <t:BaseShape>IdOnly</t:BaseShape>' +" +
    "        <t:IncludeMimeContent>true</t:IncludeMimeContent>" +
    "      </m:ItemShape >" +
    "      <m:ItemIds>" +
    `        <t:ItemId Id="${email.itemId}" />` +
    "      </m:ItemIds>" +
    "    </m:GetItem>" +
    "  </soap:Body>" +
    "</soap:Envelope>";
  const rawEMail = await new Promise<string>((resolve) => {
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(result.value, "text/xml");
      const values = doc.getElementsByTagName("t:MimeContent");
      resolve(values[0].textContent);
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
  result.raw = rawEMail;
  return result;
}

async function sendSMTPReport(
  destination: string,
  subject: string,
  lucyClientID: number = null,
  message: Message,
  comment: string = null
) {
  // Send email via EWS
  let bodyXML = "",
    commentBody = "",
    lucyClientBody = "",
    lucyCIBody = "";
  switch (message.previewType) {
    case BodyType.PLAIN:
      commentBody = comment !== null ? `X-More-Analysis: True\n${comment}\n` : "";
      lucyClientBody = lucyClientID !== null ? `X-Lucy-Client: ${lucyClientID}\n` : "";
      lucyCIBody = lucyClientID !== null ? `X-CI-Report: True\n` : "";
      bodyXML = `<t:Body BodyType="Text">${lucyClientBody}${commentBody}${lucyCIBody}\n\n-----Original Message-----\nFrom: ${message.from}\nSent: ${message.date.toString()}\nTo: ${message.to}\nSubject: ${message.subject}\n\n${message.preview}\r\n</t:Body>`;
      break;
    case BodyType.HTML:
      commentBody = comment !== null ? `X-More-Analysis: True<br />${comment}<br />` : "";
      lucyClientBody = lucyClientID !== null ? `X-Lucy-Client: ${lucyClientID}<br />` : "";
      lucyCIBody = lucyClientID !== null ? `X-CI-Report: True<br />` : "";
      bodyXML = `<t:Body BodyType="HTML"><![CDATA[${lucyClientBody}${commentBody}${lucyCIBody}<br /><br />From: ${encodeHTML(message.from)}<br />Sent: ${encodeHTML(message.date.toString())}<br />To: ${encodeHTML(message.to)}<br />Subject: ${encodeHTML(message.subject)}<br /><br />${message.preview}]]></t:Body>`;
      break;
  }
  const request =
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header><t:RequestServerVersion Version="Exchange2013" /></soap:Header>' +
    "  <soap:Body>" +
    '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
    "      <m:SavedItemFolderId>" +
    '        <t:DistinguishedFolderId Id="sentitems" />' +
    "      </m:SavedItemFolderId>" +
    "      <m:Items>" +
    "        <t:Message>" +
    `          <t:Subject>${subject}</t:Subject>` +
    `          ${bodyXML}` +
    "          <t:Attachments>" +
    "            <t:FileAttachment>" +
    `              <t:Name>email.eml</t:Name>` +
    "              <t:ContentType>application/octet-stream</t:ContentType>" +
    `              <t:Content>${message.raw}</t:Content>` +
    "            </t:FileAttachment>" +
    "          </t:Attachments>" +
    "          <t:ToRecipients>" +
    "            <t:Mailbox>" +
    `              <t:EmailAddress>${destination}</t:EmailAddress>` +
    "            </t:Mailbox>" +
    "          </t:ToRecipients>" +
    "        </t:Message>" +
    "      </m:Items>" +
    "    </m:CreateItem>" +
    "  </soap:Body>" +
    "</soap:Envelope>";
  return await new Promise<void>((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(result.value, "text/xml");
      const values = doc.getElementsByTagName("m:ResponseCode");
      if (values[0].textContent === "NoError") resolve();
      else reject();
    });
  });
}

async function moveMessageTo(email: Office.MessageRead, folder: ReportAction) {
  // Move email via EWS
  let folderId = "";
  switch (folder) {
    case ReportAction.JUNK:
      folderId = "junkemail";
      break;
    case ReportAction.TRASH:
      folderId = "deleteditems";
      break;
    case ReportAction.KEEP:
      return;
  }
  console.log(`Moving message ${email.itemId} to ${folder} folder`);
  const request =
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header><t:RequestServerVersion Version="Exchange2013" /></soap:Header>' +
    "  <soap:Body>" +
    "    <m:MoveItem>" +
    "      <m:ToFolderId>" +
    `        <t:DistinguishedFolderId Id="${folderId}"/>` +
    "      </m:ToFolderId >" +
    "      <m:ItemIds>" +
    `        <t:ItemId Id="${email.itemId}" />` +
    "      </m:ItemIds>" +
    "    </m:MoveItem>" +
    "  </soap:Body>" +
    "</soap:Envelope>";
  return await new Promise<void>((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(result.value, "text/xml");
      const values = doc.getElementsByTagName("m:ResponseCode");
      if (values[0].textContent === "NoError") resolve();
      else reject();
    });
  });
}

export async function reportMail() {
  const mail = Office.context.mailbox.item;
  const message = await parseMessage(mail);
  const comment = (<HTMLTextAreaElement>document.getElementById("reportComment")).value;
  await sendSMTPReport("cert@exchg.cert", "Phishing Report", 2, message, comment.length > 0 ? comment : null);
  await moveMessageTo(mail, getSettings().report_action);
  Office.context.ui.closeContainer();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sendFraudReport").onclick = reportMail;
  }
});
