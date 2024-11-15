/* global console, DOMParser, Office */
import { BodyType, Message, ReportAction } from "./models";
import { encodeHTML } from "./utils";

/**
 * Functionality implemented via legacy EWS due to missing equivalent methods
 * in the Outlook JavaScript API for on-premises Exchange/Outlook environments.
 * The Outlook REST APIs aren't usable due to CORS issues in that
 * setting, apparently by design: https://github.com/OfficeDev/office-js-docs-pr/issues/2166
 * The Microsoft Graph API is not available in on-premises Exchange, either.
 */

export async function getRawEmail(ewsId: string) {
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
    `        <t:ItemId Id="${ewsId}" />` +
    "      </m:ItemIds>" +
    "    </m:GetItem>" +
    "  </soap:Body>" +
    "</soap:Envelope>";
  return await new Promise<string>((resolve) => {
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(result.value, "text/xml");
      const values = doc.getElementsByTagName("t:MimeContent");
      resolve(values[0].textContent);
    });
  });
}

export async function sendSMTPReport(
  destination: string,
  subject: string,
  lucyClientID: number | null = null,
  message: Message,
  comment: string | null = null
) {
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
  return await new Promise<boolean>((resolve) => {
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(result.value, "text/xml");
      const values = doc.getElementsByTagName("m:ResponseCode");
      if (values[0].textContent === "NoError") resolve(true);
      else resolve(false);
    });
  });
}

export async function moveMessageTo(email: Office.MessageRead, folder: ReportAction) {
  let folderId = "";
  switch (folder) {
    case ReportAction.JUNK:
      folderId = "junkemail";
      break;
    case ReportAction.TRASH:
      folderId = "deleteditems";
      break;
    case ReportAction.KEEP:
      return true;
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
  return await new Promise<boolean>((resolve) => {
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(result.value, "text/xml");
      const values = doc.getElementsByTagName("m:ResponseCode");
      if (values[0].textContent === "NoError") resolve(true);
      else resolve(false);
    });
  });
}
