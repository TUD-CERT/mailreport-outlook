/* global console, Office, window */
import { BodyType, Message, Transport } from "./models";
import { fetchMessage, moveMessageTo, sendSMTPReport } from "./ews";
import { ReportAction, ReportResult } from "./models";
import { getSettings } from "./settings";
import "whatwg-fetch";

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
  const message = await fetchMessage(email.itemId);
  result.raw = message.raw;
  result.headers = message.headers;
  return result;
}

/**
 * Checks a message for Lucy headers that indicate the e-mail is part of a phishing simulation.
 */
function belongsToSimulation(message: Message) {
  for (let key in message.headers) {
    if (key.startsWith("x-lucy")) return true;
  }
  return false;
}

/**
 * Parses Lucy mail headers and returns an array of reporting URLs.
 */
function getReportingURLs(message: Message) {
  let urls = [];
  for (let key in message.headers) {
    if (key.includes("x-lucy") && key.includes("victimurl")) urls.push(message.headers[key][0]);
  }
  return urls;
}

/**
 * Parses Lucy mail headers and returns the scenario ID or null, if none was found.
 */
function getScenarioID(message: Message): string | null {
  return message.headers["x-lucy-scenario"][0] ?? null;
}

/**
 * Tries to send HTTP messages to the given Lucy URLs until one succeeds.
 * Returns a boolean to indicate whether the report was successful.
 */
async function sendHTTPReport(
  urls: string[],
  reporterAddress: string,
  message: Message,
  additionalHeaders: { [key: string]: string },
  lucyScenarioID: string | null = null,
  comment: string | null = null
) {
  const lucyReport: { [key: string]: any } = {
    email: reporterAddress,
    mail_content: message.raw,
    more_analysis: comment !== null,
    disable_incident_autoresponder: false,
    enable_comment_to_deeper_analysis_request: comment === null ? "" : comment,
  };
  if (lucyScenarioID !== null) lucyReport.scenario_id = lucyScenarioID;
  let success = false;
  for (let url of urls) {
    if ("scenario_id" in lucyReport)
      console.log(
        "Reporting simulation for scenario ",
        lucyScenarioID,
        " as ",
        reporterAddress,
        " via HTTP(S) to ",
        url
      );
    else console.log("Sending report as ", reporterAddress, " via HTTP(S) to ", url);
    // Send report
    try {
      await window.fetch(url, {
        method: "POST",
        mode: "no-cors", // Lucy server sets multiple CORS header, which Chrome/Edge doesn't like
        headers: { "Content-Type": "text/plain; Charset=UTF-8", ...additionalHeaders }, // Content-Type taken from the Lucy Outlook AddIn
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

/**
 * Returns an object with additional telemetry headers to send with each request
 * (derived from current plugin settings).
 */
function getAdditionalHeaders() {
  const headers: { [key: string]: any } = {},
    settings = getSettings();
  if (settings.send_telemetry) {
    headers["Reporting-Agent"] =
      `${Office.context.mailbox.diagnostics.hostName}/${Office.context.mailbox.diagnostics.hostVersion}`;
    headers["Reporting-Plugin"] = settings.plugin_id;
  }
  return headers;
}

export async function reportFraud(mail: Office.MessageRead, comment: string): Promise<ReportResult> {
  const message = await parseMessage(mail),
    isSimulation = belongsToSimulation(message),
    settings = getSettings(),
    transport = isSimulation ? settings.simulation_transport : settings.phishing_transport,
    parsedComment = comment.length > 0 ? comment : null,
    additionalHeaders = getAdditionalHeaders();
  let success = true;

  if (transport === Transport.HTTP || transport === Transport.HTTPSMTP) {
    let lucyReportURL = `https://${settings.lucy_server}/phishing-report`;
    if (settings.lucy_client_id !== null) lucyReportURL += `/${settings.lucy_client_id}`;
    const lucyScenarioID = isSimulation ? getScenarioID(message) : null;
    let urls = isSimulation ? getReportingURLs(message) : [lucyReportURL];
    // If invalid Lucy headers are set, fall back to the configured Lucy instance
    if (urls.length === 0) urls = [lucyReportURL];
    success =
      success &&
      (await sendHTTPReport(
        urls,
        Office.context.mailbox.userProfile.emailAddress,
        message,
        additionalHeaders,
        lucyScenarioID,
        parsedComment
      ));
  }
  if (transport === Transport.SMTP || transport === Transport.HTTPSMTP) {
    let subject = "Phishing Report";
    if (settings.smtp_use_expressive_subject) subject += `: ${message.subject}`;
    success =
      success &&
      (await sendSMTPReport(
        settings.smtp_to,
        subject,
        settings.lucy_client_id,
        message,
        additionalHeaders,
        parsedComment
      ));
  }
  if (success) await moveMessageTo(mail, settings.report_action);
  if (success) {
    if (isSimulation) return ReportResult.SIMULATION;
    return ReportResult.SUCCESS;
  }
  return ReportResult.ERROR;
}

export async function reportSpam(mail: Office.MessageRead): Promise<ReportResult> {
  const message = await parseMessage(mail),
    settings = getSettings(),
    transport = settings.phishing_transport,
    additionalHeaders = getAdditionalHeaders();
  if (belongsToSimulation(message)) {
    const result = await reportFraud(mail, "");
    // Users expect reported spam mails to be moved away even if ReportAction is KEEP
    if (settings.report_action === ReportAction.KEEP) await moveMessageTo(mail, ReportAction.JUNK);
    return result;
  }
  if (transport === Transport.HTTP) return; // HTTP endpoint does not support spam reports
  let subject = "Spam Report";
  if (settings.smtp_use_expressive_subject) subject += `: ${message.subject}`;
  let success = await sendSMTPReport(
    settings.smtp_to,
    subject,
    settings.lucy_client_id,
    message,
    additionalHeaders,
    null
  );
  await moveMessageTo(mail, ReportAction.JUNK);
  return success ? ReportResult.SUCCESS : ReportResult.ERROR;
}
