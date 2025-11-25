export enum BodyType {
  PLAIN,
  HTML,
}

export enum ReportAction {
  JUNK = "junk",
  TRASH = "trash",
  KEEP = "keep",
}

export enum ReportResultStatus {
  SUCCESS,
  SIMULATION,
  ERROR,
}

export enum Transport {
  HTTP = "http",
  SMTP = "smtp",
  HTTPSMTP = "http+smtp",
}

export enum OfficeThemeId {
  Colorful = "#000000",
  DarkGray = "#000001",
  Black = "#000002",
  White = "#000003",
}

export class Settings {
  lucy_client_id: number | null;
  lucy_server: string;
  permit_advanced_config: boolean;
  report_action: ReportAction;
  phishing_transport: Transport;
  plugin_id: string;
  send_telemetry: boolean;
  simulation_transport: Transport;
  smtp_to: string;
  smtp_use_expressive_subject: boolean;
}

export class Message {
  from: string;           // Sender of this message according to its own headers
  to: string;             // Receivers of this message according to its own headers
  reporter: string;       // E-Mail address of the account that reported this message
  date: Date;             // Date and time of the reported message
  subject: string;        // Parsed subject of the reported message
  headers: object;        // Header section of the repoted message as {key: [val1, val2, ...]} with lowercase keys
  preview: string;        // Preview of the reported message, typically just HTML or PLAIN body content
  previewType: BodyType;  // Specifies whether the preview is in HTML or PLAIN format
  raw: string;            // Raw bytes of the reported message
}

export class ReportResult {
  status: ReportResultStatus;
  diagnosis: string;

  constructor(status: ReportResultStatus) {
    this.status = status;
  }
}
