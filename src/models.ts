export enum BodyType {
  PLAIN,
  HTML,
}

export enum ReportAction {
  JUNK = "junk",
  TRASH = "trash",
  KEEP = "keep",
}

export enum ReportResult {
  SUCCESS,
  SIMULATION,
  ERROR,
}

export enum Transport {
  HTTP = "http",
  SMTP = "smtp",
  HTTPSMTP = "http+smtp",
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
  from: string;
  to: string;
  date: Date;
  subject: string;
  headers: object; // {key: [val1, val2, ...]} with lowercase keys
  preview: string;
  previewType: BodyType;
  raw: string;
}
