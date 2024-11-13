export enum BodyType {
  PLAIN,
  HTML,
}

export enum ReportAction {
  JUNK = "junk",
  TRASH = "trash",
  KEEP = "keep",
}

export enum Transport {
  HTTP = "http",
  SMTP = "smtp",
  HTTPSMTP = "http+smtp",
}

export class Settings {
  permit_advanced_config: boolean;
  report_action: ReportAction;
  smtp_to: string;
  smtp_use_expressive_subject: boolean;
}

export class Message {
  from: string;
  to: string;
  date: Date;
  subject: string;
  preview: string;
  previewType: BodyType;
  raw: string;
}
