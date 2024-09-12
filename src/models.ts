export enum BodyType {
  PLAIN,
  HTML,
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
