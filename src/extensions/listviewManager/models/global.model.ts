export type SelectedFile = {
  ID: number;
  FileRef: string;
  approveStatus: string;
  latestApprovers: string;
  latestReviewers: string;
  jiraId: string;
  comment: string;
  FSObjType: string;
  ItemChildCount: string;
  FileLeafRef: string;
  libraryId: string;
  initiatorId: any;
  CheckoutUser: any[];
  ServerRedirectedEmbedUrl: string;
  serverurl: { progid: string };
  UniqueId: string;
};

export type DocumentProps = {
  Id: number;
  Title: string;
  libraryId: string;
  comment: string;
  latestApprovers: any;
  approveStatus: string;

  ApprovalStatus: string;
  Initiator: number;
  ServerRelativeUrl: string;
};

export type File = {
  FileRef: string;
  FileLeafRef: string;
}
// Returns Email Attachment object
export class EMailAttachment {
  public FileName: string;
  public ContentBytes: string | ArrayBuffer;

  constructor(options: EMailAttachment) {
    this.FileName = options.FileName;
    this.ContentBytes = options.ContentBytes;
  }
}

// Returns Email properties object
export class EMailProperties {
  public To: string;
  public Cc: string;
  public Subject: string;
  public Body: string;
  public Attachment?: EMailAttachment[];

  constructor(options: EMailProperties) {
    this.To = options.To;
    this.Cc = options.Cc;
    this.Subject = options.Subject;
    this.Body = options.Body;
    this.Attachment = options.Attachment;
  }
}

export class EventProperties {
  public To: string;
  public optionals: string;
  public Subject: string;
  public Date: string;
  public startTime: string;
  public endTime: string;
  public onlineMeeting: boolean;
  public Body: string;
  public Attachment?: EMailAttachment[];

  constructor(options: EventProperties) {
    this.To = options.To;
    this.optionals = options.optionals
    this.Subject = options.Subject;
    this.Date = options.Date;
    this.startTime = options.startTime;
    this.endTime = options.endTime;
    this.onlineMeeting = options.onlineMeeting;
    this.Body = options.Body;
    this.Attachment = options.Attachment;
  }
}
export class DraftProperties {
  public Subject: string;
  public Body: string;
  public Attachment?: EMailAttachment[];

  constructor(options: DraftProperties) {
    this.Subject = options.Subject;
    this.Body = options.Body;
    this.Attachment = options.Attachment;
  }
}