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

// Returns Email Attachment object
export class EMailAttachment {
  public FileName: string;
  public ContentBytes: string;

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
