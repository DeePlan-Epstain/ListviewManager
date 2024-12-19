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

export type File = {
  FileRef: string;
  FileLeafRef: string;
}