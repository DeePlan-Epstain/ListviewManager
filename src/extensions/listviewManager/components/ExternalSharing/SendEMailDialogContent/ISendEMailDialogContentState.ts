export interface ISendEMailDialogContentState {
   isLoading: boolean;
   DialogTitle: string;
   MailOptionTo: string;
   MailOptionCc: string;
   MailOptionSubject: string;
   MailOptionBody: string;
   SendToError: string;
   SendCcError: string;
   SubjectError: string;
   SendEmailFailedError: boolean;
   succeed?: boolean;
   ESArray: string[],
   error?: Error;
}
