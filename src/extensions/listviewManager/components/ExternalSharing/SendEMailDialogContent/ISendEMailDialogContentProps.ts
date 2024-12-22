import { EMailProperties } from "../../../models/global.model";
import { IService } from "../models/IService";

// SendEMailDialogContent props (arguments passed into React components)
export interface ISendEMailDialogContentProps {
    close: () => void;
    submit: (eMailProperties: EMailProperties) => void;
    eMailProperties: EMailProperties;
    sendDocumentService: IService;
}
