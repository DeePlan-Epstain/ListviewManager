import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { EMailProperties } from '../../../models/global.model';
import { IService } from '../models/IService';
import { SendEMailDialogContent } from '../SendEMailDialogContent/SendEMailDialogContent';

export default class SendEMailDialog extends BaseDialog {
    private eMailProperties: EMailProperties;
    private sendDocumentService: IService;

    // When the object is created the constructor recive the object SendDocumentService from "ExternalSharingCommandSet" as input (passed into "service")
    // eMailProperties is not mandatory ("ExternalSharingCommandSet" is not obligated to pass it)
    constructor(service: IService, eMailProperties?: EMailProperties) {
        super();
        // Populate sendDocumentService with the object recived by "ExternalSharingCommandSet"
        this.sendDocumentService = service;
        // Set EMail properties
        if (eMailProperties) {
            this.eMailProperties = eMailProperties;
        }
        else {
            this.eMailProperties = new EMailProperties({
                To: "",
                Cc: "",
                Subject: `שיתוף מסמך - ${this.sendDocumentService.fileNames}`,
                Body: "",
            });
        }
    }

    // Render the component SendEMailDialogContent using ReactDOM, the component recive the eMailProperties and the service SendDocumentService as props (arguments passed into React components)
    public render(): void {
        ReactDOM.render(<SendEMailDialogContent
            close={this._close.bind(this)}
            eMailProperties={this.eMailProperties}
            submit={this._submit.bind(this)}
            sendDocumentService={this.sendDocumentService}
        />, this.domElement);
    }

    // Reset eMailProperties + remove the mounted React component from the DOM and clean up its event handlers and state
    private clear() {
        if (this.eMailProperties) {
            this.eMailProperties === undefined;
        }

        ReactDOM.unmountComponentAtNode(this.domElement);
    }

    // Close and clear the react modal
    private _close(): void {
        this.clear();
        this.close();
    }

    // Close and clear the react modal when the email sent successfully
    private _submit(eMailProperties: EMailProperties): void {
        this.clear();
        this.close();
    }
}