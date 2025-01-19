import * as React from 'react';
import { EMailAttachment, DraftProperties } from '../../models/global.model';
import { IDraftProps } from './IDraftProps';
import { jss } from "../../models/jss";
import { cacheRtl } from "../../models/cacheRtl";
import { StylesProvider } from '@material-ui/core/styles';
import { CacheProvider } from '@emotion/react';
import { Modal, CircularProgress } from '@mui/material'

import moment, { Moment } from 'moment';


export interface IDraftState {
    isLoading: boolean;
    SendEmailFailedError: boolean;
    succeed?: boolean;
    error?: Error;
}

export default class Draft extends React.Component<IDraftProps, IDraftState> {

    private _draftProperties: DraftProperties;
    private copiedFileUri: string[];

    constructor(props: IDraftProps) {
        super(props);

        this.state = {
            isLoading: false,
            succeed: false,
            SendEmailFailedError: false,
        };
        this._draftProperties = this.props.draftProperties;
        this._submit = this._submit.bind(this);
    }

    componentDidMount() {
        this._submit()
    }

    // Returns EMailAttachment object which contains the file name and its Content Encodes into base64 string
    private getEMailAttachment(): Promise<EMailAttachment[]> {
        return new Promise((resolve, reject) => {
            const { createDraft } = this.props;

            // Initialize an empty array to store the copied file URIs
            let copiedFileUris: string[] = [];

            Promise.all(createDraft.fileUris.map((fileUri: any, index: any) =>
                createDraft.CopyFileAndCleanMetadata(
                    [fileUri],
                    [createDraft.fileNames[index]],
                    [createDraft.DocumentIdUrls[index]],
                    createDraft.ServerRelativeUrl
                ).then(async (copiedFileUri: string[]) => {
                    // Add the copied file URIs to the array
                    copiedFileUris = copiedFileUris.concat(copiedFileUri);

                    return createDraft.getFileContentAsBase64(copiedFileUri).then((fileContent: string[]) => {
                        const contentBytes = Array.isArray(fileContent) ? fileContent.join('') : fileContent;

                        return new EMailAttachment({
                            FileName: createDraft.fileNames[index],
                            ContentBytes: contentBytes,
                        });
                    });
                })
            ))
                .then((attachments: EMailAttachment[]) => {
                    // Assign the accumulated file URIs to the instance property
                    this.copiedFileUri = copiedFileUris;
                    resolve(attachments);
                })
                .catch((err: any) => {
                    reject(err);
                });
        });
    }

    private createDraft(DraftProperties: DraftProperties): Promise<string> {
        return new Promise((resolve, reject) => {
            this.props.createDraft.createDraft(DraftProperties)
                .then((draftId: string) => {
                    resolve(draftId);
                })
                .catch((err: any) => {
                    reject(err);
                });
        })
    }

    // Submit the form
    public _submit() {

        this.getEMailAttachment().then((attachments: EMailAttachment[]) => {
            this._draftProperties.Attachment = attachments;
            this.createDraft(this._draftProperties)
                .then(async (draftId: string) => {
                    console.log(".then - draftId:", draftId)
                    // await this.openInOutlook(draftId)
                    this.setState({ succeed: true, isLoading: false });
                    this.props.createDraft.DeleteCopiedFile(this.copiedFileUri);
                    setTimeout(() => {
                        this.props.close(); // Close the modal after a delay for visual feedback
                    }, 1000);
                })
                .catch((err: Error) => {
                    console.error("Send Document Error", err);
                    this.setState({
                        SendEmailFailedError: true,
                        isLoading: false,
                    });
                });
        });

    }

    private async openInOutlook(draftId: string): Promise<void> {
        console.log('Opening draft email');

        // Construct the Outlook Desktop URL
        const outlookDesktopUrl = `https://outlook.office.com/mail/deeplink/compose/${draftId}`;

        // Check if the desktop protocol is supported
        try {
            // Attempt to open the desktop URL
            window.open(outlookDesktopUrl, '_blank');

            console.log('Attempting to open in Outlook Desktop:', outlookDesktopUrl);
        } catch (error) {
            console.error('Failed to open in Outlook Desktop:', error);

        }
    }

    public render() {
        return (
            <CacheProvider value={cacheRtl}>
                <StylesProvider jss={jss}>
                    <Modal
                        open={true}
                        onClose={(event, reason) => {
                            if (this.state.isLoading || this.state.succeed) {
                                return;
                            }
                            else this.props.close();
                        }}
                        aria-labelledby="modal-title"
                        aria-describedby="modal-description"
                        className="ModalCustom"
                        sx={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                        }}
                    >
                        <div
                            className="ModalContentContainer"
                            dir="rtl"
                            style={{
                                width: '250px',
                                background: 'white',
                                padding: '20px',
                                borderRadius: '5px',
                                boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                                display: 'flex',
                                justifyContent: 'center',
                                alignItems: 'center',
                                flexDirection: 'column'
                            }}
                        >
                            <h3>אנא המתן...</h3>
                            <CircularProgress />
                        </div>
                    </Modal>
                </StylesProvider>
            </CacheProvider >
        );
    }
}
