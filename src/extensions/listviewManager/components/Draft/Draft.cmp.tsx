import * as React from 'react';
import { EMailAttachment, DraftProperties } from '../../models/global.model';
import { IDraftProps } from './IDraftProps';
import { jss } from "../../models/jss";
import { cacheRtl } from "../../models/cacheRtl";
import { StylesProvider } from '@material-ui/core/styles';
import { CacheProvider } from '@emotion/react';
import { Modal, CircularProgress } from '@mui/material'
import { Dialog as DialogMicrosoft } from '@microsoft/sp-dialog';
import moment, { Moment } from 'moment';


export interface IDraftState {
    isLoading: boolean;
    SendEmailFailedError: boolean;
    succeed?: boolean;
    error?: Error;
    link: string
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
            link: '',
        };
        this._draftProperties = this.props.draftProperties;
        this._submit = this._submit.bind(this);
    }

    componentDidMount() {
        this._submit()
    }

    componentWillUnmount(): void {
        if (typeof this.state.link !== 'string') {
            DialogMicrosoft.alert('Something went wrong')
        } else {
            window.open(this.state.link, '_blank');
        }
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
                .then((link: string) => {
                    resolve(link);
                })
                .catch((err: any) => {
                    reject(err);
                });
        })
    }

    // Submit the form
    public async _submit() {
        const fileUris: string[] = this.props.createDraft.fileUris;
        const fileNames: string[] = this.props.createDraft.fileNames;
        try {
            // 1. Fetch all ArrayBuffers in parallel:
            const buffers: ArrayBuffer[] = await Promise.all(
                fileUris.map((uri: string) =>
                    this.props.sp.web
                        .getFileByServerRelativePath(uri)
                        .getBuffer()
                )
            );

            // 2a. Build attachments with a `map` (concise):
            const attachments: EMailAttachment[] = buffers.map((buf, idx) => ({
                FileName: fileNames[idx],
                ContentBytes: buf
            }));
            this._draftProperties.Attachment = attachments;
            await this.props.createDraft.createDraft(this._draftProperties).then((link) => {
                if (typeof link !== 'string') {
                    this.props.close();
                    return;
                }
                this.setState({ succeed: true, isLoading: false, link: link as string });
            }).then(() => {
                this.props.close(); // Close the modal after a delay for visual feedback
            });

        } catch (error) {
            console.error('Error submit draft', error)
        }
        return;

        this.getEMailAttachment().then((attachments: EMailAttachment[]) => {
            this._draftProperties.Attachment = attachments;
            this.createDraft(this._draftProperties)
                .then((link: string) => {

                    this.props.createDraft.DeleteCopiedFile(this.copiedFileUri).then(() => {
                        this.setState({ succeed: true, isLoading: false, link: link });
                    }).then(() => {
                        this.props.close(); // Close the modal after a delay for visual feedback
                    });
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
        console.log('Opening draft email in Outlook Desktop');

        // Construct the Outlook Desktop URL using the draft ID
        const outlookDesktopUrl = `outlook://message/${draftId}`;

        try {
            // Attempt to open the desktop URL
            window.location.href = outlookDesktopUrl;

            console.log('Attempting to open in Outlook Desktop:', outlookDesktopUrl);
        } catch (error) {
            console.error('Failed to open in Outlook Desktop:', error);
            alert('Unable to open the email in Outlook Desktop. Please ensure Outlook is installed and the protocol is supported.');
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
