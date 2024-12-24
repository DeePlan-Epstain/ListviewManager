import * as React from "react";
import "./SendEMailDialogContent.css";
import { Modal, Button, CircularProgress, TextField, Autocomplete, Snackbar, Alert, Box, IconButton } from '@mui/material'
import { ISendEMailDialogContentProps } from "./ISendEMailDialogContentProps";
import { EMailProperties, EMailAttachment } from "../../../models/global.model";
import { ISendEMailDialogContentState } from "./ISendEMailDialogContentState";
import { jss } from "../../../models/jss";
import { cacheRtl } from "../../../models/cacheRtl";
import { StylesProvider } from '@material-ui/core/styles';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import { CacheProvider } from '@emotion/react';
import SendIcon from '@mui/icons-material/Send';
import CloseIcon from '@mui/icons-material/Close';

export class SendEMailDialogContent extends React.Component<ISendEMailDialogContentProps, ISendEMailDialogContentState> {
    private _eMailProperties: EMailProperties;
    private copiedFileUri: string[];

    constructor(props: any) {
        super(props);
        // Set States (information managed within the component), When state changes, the component responds by re-rendering
        this.state = {
            isLoading: false,
            MailOptionTo: "",
            MailOptionCc: "",
            succeed: false,
            MailOptionSubject: this.props.eMailProperties.Subject,
            MailOptionBody: "",
            SendToError: "",
            ESArray: [],
            SendCcError: "",
            SubjectError: "",
            DialogTitle: "שיתוף מסמך / מסמכים עם משתמשים חיצוניים",
            SendEmailFailedError: false,
        };
        this._eMailProperties = this.props.eMailProperties;
        this._submit = this._submit.bind(this);
    }

    componentDidMount() {
        // Check local storage for existing ESArray during the initial mount
        const existingESArray = JSON.parse(localStorage.getItem("ESArray") || "[]");

        // Combine existing ESArray with Contacts from props
        const combinedContacts = [
            ...existingESArray,
            ...this.props.sendDocumentService.EmailAddress,
        ];

        // Update ESArray state with unique contacts from local storage and props
        this.setState({
            ESArray: combinedContacts,
        });
    }

    private _onChangedSubject = (e: any) => {
        this.setState({
            MailOptionSubject: e.target.value,
            SubjectError: "",
        });
        this._eMailProperties.Subject = e.target.value;
    };

    private _onChangedTo = (event: React.ChangeEvent<{}>, value: string[]) => {
        // Email validation regex
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

        // Create a list to store invalid emails
        const invalidEmails = value.filter(email => !emailRegex.test(email));

        // If there are any invalid emails, set an error message
        if (invalidEmails.length > 0) {
            this.setState({
                SendToError: `כתובת אימייל לא תקינה: ${invalidEmails.join(", ")}`,
            });
        } else {
            // Reset the error message if all emails are valid
            this.setState({
                SendToError: "",
            });
        }

        // Always update the MailOptionTo value
        this.setState(
            {
                MailOptionTo: value.join(";"),
            },
            () => {
                //console.log(this.state.MailOptionTo);
            }
        );
        this._eMailProperties.To = value.join(";");
    };

    private _onChangedCc = (event: React.ChangeEvent<{}>, value: string[]) => {
        // Email validation regex
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

        // Create a list to store invalid emails
        const invalidEmails = value.filter(email => !emailRegex.test(email));

        // If there are any invalid emails, set an error message
        if (invalidEmails.length > 0) {
            this.setState({
                SendCcError: `כתובת אימייל לא תקינה: ${invalidEmails.join(", ")}`,
            });
        } else {
            // Reset the error message if all emails are valid
            this.setState({
                SendCcError: "",
            });
        }

        // Always update the MailOptionCc value
        this.setState(
            {
                MailOptionCc: value.join(";"),
            },
            () => {
                //console.log(this.state.MailOptionCc);
            }
        );
        this._eMailProperties.Cc = value.join(";");
    };

    // Triggered every time Body is changed, set MailOptionBody(react state) and _eMailProperties(Class member) to the new value and finally reset the field validation
    private _onChangedBody = (e: any) => {
        this.setState({
            MailOptionBody: e.target.value,
        });
        this._eMailProperties.Body = e.target.value;
    };

    // Returns EMailAttachment object which contains the file name and its Content Encodes into base64 string
    private getEMailAttachment(): Promise<EMailAttachment[]> {
        return new Promise((resolve, reject) => {
            const { sendDocumentService } = this.props;

            // Initialize an empty array to store the copied file URIs
            let copiedFileUris: string[] = [];

            Promise.all(sendDocumentService.fileUris.map((fileUri: any, index: any) =>
                sendDocumentService.CopyFileAndCleanMetadata(
                    [fileUri],
                    [sendDocumentService.fileNames[index]],
                    [sendDocumentService.DocumentIdUrls[index]],
                    sendDocumentService.ServerRelativeUrl
                ).then(async (copiedFileUri: string[]) => {
                    // Add the copied file URIs to the array
                    copiedFileUris = copiedFileUris.concat(copiedFileUri);

                    return sendDocumentService.getFileContentAsBase64(copiedFileUri).then((fileContent: string[]) => {
                        const contentBytes = Array.isArray(fileContent) ? fileContent.join('') : fileContent;

                        return new EMailAttachment({
                            FileName: sendDocumentService.fileNames[index],
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

    // Send email with Attachment
    private sendEMail(eMailProperties: EMailProperties): Promise<boolean> {
        return new Promise((resolve, reject) => {
            this.props.sendDocumentService.sendEMail(eMailProperties)
                .then(() => {
                    resolve(true);
                })
                .catch((err: any) => {
                    reject(err);
                });
        });
    }

    // Validates one email format  
    private ValidateEmail = (mail: string): boolean => {
        var re =
            /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
        return re.test(mail);
    };

    // Validate the Form's fields
    private ValidateForm(eMailProperties: EMailProperties): boolean {
        let Validated = true;

        // Initialize error message variables for 'To' and 'Cc'
        let toErrors: string[] = [];
        let ccErrors: string[] = [];

        // Validate 'To' field
        if (eMailProperties.To.trim() === "") {
            Validated = false;
            this.setState({
                SendToError: "שדה 'אל' לא יכול להישאר ריק",
            });
        } else {
            // Validate emails in 'To' field
            const ToArray = eMailProperties.To.split(";");
            for (let i = 0; i < ToArray.length; i++) {
                const email = ToArray[i].trim();
                if (email !== "" && !this.ValidateEmail(email)) {
                    toErrors.push(email); // Collect invalid emails
                    Validated = false;
                }
            }
            // If there are invalid emails, set the error message
            if (toErrors.length > 0) {
                this.setState({
                    SendToError: `אחת או יותר מכתובות הדואר האלקטרוני ב-'אל' שגויות: ${toErrors.join(", ")}`,
                });
            }
        }

        // Validate 'Cc' field
        if (eMailProperties.Cc.trim() !== "") {
            const CcArray = eMailProperties.Cc.split(";");
            for (let i = 0; i < CcArray.length; i++) {
                const email = CcArray[i].trim();
                if (email !== "" && !this.ValidateEmail(email)) {
                    ccErrors.push(email); // Collect invalid emails
                    Validated = false;
                }
            }
            // If there are invalid emails, set the error message
            if (ccErrors.length > 0) {
                this.setState({
                    SendCcError: `אחת או יותר מכתובות הדואר האלקטרוני ב-'מכותבים' שגויות: ${ccErrors.join(", ")}`,
                });
            }
        }

        // Validate 'Subject' field
        if (eMailProperties.Subject.trim() === "") {
            this.setState({
                SubjectError: "שדה 'נושא' לא יכול להישאר ריק",
            });
            Validated = false;
        }

        return Validated;
    }

    // Submit the form
    private _submit() {
        // Validate the Form
        if (this.ValidateForm(this._eMailProperties)) {
            // Activate spinner
            this.setState({ isLoading: true, succeed: false });  // Reset success to false
            // Get the Content of the file Encodes into base64 string
            this.getEMailAttachment().then((attachments: EMailAttachment[]) => {
                this._eMailProperties.Attachment = attachments;
                this.sendEMail(this._eMailProperties)
                    .then(() => {
                        this.setState({ succeed: true, isLoading: false });
                        this.props.sendDocumentService.DeleteCopiedFile(this.copiedFileUri);
                        setTimeout(() => {
                            this.props.close(); // Close the modal after a delay for visual feedback
                        }, 1000);
                    })
                    .catch((err) => {
                        console.error("Send Document Error", err);
                        this.setState({
                            SendEmailFailedError: true,
                            isLoading: false,
                        });
                    });
            });
        }
    }

    public render(): React.ReactElement<ISendEMailDialogContentProps> {

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
                            zIndex: '9999999',
                        }}
                    >
                        <div
                            className="ModalContentContainer"
                            dir="rtl"
                            style={{
                                width: '450px',
                                background: 'white',
                                padding: '20px',
                                borderRadius: '5px',
                                boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                            }}
                        >
                            <span id="modal-title">{this.state.DialogTitle}</span>
                            <div className="top-spacer" />
                            <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "0px 2px" }}>אל:</span>
                            <div className="top-spacerLabel" />
                            <div className="SendDocumentContainer" dir="rtl">
                                <Autocomplete
                                    onChange={(event, value) => this._onChangedTo(event, value)}
                                    dir="rtl"
                                    disablePortal
                                    multiple
                                    freeSolo
                                    disabled={this.state.isLoading || this.state.succeed}
                                    options={this.state.ESArray}
                                    ListboxProps={{ style: { maxHeight: '15rem', background: "white" } }}
                                    renderInput={(params: any) => (
                                        <TextField
                                            dir="rtl"
                                            type="email"
                                            fullWidth
                                            size="small"
                                            {...params}
                                            //label="אל"
                                            sx={{
                                                '& .MuiOutlinedInput-root': {
                                                    padding: '0px !important',
                                                }
                                            }}
                                            required={true}
                                            helperText={
                                                <span style={{ color: 'red' }}>{this.state.SendToError}</span>
                                            }
                                        />
                                    )}
                                />
                                <div className="top-spacer" />
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>מכותבים:</span>
                                <div className="top-spacerLabel" />
                                <Autocomplete
                                    multiple
                                    disablePortal
                                    freeSolo
                                    disabled={this.state.isLoading || this.state.succeed}
                                    ListboxProps={{ style: { maxHeight: '15rem' } }}
                                    onChange={(event, value) => this._onChangedCc(event, value)}
                                    options={this.state.ESArray}
                                    renderInput={(params: any) => (
                                        <TextField
                                            dir="rtl"
                                            type="email"
                                            {...params}
                                            //label="מכותבים"
                                            sx={{
                                                '& .MuiOutlinedInput-root': {
                                                    padding: '0px !important',
                                                }
                                            }}
                                            helperText={
                                                <span style={{ color: 'red' }}>{this.state.SendCcError}</span>
                                            }
                                        />
                                    )}
                                />
                                <div className="top-spacer" />
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>נושא:</span>
                                <div className="top-spacerLabel" />
                                <TextField
                                    style={{ direction: 'rtl' }}
                                    helperText={
                                        <span style={{ color: 'red' }}>{this.state.SubjectError}</span>
                                    }
                                    onChange={this._onChangedSubject}
                                    //label="נושא"
                                    disabled={this.state.isLoading || this.state.succeed}
                                    name="MailOptionSubject"
                                    required={true}
                                    value={this.state.MailOptionSubject}
                                    fullWidth
                                    size="small"
                                />
                                <div className="top-spacer" />
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>תוכן המייל:</span>
                                <div className="top-spacerLabel" />
                                <TextField
                                    style={{ direction: 'rtl' }}
                                    onChange={this._onChangedBody}
                                    //label="תוכן המייל"
                                    disabled={this.state.isLoading || this.state.succeed}
                                    name="MailOptionBody"
                                    multiline
                                    minRows={3}
                                    maxRows={3}
                                    value={this.state.MailOptionBody}
                                    fullWidth
                                    size="small"
                                />
                                <div className="top-spacer" />
                            </div>
                            <div className="ModalFooter">
                                <Box
                                    sx={{
                                        display: 'flex',
                                        justifyContent: 'end',
                                        gap: '16px',
                                    }}
                                >
                                    <Button
                                        color="error"
                                        disabled={this.state.isLoading || this.state.succeed}
                                        onClick={this.props.close}
                                        startIcon={ <IconButton disableRipple style={{color: "#f58383", paddingLeft: 0, margin: "0px !important"}}><CloseIcon /></IconButton>}
                                        sx={{
                                            "& .MuiButton-startIcon": {
                                                margin: 0, // Removes default margin
                                            },
                                        }}
                                    >
                                        ביטול
                                    </Button>
                                    <Button
                                        onClick={this._submit}
                                        disabled={this.state.isLoading || this.state.succeed}
                                        endIcon={(!this.state.isLoading || !this.state.succeed) && <IconButton disableRipple style={{ transform: "rotate(180deg)", color:"#1976d2", "padding": 0 }} ><SendIcon/></IconButton>}
                                        startIcon={this.state.isLoading ? (
                                            <CircularProgress size={20} color="inherit" />
                                        ) : (
                                            this.state.succeed && <CheckCircleIcon />
                                        )}
                                    >
                                        {this.state.isLoading ? 'שליחה...' : (this.state.succeed ? 'נשלח' : 'שליחת מייל')}
                                    </Button>
                                </Box>
                            </div>

                        </div>
                    </Modal>
                </StylesProvider>
            </CacheProvider >

        );
    }
}