import * as React from 'react';
import { EMailAttachment, EventProperties } from '../../models/global.model';
import { IMeetingInvProps } from './IMeetingInvProps';
import { jss } from "../../models/jss";
import { cacheRtl } from "../../models/cacheRtl";
import { StylesProvider } from '@material-ui/core/styles';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import { CacheProvider } from '@emotion/react';
import SendIcon from '@mui/icons-material/Send';
import CloseIcon from '@mui/icons-material/Close';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { Modal, Button, CircularProgress, TextField, Autocomplete, Snackbar, Alert, Box, IconButton, Switch } from '@mui/material'
import { DatePicker } from '@mui/x-date-pickers/DatePicker';
import { TimePicker } from '@mui/x-date-pickers/TimePicker';
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider';
import { AdapterMoment } from '@mui/x-date-pickers/AdapterMoment';
import moment, { Moment } from 'moment';
import styles from './MeetingInv.module.scss'

export interface IMeetingInvState {
    isLoading: boolean;
    DialogTitle: string;
    MailOptionTo: string;
    MailOptional: string;
    MailOptionSubject: string;
    MailOptionBody: string;
    SendToError: string;
    SendOptinalsError: string;
    SubjectError: string;
    SendEmailFailedError: boolean;
    succeed?: boolean;
    ESArray: string[],
    error?: Error;
    date: Moment,
    startTime: any,
    endTime: any,
    dateAndTimeError: string
    onlineMeeting: boolean
    link: string
}

export default class MeetingInv extends React.Component<IMeetingInvProps, IMeetingInvState> {

    private _eventProperties: EventProperties;
    private copiedFileUri: string[];

    constructor(props: IMeetingInvProps) {
        super(props);

        this.state = {
            isLoading: false,
            MailOptionTo: "",
            MailOptional: "",
            succeed: false,
            MailOptionSubject: this.props.eventProperties.Subject,
            MailOptionBody: "",
            SendToError: "",
            ESArray: [],
            SendOptinalsError: "",
            SubjectError: "",
            DialogTitle: "זימון פגישה",
            SendEmailFailedError: false,
            date: moment(),
            startTime: this.props.eventProperties.startTime,
            endTime: this.props.eventProperties.endTime,
            dateAndTimeError: "",
            onlineMeeting: false,
            link: '',
        };
        this._eventProperties = this.props.eventProperties;
        this._submit = this._submit.bind(this);
    }

    componentDidMount() {

        // Combine existing ESArray with Contacts from props
        const combinedContacts = [
            ...this.props.createEvent.EmailAddress,
        ];

        // Update ESArray state with unique contacts from local storage and props
        this.setState({
            ESArray: combinedContacts,
        });
        this._onStart()
    }

    componentWillUnmount(): void {
        window.open(this.state.link, '_blank');
    }

    private _onChangedSubject = (e: any) => {
        this.setState({
            MailOptionSubject: e.target.value,
            SubjectError: "",
        });
        this._eventProperties.Subject = e.target.value;
    };

    private _onChangedTo = (event: React.ChangeEvent<{}>, value: string[]) => {
        // Email validation regex
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

        // Create a list to store invalid emails
        const invalidEmails = value.filter(email => !emailRegex.test(email));

        // If there are any invalid emails, set an error message
        if (!invalidEmails.length) {
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
        this._eventProperties.To = value.join(";");
    };

    private _onChangedMailOptional = (event: React.ChangeEvent<{}>, value: string[]) => {
        // Email validation regex
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

        // Create a list to store invalid emails
        const invalidEmails = value.filter(email => !emailRegex.test(email));

        // If there are any invalid emails, set an error message
        if (!invalidEmails.length) {
            // Reset the error message if all emails are valid
            this.setState({
                SendOptinalsError: "",
            });
        }

        // Always update the MailOptionCc value
        this.setState(
            {
                MailOptional: value.join(";"),
            },
            () => {
                //console.log(this.state.MailOptionCc);
            }
        );
        this._eventProperties.optionals = value.join(";");
    };

    // Triggered every time Body is changed, set MailOptionBody(react state) and _eventProperties(Class member) to the new value and finally reset the field validation
    private htmlToPlainText(html: any) {
        const tempElement = document.createElement('div');
        tempElement.innerHTML = html;
        return tempElement.innerText;
    }

    // Triggered every time Body is changed, set MailOptionBody(react state) and _eventProperties(Class member) to the new value and finally reset the field validation
    private _onChangedBody = (newValue: string): string => {
        const plainText = this.htmlToPlainText(newValue);
        this.setState({
            MailOptionBody: newValue,
        });
        this._eventProperties.Body = newValue;
        return newValue;
    };


    // Returns EMailAttachment object which contains the file name and its Content Encodes into base64 string
    private getEMailAttachment(): Promise<EMailAttachment[]> {
        return new Promise((resolve, reject) => {
            const { createEvent } = this.props;

            // Initialize an empty array to store the copied file URIs
            let copiedFileUris: string[] = [];

            Promise.all(createEvent.fileUris.map((fileUri: any, index: any) =>
                createEvent.CopyFileAndCleanMetadata(
                    [fileUri],
                    [createEvent.fileNames[index]],
                    [createEvent.DocumentIdUrls[index]],
                    createEvent.ServerRelativeUrl
                ).then(async (copiedFileUri: string[]) => {
                    // Add the copied file URIs to the array
                    copiedFileUris = copiedFileUris.concat(copiedFileUri);

                    return createEvent.getFileContentAsBase64(copiedFileUri).then((fileContent: string[]) => {
                        const contentBytes = Array.isArray(fileContent) ? fileContent.join('') : fileContent;

                        return new EMailAttachment({
                            FileName: createEvent.fileNames[index],
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

    private createEvent(EventProperties: EventProperties): Promise<string> {
        return new Promise((resolve, reject) => {
            this.props.createEvent.createEvent(EventProperties)
                .then((link: string) => {
                    resolve(link);
                })
                .catch((err: any) => {
                    reject(err);
                });
        })
    }

    // Validates one email format  
    private ValidateEmail = (mail: string): boolean => {
        var re =
            /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
        return re.test(mail);
    };

    // Validate the Form's fields
    private ValidateForm(EventProperties: EventProperties): boolean {
        let Validated = true;
        const { date, startTime, endTime } = this.state
        // Initialize error message variables for 'To' and 'Cc'
        let toErrors: string[] = [];
        let optionalsErrors: string[] = [];

        // Validate 'To' field
        if (EventProperties.To.trim() === "") {
            Validated = false;
            this.setState({
                SendToError: "שדה 'משתתפים' לא יכול להישאר ריק",
            });
        } else {
            // Validate emails in 'To' field
            const ToArray = EventProperties.To.split(";");
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
                    SendToError: `אחת או יותר מכתובות הדואר האלקטרוני ב-'משתתפים' שגויות: ${toErrors.join(", ")}`,
                });
            }
        }

        // Validate 'optionals'
        const ToArray = EventProperties.optionals.split(";");
        for (let i = 0; i < ToArray.length; i++) {
            const email = ToArray[i].trim();
            if (email !== "" && !this.ValidateEmail(email)) {
                optionalsErrors.push(email); // Collect invalid emails
                Validated = false;
            }
        }
        // If there are invalid emails, set the error message
        if (optionalsErrors.length > 0) {
            this.setState({
                SendOptinalsError: `אחת או יותר מכתובות הדואר האלקטרוני ב-'משתתפים אופציונלים' שגויות: ${optionalsErrors.join(", ")}`,
            });
        }

        // Validate 'Subject' field
        if (EventProperties.Subject.trim() === "") {
            this.setState({
                SubjectError: "שדה 'נושא' לא יכול להישאר ריק",
            });
            Validated = false;
        }

        if (date.isValid() === false || (startTime === "Invalid date" || startTime === "") || (endTime === "Invalid date" || endTime === "")) {
            this.setState({
                dateAndTimeError: "אחד מן השדות 'תאריך', 'משעה' או 'עד שעה' ריקים"
            })
            Validated = false
        } else {
            this.setState({
                dateAndTimeError: ""
            })
        }

        return Validated;
    }

    public validationDateAndTimeOnChange() {
        const { date, startTime, endTime } = this.state

        if (date.isValid() === false || (startTime === "Invalid date" || startTime === "") || (endTime === "Invalid date" || endTime === "")) {
            this.setState({
                dateAndTimeError: "אחד מן השדות 'תאריך', 'משעה' או 'עד שעה' ריקים"
            })
        } else {
            this.setState({
                dateAndTimeError: ""
            })
        }
    }

    public formatMeetingTimeUTC = (date: any, time: any): string => {
        const [hours, minutes] = time.split(':')
        return moment(date)
            .set({
                hour: parseInt(hours),
                minute: parseInt(minutes),
                second: 0,
                millisecond: 0,
            })
            .format('YYYY-MM-DDTHH:mm:ss');
    };

    // Submit the form
    public _submit() {
        if (this.ValidateForm(this._eventProperties)) {

            const { date, startTime, endTime, onlineMeeting } = this.state

            // Activate spinner
            this.setState({ isLoading: true, succeed: false });  // Reset success to false

            const startTimeDate = this.formatMeetingTimeUTC(date, startTime)
            const endTimeDate = this.formatMeetingTimeUTC(date, endTime)

            this._eventProperties.startTime = startTimeDate
            this._eventProperties.endTime = endTimeDate
            this._eventProperties.onlineMeeting = onlineMeeting

            this.getEMailAttachment().then((attachments: EMailAttachment[]) => {
                this._eventProperties.Attachment = attachments;
                this.createEvent(this._eventProperties)
                    .then(() => {
                        this.setState({ succeed: true, isLoading: false });
                        this.props.createEvent.DeleteCopiedFile(this.copiedFileUri).then(() => {
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

    }

    // _onStart without the form
    public _onStart() {
        this.getEMailAttachment().then((attachments: EMailAttachment[]) => {
            this._eventProperties.Attachment = attachments;
            this.createEvent(this._eventProperties)
                .then((link: string) => {
                    this.props.createEvent.DeleteCopiedFile(this.copiedFileUri).then(() => {
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
                        {/* <div
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
                            <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "0px 2px" }}>משתתפים:</span>
                            <div className="top-spacerLabel" />
                            <div className="SendDocumentContainer" id={styles.containerSmall} dir="rtl">
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
                                            helperText={this.state.SendToError ?
                                                <span className="errorSpan">{this.state.SendToError}</span> : <span>לאחר הקלדת אימייל, לחץ על מקש Enter.</span>
                                            }
                                        />
                                    )}
                                />
                                <div className="top-spacer" />
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "0px 2px" }}>משתתפים אופציונלים:</span>
                                <div className="top-spacerLabel" />
                                <div className="SendDocumentContainer" dir="rtl">
                                    <Autocomplete
                                        onChange={(event, value) => this._onChangedMailOptional(event, value)}
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
                                                helperText={this.state.SendOptinalsError ?
                                                    <span className="errorSpan">{this.state.SendOptinalsError}</span> : <span>לאחר הקלדת אימייל, לחץ על מקש Enter.</span>
                                                }
                                            />
                                        )}
                                    />
                                </div>
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
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>תאריך ושעה:</span>
                                <div className="top-spacerLabel" style={{ paddingBottom: '1em' }} />
                                <div style={{ display: 'flex', flexDirection: 'column', }}>
                                    <LocalizationProvider dateAdapter={AdapterMoment} adapterLocale="he">
                                        <DatePicker
                                            sx={{ paddingBottom: '1em' }}
                                            label="תאריך"
                                            disablePast
                                            value={this.state.date}
                                            onChange={(newValue: Moment) => {
                                                this.setState({ date: moment(newValue).startOf('day') },
                                                    () => this.validationDateAndTimeOnChange()); // Strip time component
                                            }}
                                        />
                                        <div style={{ display: 'flex', flexDirection: 'row', gap: '1em' }}>
                                            <TimePicker
                                                ampm={false}
                                                label="משעה"
                                                disablePast
                                                value={this.state.startTime ? moment(this.state.startTime, 'HH:mm') : null}
                                                maxTime={this.state.endTime ? moment(this.state.endTime, 'HH:mm') : undefined}
                                                onChange={(newValue: Moment) => {
                                                    this.validationDateAndTimeOnChange()
                                                    this.setState({ startTime: moment(newValue).format('HH:mm') },
                                                        () => this.validationDateAndTimeOnChange()); // Ensure time in "HH:mm"
                                                }}
                                            />
                                            <TimePicker
                                                ampm={false}
                                                label="עד שעה"
                                                disablePast
                                                value={this.state.endTime ? moment(this.state.endTime, 'HH:mm') : null}
                                                minTime={this.state.startTime ? moment(this.state.startTime, 'HH:mm') : undefined}
                                                onChange={(newValue: Moment) => {
                                                    this.setState({ endTime: moment(newValue).format('HH:mm') },
                                                        () => this.validationDateAndTimeOnChange()); // Ensure time in "HH:mm"
                                                }}
                                            />
                                        </div>
                                    </LocalizationProvider>
                                    <div style={{ display: 'flex', width: '100%', paddingTop: '5px' }}>
                                        <span className={styles.errorMessage} style={{ color: 'red' }}>{this.state.dateAndTimeError}</span>
                                    </div>
                                </div>

                                <div className="top-spacer" />
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>פגישה מקוונת:</span>
                                <div className="top-spacerLabel" />
                                <Switch onClick={() => this.setState({ onlineMeeting: !this.state.onlineMeeting })}></Switch>

                                <div className="top-spacer" />
                                <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>תוכן המייל:</span>
                                <div className="top-spacerLabel" />
                                <RichText
                                    value={this.state.MailOptionBody}
                                    isEditMode={!this.state.isLoading && !this.state.succeed}
                                    className="richTextEditor"
                                    onChange={(newValue: string) => this._onChangedBody(newValue)}
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
                                        startIcon={<IconButton disableRipple style={{ color: this.state.isLoading || this.state.succeed ? 'inherit' : "#f58383", paddingLeft: 0, margin: "0px !important" }}><CloseIcon /></IconButton>}
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
                                        endIcon={(!this.state.isLoading || !this.state.succeed) && <IconButton disableRipple style={{
                                            transform: "rotate(180deg)",
                                            color: this.state.isLoading || this.state.succeed ? 'inherit' : "#1976d2",
                                            "padding": 0
                                        }} ><SendIcon /></IconButton>}
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

                        </div> */}
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
