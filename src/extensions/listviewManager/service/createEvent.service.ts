import {
    MSGraphClientFactory,
    MSGraphClientV3
} from '@microsoft/sp-http';
import { Constants } from '../models/Constants';
import { EventProperties } from '../models/global.model';
import { IService } from '../components/MeetingInv/models/IService';
import { getSP } from "../../../pnpjs-config";
import { SPFI } from '@pnp/sp';


export class CreateEvent implements IService {
    public context: any;
    public webUri: string;
    public msGraphClientFactory: MSGraphClientFactory;
    public fileNames: string[];
    public fileUris: string[];
    public DocumentIdUrls: string[];
    public ServerRelativeUrl: string;
    public EmailAddress: any;
    private sp: SPFI;
    private static instance: CreateEvent;

    // private constructor(context?: any) {
    //     this.context = context;
    //     this.sp = getSP(this.context);
    // }

    // Return the same object if not changed or a new one
    public static getInstance() {
        if (!CreateEvent.instance) {
            CreateEvent.instance = new CreateEvent();
        }
        return CreateEvent.instance;
    }

    /**
     *  PUBLIC METHODS
     */

    // Set the "DocID" Field of the Document library which holds a temporary copy of the file to read only.
    public SetDocIdReadOnlyField(ReadOnlyFieldValue: boolean): Promise<boolean> {
        const sp = getSP(this.context);
        return new Promise((resolve, reject) => {
            // Get the Field from the list "COPY_DOCUMENT_LIBRARY_NAME" (go to constants in order to watch the list's name) and update ReadOnlyField property to true/false
            sp.web.lists.getByTitle(Constants.COPY_DOCUMENT_LIBRARY_NAME).fields.getByInternalNameOrTitle("DocID").update({ ReadOnlyField: ReadOnlyFieldValue }).then(() => {
                resolve(true);
            })
                .catch((err: any) => {
                    reject(err);
                });
        });
    }

    // Delete the temporary file which was a copy of the original file after he did his purpose (created in order to clean the original file metadata)  
    public DeleteCopiedFile(FileUrisToDelete: string[]): Promise<boolean> {
        const sp = getSP(this.context);
        return new Promise((resolve, reject) => {
            // Create an array of promises for each file deletion
            const deletePromises = FileUrisToDelete.map((fileUri) =>
                sp.web.getFileByUrl(fileUri)
                    .getItem()
                    .then((item: any) => item.delete())
                    .catch((err: string) => {
                        console.error(`Error deleting file ${fileUri}:`, err);
                        return false;
                    })
            );

            // Wait for all delete promises to finish
            Promise.all(deletePromises)
                .then((results) => {
                    // If all files were successfully deleted, resolve with true
                    if (results.every((result: boolean) => result === true)) {
                        resolve(true);
                    } else {
                        reject(new Error('Some files could not be deleted'));
                    }
                })
                .catch((err: any) => {
                    reject(err); // Handle any unexpected errors
                });
        });
    }

    // Copy the file into a mediator Document library in order to clean its metadata
    public CopyFileAndCleanMetadata(fileUris: string[], fileNames: string[], DocumentIdUrls: string[], ServerRelativeUrl: string): Promise<string[]> {
        const sp = getSP(this.context)
        return new Promise((resolve, reject) => {
            if (fileUris.length !== fileNames.length || fileNames.length !== DocumentIdUrls.length) {
                return reject("Input arrays must have the same length.");
            }

            // let web = Web(this.webUri);

            const filePromises = fileUris.map((fileUri: string, index) => {
                const fileName = fileNames[index];
                const documentIdUrl = DocumentIdUrls[index];
                const FilePath = `${ServerRelativeUrl}Shared%20Documents/${fileName}`;

                // Copy each file and update its metadata
                return sp.web.getFileByUrl(fileUri)
                    .copyTo(FilePath, true)
                    .then(() =>
                        sp.web.getFileByUrl(FilePath).getItem()
                    )
                    .then((item: any) =>
                        this.SetDocIdReadOnlyField(false).then((ReadOnlyFieldValue) => {
                            if (ReadOnlyFieldValue) {
                                const CurrentDocumentId = documentIdUrl.split('ID=')[1];
                                return item.update({ DocID: CurrentDocumentId }).then(() =>
                                    this.SetDocIdReadOnlyField(true).then((ResetReadOnlyFieldValue) => {
                                        if (ResetReadOnlyFieldValue) {
                                            return FilePath; // Return the copied file path
                                        }
                                        throw new Error("Failed to reset DocId read-only field.");
                                    })
                                );
                            }
                            throw new Error("Failed to set DocId read-only field to false.");
                        })
                    );
            });

            // Wait for all file operations to complete
            Promise.all(filePromises)
                .then((copiedFilePaths) => resolve(copiedFilePaths))
                .catch((error) => reject(error));
        });
    }

    // Returns the Content of the file Encodes into base64 string
    public getFileContentAsBase64(fileUris: string[]): Promise<string[]> {
        const sp = getSP(this.context);
        return new Promise((resolve, reject) => {
            // let web = Web(this.webUri);

            // יצירת הבטחות לכל קובץ
            const filePromises = fileUris.map((fileUri: string) => {
                return sp.web
                    .getFileByUrl(fileUri)
                    .getBuffer()
                    .then((buffer: ArrayBuffer) => {
                        return this.base64ArrayBuffer(buffer);
                    });
            });

            // המתנה לכל ההבטחות
            Promise.all(filePromises)
                .then((base64Files) => {
                    resolve(base64Files); // מחזיר את המערך של Base64
                })
                .catch((err: any) => {
                    reject(err); // דוחה במקרה של שגיאה
                });
        });
    }

    // Send the email with its content
    public createEvent(eventProperties: EventProperties): Promise<boolean | Error> {

        type Attendee = {
            emailAddress: {
                address: string;
                name?: string;
            },
            type: string;
        }

        type Attachment = {
            "@odata.type": string;
            name: string;
            contentBytes: string;
        }
        // Split emails into arrays
        const attendees: Attendee[] = eventProperties.To.split(';').map(email => {
            return {
                emailAddress: {
                    address: email, // Email address is required
                },
                type: "required",
            }
        })


        const attachments: Attachment[] = eventProperties.Attachment.map(attachment => {
            return {
                "@odata.type": "#microsoft.graph.fileAttachment",
                name: attachment.FileName,
                contentBytes: attachment.ContentBytes
            }
        })
        console.log("createEvent - attachments:", attachments)

        const newEvent = {
            subject: "Quarterly Planning Meeting with Attachments",
            body: {
                contentType: "HTML",
                content: "Please review the attached agenda before the meeting.",
            },
            start: {
                dateTime: "2025-01-01T09:00:00",
                timeZone: "Asia/Jerusalem",
            },
            end: {
                dateTime: "2025-01-01T10:00:00",
                timeZone: "Asia/Jerusalem",
            },
            attendees: attendees,

        };

        // Get the client from sharepoint and make an api call in his name to "MSGraphClient" in order to send the email
        return new Promise((resolve, reject) => {
            this.msGraphClientFactory
                .getClient('3')
                .then((client: MSGraphClientV3) => {
                    client
                        .api(`${Constants.GRAPH_API_BASE_URI}${Constants.GRAPH_API_CREATE_EVENT}`)
                        .post(newEvent)
                        .then((event: any) => {
                            if (attachments && attachments.length > 0) {
                                const attachmentPromises = attachments.map((attachment) => {
                                    client
                                        .api(`${Constants.GRAPH_API_BASE_URI}${Constants.GRAPH_API_CREATE_EVENT}/${event.id}/attachments`)
                                        .post(attachment)
                                })

                                // Wait for all attachments to be added
                                Promise.all(attachmentPromises)
                                    .then(() => {
                                        resolve(true)
                                    })
                                    .catch(() => {
                                        reject(new Error('Event created but failed to add attachments'));
                                    });
                            } else {
                                resolve(true);
                            }
                        })
                        .catch(() => {
                            reject(new Error('Failed to send email'));
                        });
                })
                .catch((error) => {
                    reject(new Error('Failed to get Graph client'));
                });
        });
    }

    /**
     *  PRIVATE METHODS
     */

    // Encodes arrayBuffer into base64 string 
    private base64ArrayBuffer(arrayBuffer: any) {
        var base64 = '';
        var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

        var bytes = new Uint8Array(arrayBuffer);
        var byteLength = bytes.byteLength;
        var byteRemainder = byteLength % 3;
        var mainLength = byteLength - byteRemainder;

        var a, b, c, d;
        var chunk;

        // Main loop deals with bytes in chunks of 3
        for (var i = 0; i < mainLength; i = i + 3) {
            // Combine the three bytes into a single integer
            chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];

            // Use bitmasks to extract 6-bit segments from the triplet
            a = (chunk & 16515072) >> 18; // 16515072 = (2^6 - 1) << 18
            b = (chunk & 258048) >> 12; // 258048   = (2^6 - 1) << 12
            c = (chunk & 4032) >> 6; // 4032     = (2^6 - 1) << 6
            d = chunk & 63;       // 63       = 2^6 - 1

            // Convert the raw binary segments to the appropriate ASCII encoding
            base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
        }

        // Deal with the remaining bytes and padding
        if (byteRemainder == 1) {
            chunk = bytes[mainLength];

            a = (chunk & 252) >> 2; // 252 = (2^6 - 1) << 2

            // Set the 4 least significant bits to zero
            b = (chunk & 3) << 4; // 3   = 2^2 - 1

            base64 += encodings[a] + encodings[b] + '==';
        } else if (byteRemainder == 2) {
            chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];

            a = (chunk & 64512) >> 10; // 64512 = (2^6 - 1) << 10
            b = (chunk & 1008) >> 4; // 1008  = (2^6 - 1) << 4

            // Set the 2 least significant bits to zero
            c = (chunk & 15) << 2; // 15    = 2^4 - 1

            base64 += encodings[a] + encodings[b] + encodings[c] + '=';
        }

        return base64;
    }
}
export default CreateEvent.getInstance();