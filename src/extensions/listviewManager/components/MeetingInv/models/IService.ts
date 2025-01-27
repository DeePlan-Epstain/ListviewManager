import { EventProperties } from "../../../models/global.model";
import { MSGraphClientFactory } from "@microsoft/sp-http";

// Service members and methods
export interface IService {
    context: any;
    webUri: string;
    msGraphClientFactory: MSGraphClientFactory;
    fileNames: string[];
    fileUris: string[];
    DocumentIdUrls: string[];
    ServerRelativeUrl: string;
    EmailAddress: string[];
    createEvent(eventProperties: EventProperties): Promise<string | Error>;
    getFileContentAsBase64(fileUris: string[]): Promise<string[]>;
    CopyFileAndCleanMetadata(fileUris: string[], fileNames: string[], DocumentIdUrls: string[], ServerRelativeUrl: string): Promise<string[]>;
    DeleteCopiedFile(fileUri: string[]): Promise<boolean>;
}