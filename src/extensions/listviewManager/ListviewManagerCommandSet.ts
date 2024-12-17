import React from "react";
import ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import { getSP, getGraph } from "../../pnpjs-config";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { SelectedFile } from "./models/global.model";
import ApproveDocument, {
  ApproveDocumentProps,
} from "./components/ApproveDocument/ApproveDocument.cmp";
import ModalExt from "../../extensions/listviewManager/components/FolderHierarchy/ModalExt";
import MoveFile, { MoveFileProps } from "./components/MoveFile/MoveFile.cmp";
// import RenameFile, {
//   IRenameFileProps,
// } from "./components/RenameFile/RenameFile";
import { PermissionKind } from "@pnp/sp/security";
import { decimalToBinaryArray } from "./service/util.service";
import { ModalExtProps } from "./components/FolderHierarchy/ModalExtProps";
import { ConvertToPdf, getConvertibleTypes } from "./service/pdf.service";
import { GraphFI } from "@pnp/graph";
import SendDocumentService from "./service/sendDocument.service";
import SendEMailDialog from "./components/ExternalSharing/SendEMailDialog/SendEMailDialog";

const { solution } = require("../../../config/package-solution.json");

export interface IListviewManagerCommandSetProperties {
  sampleTextOne: string;
}

const LOG_SOURCE: string = "ListviewManagerCommandSet";

export default class ListviewManagerCommandSet extends BaseListViewCommandSet<IListviewManagerCommandSetProperties> {
  private dialogContainer: HTMLDivElement;
  private sp: SPFI;
  private graph: GraphFI;
  private currUser: ISiteUserInfo;
  private isAllowedToMoveFile: boolean = false;
  private typeSet: Set<string>;

  private allowedUsers: string[] = [
    "EpsteinSystem@Epstein.co.il",
  ].map((e) => e.toLocaleLowerCase());

  public async onInit(): Promise<void> {
    console.log(solution.name + ":", solution.version);
    Log.info(LOG_SOURCE, "Initialized ListviewManagerCommandSet");
    this.sp = getSP(this.context);
    this.graph = getGraph(this.context);
    this.isAllowedToMoveFile = await this._checkUserPermissionToMoveFile();
    console.log("this.isAllowedToMoveFile", this.isAllowedToMoveFile);


    this.currUser = await this.sp.web.currentUser();
    if (!this.allowedUsers.includes(this.currUser.Email.toLocaleLowerCase())) {
      this.startDnDBlock();
    }
    this.dialogContainer = document.body.appendChild(
      document.createElement("div")
    );

    const compareOneCommand: Command = this.tryGetCommand("Approval_Document");
    compareOneCommand.visible = false;
    const compareTwoCommand: Command = this.tryGetCommand("folderHierarchy");
    if (this.isAllowedToMoveFile === false) {
      compareTwoCommand.visible = false;
    }
    const compareFiveCommand: Command = this.tryGetCommand("convertToPDF");
    compareFiveCommand.visible = false;
    // const compareThreeCommand: Command = this.tryGetCommand("Move_File");
    // compareThreeCommand.visible = false;
    // const compareFourCommand: Command = this.tryGetCommand("RenameFile");
    // compareFourCommand.visible = false;

    // File sharing by email
    const externalSharingCompareOneCommand: Command = this.tryGetCommand("External_Sharing");
    externalSharingCompareOneCommand.visible = false;

    const isUserAllowed = this.allowedUsers.includes(this.currUser.Email);
    if (!isUserAllowed) {
      require("./styles/createNewFolder.module.scss"); // hide the button create new folder if the user is not allowed
    }
    this.typeSet = await getConvertibleTypes(this.context);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {

    const fullUrl = window.location.href;
    console.log(fullUrl);

    // Extract the "id" parameter from the query string
    const urlParams = new URLSearchParams(fullUrl.split('?')[1]);
    const folderPath = urlParams.get('id');

    // If "id" exists, decode it to get the relative path
    let finalPath = folderPath ? decodeURIComponent(folderPath) : null;

    if (finalPath) {
      console.log("Final Server Relative URL:", finalPath);
    } else {

      finalPath = fullUrl.split('/Forms')[0] // Assuming "Forms" is in the URL structure
    }
    console.log(finalPath);


    const selectedFiles = event.selectedRows.map((SR: any) => {
      const keys = SR._values.keys();
      const row: any = {};
      Array.from(keys).forEach(
        (key: any) => (row[key] = SR.getValueByName(key))
      );
      return row;
    });



    switch (event.itemId) {
      case "Approval_Document":
        this._renderApproveDocumentModal(selectedFiles[0], "Approval");
        break;
      case "folderHierarchy":
        this._renderfolderHierarchytModal(finalPath, "Approval");
        break;
      case "convertToPDF":
        ConvertToPdf(this.context, selectedFiles[0])
        break;
      case "External_Sharing":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.processSelectedRowsAndContacts(Array.from(event.selectedRows));
        }
        break;
      // case "Move_File":
      //   this._renderMoveFileModal(selectedFiles);
      //   break;
      // case "RenameFile":
      //   if (libraryName && libraryID) {
      //     this._renderRenameFileModal(selectedFiles[0], libraryName, libraryID);
      //   }
      //   break;
      default:
        throw new Error("Unknown command");
    }
  }
  private startDnDBlock() {
    const dropZoneArea = document.querySelectorAll("[role=presentation]");
    const testArr: Element[] = [];

    dropZoneArea.forEach((dz) => {
      if (
        dz.className.includes("root") &&
        dz.className.includes("absolute") &&
        dz.attributes.getNamedItem("data-drop-target-key")
      )
        testArr.push(dz);
    });

    testArr[0].addEventListener("drop", this.preventFolderDrop, false);
  }

  private preventFolderDrop(event: DragEvent): void {
    // Check if the item being dragged is a folder
    if (event?.dataTransfer?.items?.length) {
      for (let i = 0; i < event.dataTransfer.items.length; i++) {
        const item = event.dataTransfer.items[i].webkitGetAsEntry();
        if (item && item.isDirectory) {
          event.preventDefault();
          event.stopPropagation();
          alert("Folder upload is not allowed.");
          return;
        }
      }
    }
  }

  extractLibraryDetails = async (
    fileRef: string
  ): Promise<{ libraryName: string; libraryID: string }> => {
    console.log(fileRef);

    const parts = fileRef.split("/");
    const libraryUrlPart = parts.length > 2 ? parts[3] : "";

    // Fetch all libraries and find the one matching the extracted URL part
    const libraries = await this.sp.web.lists
      .select("Id", "Title", "RootFolder/Name")
      .expand("RootFolder")();
    const library = libraries.find(
      (lib: any) => lib.RootFolder.Name === libraryUrlPart
    );

    if (library) {
      return { libraryName: libraryUrlPart, libraryID: library.Id };
    } else {
      throw new Error(`Library with URL part '${libraryUrlPart}' not found`);
    }
  };

  private _closeDialogContainer = () => {
    ReactDom.unmountComponentAtNode(this.dialogContainer!);
  };

  private async _checkUserPermissionToMoveFile(): Promise<boolean> {
    try {
      return await this.sp.web.lists
        .getById(this.context.pageContext.list.id["_guid"])
        .currentUserHasPermissions(PermissionKind.EditListItems);
    } catch (error) {
      console.error("Error while checking user permission to move file", error);
    }
  }

  private _renderMoveFileModal(selectedRows: any[]) {
    const element: React.ReactElement<MoveFileProps> = React.createElement(
      MoveFile,
      {
        sp: this.sp,
        context: this.context,
        selectedRows,
        unMountDialog: this._closeDialogContainer,
      }
    );

    ReactDom.render(element, this.dialogContainer);
  }

  // private _renderRenameFileModal(
  //   selectedRow: any,
  //   libraryName: any,
  //   libraryID: any
  // ) {
  //   const element: React.ReactElement<IRenameFileProps> = React.createElement(
  //     RenameFile,
  //     {
  //       sp: this.sp,
  //       context: this.context,
  //       selectedRow,
  //       libraryName,
  //       libraryID,
  //       unMountDialog: this._closeDialogContainer,
  //     }
  //   );
  //   ReactDom.render(element, this.dialogContainer);
  // }

  private _renderApproveDocumentModal(selectedRow: any, modalInterface: "Review" | "Approval") {
    const element: React.ReactElement<ApproveDocumentProps> =
      React.createElement(ApproveDocument, {
        sp: this.sp,
        context: this.context as any,
        selectedRow,
        modalInterface,
        currUser: this.currUser,
        unMountDialog: this._closeDialogContainer,
      });

    ReactDom.render(element, this.dialogContainer);
  }

  private _renderfolderHierarchytModal(finalPath: any, modalInterface: "Review" | "Approval") {
    const element: React.ReactElement<ModalExtProps> =
      React.createElement(ModalExt, {
        sp: this.sp,
        context: this.context as any,
        finalPath,
        modalInterface,
        currUser: this.currUser,
        unMountDialog: this._closeDialogContainer,
      });

    ReactDom.render(element, this.dialogContainer);
  }

  private async processSelectedRowsAndContacts(selectedRows: any[]): Promise<void> {
    // Initialize arrays to store file information
    const fileNames: string[] = [];
    const fileRefs: string[] = [];
    const documentIdUrls: string[] = [];

    // Iterate through selected rows to gather file information
    selectedRows.forEach(row => {
      const fileName = row.getValueByName("FileLeafRef").toString();
      const fileRef = row.getValueByName("FileRef").toString();
      const documentIdUrl = row.getValueByName("ServerRedirectedEmbedUrl").toString();

      fileNames.push(fileName);
      fileRefs.push(fileRef);
      documentIdUrls.push(documentIdUrl);
    });

    // Retrieve user contacts
    const contact = await this.graph.me.contacts();
    const emails = contact.flatMap((c: any) => c.emailAddresses.map((email: any) => email.address));

    // Update SendDocumentService properties
    SendDocumentService.EmailAddress = emails;
    SendDocumentService.fileNames = fileNames;
    SendDocumentService.fileUris = fileRefs;
    SendDocumentService.DocumentIdUrls = documentIdUrls;
    SendDocumentService.webUri = this.context.pageContext.web.absoluteUrl;
    SendDocumentService.context = this.context;

    // Set MS Graph client factory
    if (this.context && this.context.msGraphClientFactory) {
      SendDocumentService.msGraphClientFactory = this.context.msGraphClientFactory;
    } else {
      console.error("MSGraphClientFactory is undefined.");
      return;
    }

    // Set server relative URL
    const currentRelativeUrl = this.context.pageContext.site.serverRelativeUrl;
    SendDocumentService.ServerRelativeUrl = currentRelativeUrl;

    // Create and display the email dialog
    const dialog: SendEMailDialog = new SendEMailDialog(SendDocumentService);
    await dialog.show();
  }

  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    Log.info(LOG_SOURCE, "List view state changed");

    let LibraryName = this.context.pageContext.list.title;

    const compareOneCommand: Command = this.tryGetCommand("Approval_Document");
    // const compareThreeCommand: Command = this.tryGetCommand("Move_File");
    // const compareFourCommand: Command = this.tryGetCommand("RenameFile");
    const compareFiveCommand: Command = this.tryGetCommand("convertToPDF");
    const externalSharingCompareOneCommand: Command = this.tryGetCommand("External_Sharing");


    if (compareOneCommand) {
      compareOneCommand.visible = event.selectedRows?.length === 1 && event.selectedRows[0]?.getValueByName('FSObjType') == 0
    }

    // if there is only one selected item and its a file and its a file type that can be converted to pdf
    if (compareFiveCommand) {
      compareFiveCommand.visible = event.selectedRows?.length === 1 &&
        event.selectedRows[0]?.getValueByName('FSObjType') == 0 &&
        this.typeSet.has(event.selectedRows[0]?.getValueByName(".fileType"));
    }

    // if there is one selected item or more and its a file
    if (externalSharingCompareOneCommand) {
      if (event.selectedRows?.length > 0) {
        const fileExt = event.selectedRows[0].getValueByName(".fileType")
        if (fileExt.toLowerCase() !== "") externalSharingCompareOneCommand.visible = true;
      } else externalSharingCompareOneCommand.visible = false;
    }

    // if (compareThreeCommand) {
    //   if (event.selectedRows?.length > 0) {
    //     const isFolder = Boolean(
    //       event.selectedRows.find(
    //         (r) => r.getValueByName("ContentType") === "Folder"
    //       )
    //     );
    //     compareThreeCommand.visible = !isFolder;
    //   } else compareThreeCommand.visible = false;
    // }

    // if (compareFourCommand) {
    //   compareFourCommand.visible = event.selectedRows?.length === 1;
    // }

    // const compareTwoCommand: Command = this.tryGetCommand("folderHierarchy");
    // if (compareTwoCommand) {
    //   compareTwoCommand.visible = event.selectedRows?.length === 1;
    // }
  }
}
