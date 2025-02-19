import React from "react";
import ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import { getSP, getGraph, getSPByPath } from "../../pnpjs-config";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { EMailProperties, EventProperties, DraftProperties, SelectedFile } from "./models/global.model";
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
import LinkToCategory, { LinkToCategoryProps } from "./components/LinkToCategory/LinkToCategory";
import { ConvertToPdf, getConvertibleTypes } from "./service/pdf.service";
import { GraphFI } from "@pnp/graph";
import SendDocumentService from "./service/sendDocument.service";
import CreateEvent from "./service/createEvent.service";
import ExportZipModal from "./components/ExportZip/ExportZip.cmp";
import Swal from 'sweetalert2';
import { ISendEMailDialogContentProps } from "./components/ExternalSharing/SendEMailDialogContent/ISendEMailDialogContentProps";
import { SendEMailDialogContent } from "./components/ExternalSharing/SendEMailDialogContent/SendEMailDialogContent";
import MeetingInv from "./components/MeetingInv/MeetingInv";
import { IMeetingInvProps } from "./components/MeetingInv/IMeetingInvProps";
import { IDraftProps } from "./components/Draft/IDraftProps";
import Draft from "./components/Draft/Draft.cmp"
import CreateDraft from "./service/createDraft.service";
import toast, { Toaster } from 'react-hot-toast'; // Importing react-hot-toast
import { spfi, SPFx } from "@pnp/sp";
import './../ext.css'

const { solution } = require("../../../config/package-solution.json");

export interface IListviewManagerCommandSetProperties {
  sampleTextOne: string;
}

const LOG_SOURCE: string = "ListviewManagerCommandSet";
const FAVORITES_LIST_ID = '6f3d6257-4a9b-41fe-a847-487c942cd628'

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

  private favorites: any[] = []
  private spPortal: SPFI = null

  public async onInit(): Promise<void> {
    console.log(solution.name + ":", solution.version);
    Log.info(LOG_SOURCE, "Initialized ListviewManagerCommandSet");
    this.sp = getSP(this.context);
    this.graph = getGraph(this.context);
    try {
      this.isAllowedToMoveFile = await this._checkUserPermissionToMoveFile();
      this.currUser = await this.sp.web.currentUser();
      this.spPortal = getSPByPath("https://epstin100.sharepoint.com/sites/EpsteinPortal", this.context);
      // Favorites list
      const allListItemsFavorites = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items()

      const { Id, Email } = this.currUser
      const userFound = allListItemsFavorites.find(user => user?.email.trim().toLocaleLowerCase() === this.currUser.Email.trim().toLocaleLowerCase())
      if (allListItemsFavorites && userFound) {
        // user exists in the list
        this.favorites = JSON.parse(userFound.favorites)
      } else {
        // user do not exist in the list
        const userPortal = await this.spPortal.web.siteUsers.getByEmail(Email)()
        await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.add({
          Title: Email,
          userId: userPortal.Id,
          email: Email,
          favorites: JSON.stringify([])
        })
      }

    } catch (error) {
      console.error('onInit error:', error)
    }

    this.dialogContainer = document.body.appendChild(
      document.createElement("div")
    );

    // Initialize the toast container when the extension is loaded
    const container = document.createElement('div');
    document.body.appendChild(container);

    // Render the Toaster (this will allow toasts to show)
    ReactDom.render(React.createElement(Toaster, { position: "top-left", }), container);

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

    // MeetingInv
    const meetingInvCompareOneCommand: Command = this.tryGetCommand('MeetingInv')
    meetingInvCompareOneCommand.visible = false

    // MeetingInv
    const draftCompareOneCommand: Command = this.tryGetCommand('draft')
    draftCompareOneCommand.visible = false

    // shoppingCart
    const shoppingCartCompareOneCommand: Command = this.tryGetCommand('shoppingCart')
    shoppingCartCompareOneCommand.visible = false

    // addToFavorites
    const addToFavoritesCompareOneCommand: Command = this.tryGetCommand('addToFavorites')
    addToFavoritesCompareOneCommand.visible = false

    // deleteFromFavorites
    const deleteFromFavoritesCompareOneCommand: Command = this.tryGetCommand('deleteFromFavorites')
    deleteFromFavoritesCompareOneCommand.visible = false


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

      case "linkToCategory":
        this._renderLinkToCategoryModal(selectedFiles);
        break;

      case "convertToPDF":
        ConvertToPdf(this.context, selectedFiles[0]);
        Swal.fire({
          title: "בקשתך נקלטה בהצלחה",
          text: "המרת הקובץ תחל בשניות הקרובות",
          icon: "success",
          allowOutsideClick: false,
          didOpen: () => {
            Swal.showLoading(); // Show loading spinner

          },
        });
        break;
      case "ExportToZip":
        this._renderExportZipModal(selectedFiles);
        break;
      case "External_Sharing":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsToShareDocuments(Array.from(event.selectedRows));
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
      case "MeetingInv":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsToMeetingInv(Array.from(event.selectedRows));
        }
        break;
      case "draft":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsToDraft(Array.from(event.selectedRows));
        }
        break;
      case "shoppingCart":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsToShoppingCart(Array.from(event.selectedRows));
        }
        break;
      case "addToFavorites":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsAddToFavorites(Array.from(event.selectedRows));
        }
        break;
      case "deleteFromFavorites":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsDeleteFromFavorites(Array.from(event.selectedRows));
        }
        break;
      default:
        throw new Error("Unknown command");
    }
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

  private _renderLinkToCategoryModal(selectedRows: any[]) {
    const element: React.ReactElement<MoveFileProps> = React.createElement(
      LinkToCategory,
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

  private _renderExportZipModal(selectedFiles: SelectedFile[]) {
    const element: React.ReactElement<ModalExtProps> =
      React.createElement(ExportZipModal, {
        sp: this.sp,
        context: this.context,
        selectedItems: selectedFiles,
        unMountDialog: this._closeDialogContainer,
        status: true
      });

    ReactDom.render(element, this.dialogContainer);
  }

  private async selectedRowsToShareDocuments(selectedRows: any[]): Promise<void> {
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

    const element: React.ReactElement<ISendEMailDialogContentProps> = React.createElement(
      SendEMailDialogContent,
      {
        close: this._closeDialogContainer,
        eMailProperties: new EMailProperties({
          To: "",
          Cc: "",
          Subject: `שיתוף מסמך - ${SendDocumentService.fileNames}`,
          Body: "",
        }),
        sendDocumentService: SendDocumentService,
        submit: () => {
          // Clear eMailProperties values
          new EMailProperties({
            To: "",
            Cc: "",
            Subject: "",
            Body: "",
          });
          // Close the dialog container
          this._closeDialogContainer();
        },
      }
    );

    ReactDom.render(element, this.dialogContainer);
  }

  private async selectedRowsToMeetingInv(selectedRows: any[]): Promise<void> {

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

    // Update CreateEvent properties
    CreateEvent.EmailAddress = emails;
    CreateEvent.fileNames = fileNames;
    CreateEvent.fileUris = fileRefs;
    CreateEvent.DocumentIdUrls = documentIdUrls;
    CreateEvent.webUri = this.context.pageContext.web.absoluteUrl;
    CreateEvent.context = this.context;

    // Set MS Graph client factory
    if (this.context && this.context.msGraphClientFactory) {
      CreateEvent.msGraphClientFactory = this.context.msGraphClientFactory;
    } else {
      console.error("MSGraphClientFactory is undefined.");
      return;
    }

    // Set server relative URL
    const currentRelativeUrl = this.context.pageContext.site.serverRelativeUrl;
    CreateEvent.ServerRelativeUrl = currentRelativeUrl;
    const element: React.ReactElement<IMeetingInvProps> = React.createElement(
      MeetingInv,
      {
        close: this._closeDialogContainer,
        eventProperties: new EventProperties({
          To: "",
          optionals: "",
          Subject: `זימון פגישה - ${CreateEvent.fileNames}`,
          Date: "",
          startTime: "",
          endTime: "",
          onlineMeeting: false,
          Body: "",
        }),
        createEvent: CreateEvent,
        submit: () => {
          // Clear eMailProperties values
          new EventProperties({
            To: "",
            optionals: "",
            Subject: "",
            Date: null,
            startTime: "",
            endTime: "",
            onlineMeeting: false,
            Body: "",
          });
          // Close the dialog container
          this._closeDialogContainer();
        },
      },
    )

    ReactDom.render(element, this.dialogContainer)
  }

  private async selectedRowsToDraft(selectedRows: any[]): Promise<void> {

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

    // Update CreateDraft properties
    CreateDraft.EmailAddress = emails;
    CreateDraft.fileNames = fileNames;
    CreateDraft.fileUris = fileRefs;
    CreateDraft.DocumentIdUrls = documentIdUrls;
    CreateDraft.webUri = this.context.pageContext.web.absoluteUrl;
    CreateDraft.context = this.context;

    // Set MS Graph client factory
    if (this.context && this.context.msGraphClientFactory) {
      CreateDraft.msGraphClientFactory = this.context.msGraphClientFactory;
    } else {
      console.error("MSGraphClientFactory is undefined.");
      return;
    }

    // Set server relative URL
    const currentRelativeUrl = this.context.pageContext.site.serverRelativeUrl;
    CreateDraft.ServerRelativeUrl = currentRelativeUrl;
    const element: React.ReactElement<IDraftProps> = React.createElement(
      Draft,
      {
        close: this._closeDialogContainer,
        draftProperties: new DraftProperties({
          Subject: `טיוטה - ${CreateDraft.fileNames}`,
          Body: "",
        }),
        createDraft: CreateDraft,
        submit: () => {
          // Clear eMailProperties values
          new DraftProperties({
            Subject: "",
            Body: "",
          });
          // Close the dialog container
          this._closeDialogContainer();
        },
      },
    )

    ReactDom.render(element, this.dialogContainer)
  }

  private async selectedRowsToShoppingCart(selectedRows: any[]): Promise<void> {

    for (const row of selectedRows) {
      try {
        const fileName = row.getValueByName("FileLeafRef").toString();
        const fileRef = row.getValueByName("FileRef").toString();
        const fileId = row.getValueByName("ID").toString();

        // Check if the item already exists in the shopping cart
        const existingItems = await this.spPortal.web.lists.getById("a9b46017-b0f0-4729-aeb0-a139aa421bc5")
          .items.filter(`RelativePath eq '${fileRef}'`)();

        if (existingItems.length > 0) {
          // Item already exists
          toast.error(`הקובץ ${fileName} כבר קיים בסל`);
        } else {
          // Add the item to the shopping cart
          await this.spPortal.web.lists.getById("a9b46017-b0f0-4729-aeb0-a139aa421bc5").items.add({
            Title: fileName,
            RelativePath: fileRef,
            itemID: fileId,
          });

          // Show success toast notification
          toast.success(`הקובץ ${fileName} נוסף לסל בהצלחה`);
        }

      } catch (error) {
        // Show error toast notification
        toast.error(`לא ניתן להוסיף את הקובץ לסל. אנא נסה שוב`);
      }
    }
  }

  private async selectedRowsAddToFavorites(selectedRows: any[]): Promise<void> {
    console.log("selectedRowsAddToFavorites - selectedRows:", selectedRows)

    // Initialize arrays to store file information
    const fileNames: string[] = [];
    const fileRefs: string[] = [];
    const documentIdUrls: string[] = [];
    const itemIds: string[] = [];
    const FSObjTypes: string[] = [];

    // Iterate through selected rows to gather file information
    selectedRows.forEach(row => {
      const fileName = row.getValueByName("FileLeafRef").toString();
      const fileRef = row.getValueByName("FileRef").toString();
      const documentIdUrl = row.getValueByName("ServerRedirectedEmbedUrl").toString();
      const itemId = row.getValueByName("ID").toString();
      const FSObjType = row.getValueByName("FSObjType").toString();

      fileNames.push(fileName);
      fileRefs.push(fileRef);
      documentIdUrls.push(documentIdUrl);
      itemIds.push(itemId)
      FSObjTypes.push(FSObjType)
    });
    const navTreeListIds = '41d92fdd-1469-475b-8d19-9fe47cca24be'
    const siteId = this.context.pageContext.site.id
    let linkTitle = null

    try {
      linkTitle = await this.spPortal.web.lists.getById(navTreeListIds).items.filter(`SiteId eq '${siteId}'`)()
    } catch (error) {
      console.error(error)
      return
    }

    const path = linkTitle ? `${linkTitle[0].Title}/${this.context.pageContext.list.title}` : ''

    const recoveryList = this.favorites

    if (!this.favorites.find(fav => fav.fileName === fileNames[0])) {
      try {
        const payload = {
          fileName: fileNames[0],
          fileRef: fileRefs[0],
          documentIdUrl: documentIdUrls[0],
          serverRelativeUrl: this.context.pageContext.site.serverRelativeUrl,
          absoluteUrl: this.context.pageContext.site.absoluteUrl,
          itemId: itemIds[0],
          libraryId: this.context.pageContext.list.id["_guid"],
          FSObjType: FSObjTypes[0],
          path: path
        }

        this.favorites.push(payload)
        const item = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.filter(`email eq '${this.currUser.Email}'`)()

        await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.getById(item[0].Id).update({
          favorites: JSON.stringify(this.favorites)
        }).then(() => {
          toast.success(`הקובץ ${fileNames[0]} נוסף למועדפים בהצלחה!`);
        })

      } catch (error) {
        this.favorites = recoveryList
        console.error(error)
        toast.error(`הוספת הקובץ ${fileNames[0]} נכשלה.`)
      }

    } else {
      toast.success(`הקובץ ${fileNames[0]} קיים במועדפים`)
    }
  }

  private async selectedRowsDeleteFromFavorites(selectedRows: any[]): Promise<void> {
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

    const recoveryList = this.favorites

    if (this.favorites.find(fav => fav.fileName === fileNames[0])) {
      try {

        this.favorites = this.favorites.filter(fav => fav.fileName !== fileNames[0])

        const item = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.filter(`email eq '${this.currUser.Email}'`)()

        await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.getById(item[0].Id).update({
          favorites: JSON.stringify(this.favorites)
        }).then(() => {
          toast.success(`הקובץ ${fileNames[0]} הוסר מהמועדפים בהצלחה!`);
        })

      } catch (error) {
        this.favorites = recoveryList
        console.error(error)
        toast.error(`הסרת הקובץ ${fileNames[0]} נכשלה.`)
      }

    } else {
      toast.success(`הקובץ לא ${fileNames[0]} קיים במועדפים`)
    }
  }

  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    Log.info(LOG_SOURCE, "List view state changed");

    let LibraryName = this.context.pageContext.list.title;

    const compareOneCommand: Command = this.tryGetCommand("Approval_Document");
    const compareThreeCommand: Command = this.tryGetCommand("linkToCategory");
    // const compareFourCommand: Command = this.tryGetCommand("RenameFile");
    const compareFiveCommand: Command = this.tryGetCommand("convertToPDF");
    const compareSixCommand: Command = this.tryGetCommand("ExportToZip");
    // const externalSharingCompareOneCommand: Command = this.tryGetCommand("External_Sharing");
    const meetingInvCompareOneCommand: Command = this.tryGetCommand('MeetingInv')
    const draftCompareOneCommand: Command = this.tryGetCommand('draft')
    const shoppingCartCompareOneCommand: Command = this.tryGetCommand('shoppingCart')
    const addToFavoritesCompareOneCommand: Command = this.tryGetCommand('addToFavorites')
    const deleteFromFavoritesCompareOneCommand: Command = this.tryGetCommand('deleteFromFavorites')

    if (compareOneCommand) {
      compareOneCommand.visible = event.selectedRows?.length === 1 && event.selectedRows[0]?.getValueByName('FSObjType') == 0
    }

    if (compareThreeCommand) {
      compareThreeCommand.visible = event.selectedRows?.length === 1;
      // if there is only one selected item and its a file and its a file type that can be converted to pdf
      if (compareFiveCommand) {
        compareFiveCommand.visible = event.selectedRows?.length === 1
          && event.selectedRows[0]?.getValueByName('FSObjType') == 0
          && this.typeSet.has(event.selectedRows[0]?.getValueByName(".fileType"));// &&
        // event.selectedRows[0].getValueByName("fileSize") > 0;
      }

      // // if there is one selected item or more and its a file
      // if (externalSharingCompareOneCommand) {
      //   if (event.selectedRows?.length > 0) {
      //     const fileExt = event.selectedRows[0].getValueByName(".fileType")
      //     if (fileExt.toLowerCase() !== "") externalSharingCompareOneCommand.visible = true;
      //   } else externalSharingCompareOneCommand.visible = false;
      // }

      // MeetingInv
      if (meetingInvCompareOneCommand) {
        if (event.selectedRows?.length > 0) {
          const fileExt = event.selectedRows[0].getValueByName(".fileType")
          if (fileExt.toLowerCase() !== "") meetingInvCompareOneCommand.visible = true;
        } else meetingInvCompareOneCommand.visible = false;
      }

      // Draft
      if (draftCompareOneCommand) {
        if (event.selectedRows?.length > 0) {
          const fileExt = event.selectedRows[0].getValueByName(".fileType")
          if (fileExt.toLowerCase() !== "") draftCompareOneCommand.visible = true;
        } else draftCompareOneCommand.visible = false;
      }

      // shoppingCart
      if (shoppingCartCompareOneCommand) {
        if (event.selectedRows?.length > 0) {
          const fileExt = event.selectedRows[0].getValueByName(".fileType")
          if (fileExt.toLowerCase() !== "") shoppingCartCompareOneCommand.visible = true;
        } else shoppingCartCompareOneCommand.visible = false;
      }

      // addToFavorites
      if (addToFavoritesCompareOneCommand) {

        if (event.selectedRows?.length === 1 && !this.favorites.find(fav => fav.fileName === event.selectedRows[0].getValueByName('FileLeafRef'))) {
          addToFavoritesCompareOneCommand.visible = true
        } else {
          addToFavoritesCompareOneCommand.visible = false
        }
      }

      // deleteFromFavorites
      if (deleteFromFavoritesCompareOneCommand) {
        if (event.selectedRows?.length === 1 && this.favorites.find(fav => fav.fileName === event.selectedRows[0].getValueByName('FileLeafRef'))) {
          deleteFromFavoritesCompareOneCommand.visible = true
        } else {
          deleteFromFavoritesCompareOneCommand.visible = false
        }
      }

      if (compareSixCommand) {
        compareSixCommand.visible = event.selectedRows?.length >= 1;
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
}
