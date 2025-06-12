import React from "react";
import ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import { getSP, getSPByPath } from "../../pnpjs-config";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters
} from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { EMailProperties, EventProperties, DraftProperties, SelectedFile } from "./models/global.model";
import ApproveDocument, {
  ApproveDocumentProps,
} from "./components/ApproveDocument/ApproveDocument.cmp";
import ModalExt from "../../extensions/listviewManager/components/FolderHierarchy/ModalExt";
import { MoveFileProps } from "./components/MoveFile/MoveFile.cmp";
import { PermissionKind } from "@pnp/sp/security";
import { ModalExtProps } from "./components/FolderHierarchy/ModalExtProps";
import { MergePDFProps } from "./components/MergePDF/MergePDF.cmp";
import LinkToCategory from "./components/LinkToCategory/LinkToCategory";
import { ConvertToPdf, getConvertibleTypes } from "./service/pdf.service";
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
import toast, { Toaster } from 'react-hot-toast';
import './../ext.css'
import MergePDF from "./components/MergePDF/MergePDF.cmp";
import { Version } from '@microsoft/sp-core-library';
import { clickEvent } from './service/thirdPartyOpen.service'
const { solution } = require("../../../config/package-solution.json");

export interface IListviewManagerCommandSetProperties {
  sampleTextOne: string;
}

const LOG_SOURCE: string = "ListviewManagerCommandSet";
const FAVORITES_LIST_ID = '6f3d6257-4a9b-41fe-a847-487c942cd628'
const FAVORITES_ADDIN_LIST_ID = 'eccc2588-4c91-4259-bd18-f2d7c780803d';
const CONVERTIBLE_TYPES_ID = 'b748b7b9-6b49-44f9-b889-ef7e99ebdb47'
const CONVERTIBLE_TYPES = ['doc', 'docx', 'dot', 'dotx', 'docm', 'dotm', 'rtf', 'txt', 'odt', 'xls', 'xlsx', 'xlsm', 'xltx', 'xltm', 'xml', 'xml', 'ods', 'ppt', 'pptx', 'ppsx', 'potx', 'pps', 'pot', 'odp', 'pdf']

export default class ListviewManagerCommandSet extends BaseListViewCommandSet<IListviewManagerCommandSetProperties> {
  private dialogContainer: HTMLDivElement;
  private sp: SPFI;
  private currUser: ISiteUserInfo;
  private typeSet: Set<string> = new Set(CONVERTIBLE_TYPES);
  private typeToConvert: Set<string> = new Set(CONVERTIBLE_TYPES);

  private allowedUsers: string[] = [
    "EpsteinSystem@Epstein.co.il",
  ].map((e) => e.toLocaleLowerCase());

  private favorites: any[] = []
  private favoritesAddin: any[] = []
  private spPortal: SPFI = null

  public async onInit(): Promise<void> {
    console.log(solution.name + ":", solution.version);

    Log.info(LOG_SOURCE, "Initialized ListviewManagerCommandSet");
    this.sp = getSP(this.context);
    this.handleCmds();
    this.initExt();

    return Promise.resolve();
  }

  private handleCmds() {
    const commandConfigs = this.context.manifest.items;

    const cmdArr: any[] = Object.entries(commandConfigs).map(([id, config]) => ({
      ...config,
      id,
    }));

    cmdArr.forEach((cmd) => {
      const cmdObj: Command = this.tryGetCommand(cmd.id);

      if (cmdObj) {
        cmdObj.visible = cmd.visible || false;
      }
    });
  }

  private async initExt() {
    const t0 = performance.now()
    clickEvent(this.context);
    try {
      this.currUser = await this.sp.web.currentUser();
      this.spPortal = getSPByPath("https://epstin100.sharepoint.com/sites/EpsteinPortal", this.context);
      // Favorites list
      const allListItemsFavorites = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items()
      // Convertible Types list
      this.typeToConvert = new Set((await this.spPortal.web.lists.getById(CONVERTIBLE_TYPES_ID).items.select("Title")()).map(item => item.Title));

      const { Email } = this.currUser
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

      const allListItemsFavoritesAddin = await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items()

      const userFoundAddin = allListItemsFavoritesAddin.find(
        user =>
          String(user?.Title?.trim() + '@Epstein.co.il').toLowerCase() ===
          this.currUser.Email?.trim().toLowerCase()
      );

      if (allListItemsFavoritesAddin && userFoundAddin) {
        // user exists in the list
        this.favoritesAddin = JSON.parse(userFoundAddin.Items)
      } else {
        // user do not exist in the list                
        await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items.add({
          Title: Email?.split('@')[0] || 'User',
          Items: JSON.stringify([])
        })
      }
      this.typeSet = await getConvertibleTypes(this.context);
      this._checkUserPermissionToMoveFile()
    } catch (error) {
      console.error('onInit error:', error)
    }

    // Listen for updates from the modal
    window.addEventListener('favoritesUpdated', this.refreshFavorites.bind(this));

    this.dialogContainer = document.body.appendChild(
      document.createElement("div")
    );

    // Initialize the toast container when the extension is loaded
    const container = document.createElement('div');
    document.body.appendChild(container);

    // Render the Toaster (this will allow toasts to show)
    ReactDom.render(React.createElement(Toaster, { position: "top-left", }), container);

    const isUserAllowed = this.allowedUsers.includes(this.currUser.Email);
    if (!isUserAllowed) {
      require("./styles/createNewFolder.module.scss"); // hide the button create new folder if the user is not allowed
    }

    console.log("initExt took " + (performance.now() - t0) + " milliseconds.");
  }

  protected get dataVersion(): Version {
    return Version.parse(solution.version);
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {

    const fullUrl = window.location.href;
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
      case "mergeToPDF":
        // Check if the user selected some items
        if (event.selectedRows.length > 0) {
          // Process the selected rows and retrieve contacts
          await this.selectedRowsToMergePDF(Array.from(event.selectedRows));
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

  private async _checkUserPermissionToMoveFile(): Promise<boolean> {
    try {
      return await this.sp.web.lists
        .getById(this.context.pageContext.list.id["_guid"])
        .currentUserHasPermissions(PermissionKind.EditListItems);
    } catch (error) {
      console.error("Error while checking user permission to move file", error);
    }
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

    // Update SendDocumentService properties
    SendDocumentService.EmailAddress = [] //emails;
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

    // Update CreateEvent properties
    CreateEvent.EmailAddress = [] //emails;
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
        sp: this.sp,
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

    // Update CreateDraft properties
    CreateDraft.EmailAddress = [] //emails;
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
        sp: this.sp,
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

  private async getFavorites(): Promise<any[]> {
    const allListItemsFavorites = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items()

    const { Id, Email } = this.currUser
    const userFound = allListItemsFavorites.find(user => user?.email.trim().toLocaleLowerCase() === Email.trim().toLocaleLowerCase())
    if (allListItemsFavorites && userFound) {
      // user exists in the list
      return JSON.parse(userFound.favorites)
    } else {
      return []
    }
  }

  private async refreshFavorites(): Promise<void> {
    console.log('favoritesUpdated event received');
    try {
      const allListItemsFavorites = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items()

      const { Id, Email } = this.currUser
      const userFound = allListItemsFavorites.find(user => user?.email.trim().toLocaleLowerCase() === Email.trim().toLocaleLowerCase())
      if (allListItemsFavorites && userFound) {
        // user exists in the list
        this.favorites = JSON.parse(userFound.favorites)
        console.log('favoritesUpdated updated');
      } else {
        this.favorites = []
      }
      this.favoritesAddin = await this.getFavoritesAddin()
    } catch (error) {
      console.log('Error in refreshFavorites: ', error)
    }
  }

  private async getFavoritesAddin(): Promise<any[]> {
    const allListItemsFavoritesAddin = await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items()

    const userFound = allListItemsFavoritesAddin.find(
      user =>
        String(user?.Title?.trim() + '@Epstein.co.il').toLowerCase() ===
        this.currUser.Email?.trim().toLowerCase()
    );

    if (allListItemsFavoritesAddin && userFound) {
      // user exists in the list
      return JSON.parse(userFound.Items)
    } else {
      return []
    }
  }

  private async selectedRowsAddToFavorites(selectedRows: any[]): Promise<void> {

    // Initialize arrays to store file information
    const fileNames: string[] = [];
    const fileRefs: string[] = [];
    const documentIdUrls: string[] = [];
    const itemIds: string[] = [];
    const FSObjTypes: string[] = [];
    const Projects: string[] = [];

    // Iterate through selected rows to gather file information
    selectedRows.forEach(row => {
      const fileName = row.getValueByName("FileLeafRef").toString();
      const fileRef = row.getValueByName("FileRef").toString();
      const documentIdUrl = row.getValueByName("ServerRedirectedEmbedUrl").toString();
      const itemId = row.getValueByName("ID").toString();
      const FSObjType = row.getValueByName("FSObjType").toString();
      const Project = row.getValueByName("Project");

      fileNames.push(fileName);
      fileRefs.push(fileRef);
      documentIdUrls.push(documentIdUrl);
      itemIds.push(itemId)
      FSObjTypes.push(FSObjType)
      Projects.push(Project)
    });
    const libraryRoot = this.context.pageContext.list.serverRelativeUrl;
    let level2 = ''
    if (libraryRoot.split('/').length + 1 === fileRefs[0].split('/').length) {
    } else {
      level2 = '/' + fileRefs[0].split('/')[4]
    }

    const navTreeListIds = '41d92fdd-1469-475b-8d19-9fe47cca24be'
    const siteId = this.context.pageContext.site.id
    let linkTitle = null

    try {
      console.log('this.spPortal:', this.spPortal)
      linkTitle = await this.spPortal.web.lists.getById(navTreeListIds).items.filter(`SiteId eq '${siteId}'`)()
    } catch (error) {
      console.error('Error linkTitle', error)
      return
    }

    try {
      this.favorites = await this.getFavorites()
    } catch (error) {
      console.error('Error getFavorites', error)
    }

    let isModified: boolean = false

    try {
      this.favoritesAddin = await this.getFavoritesAddin()
    } catch (error) {
      console.error('Error getFavoritesAddin', error)
    }

    const path = linkTitle ? `${linkTitle[0].Title}/${this.context.pageContext.list.title}${level2}` : ''

    const recoveryList = this.favorites

    for (let i = 0; i < selectedRows.length; i++) {
      if (!this.favorites.find(fav => fav.fileRef === fileRefs[i])) {

        if (selectedRows[i].getValueByName("FSObjType").toString() === '0') {
          const payload = {
            fileName: fileNames[i],
            fileRef: fileRefs[i],
            documentIdUrl: documentIdUrls[i],
            serverRelativeUrl: this.context.pageContext.site.serverRelativeUrl,
            absoluteUrl: this.context.pageContext.site.absoluteUrl,
            itemId: itemIds[i],
            libraryId: this.context.pageContext.list.id["_guid"],
            FSObjType: FSObjTypes[i],
            path: path,
            project: Projects[i]
          }
          this.favorites.push(payload)
        }

        if (selectedRows[i].getValueByName("FSObjType").toString() === '1'
          && !this.favoritesAddin.find(favAddin => favAddin.Path === selectedRows[i].getValueByName("FileRef").toString())) {
          console.log('isModified')
          isModified = true
          const payloadAddin = {
            Title: fileNames[i],
            Path: fileRefs[i],
            Type: 'Folder'
          }
          this.favoritesAddin.push(payloadAddin)
        }
      }
    }

    if (this.favoritesAddin.length > 0 && isModified) {
      try {
        const allListItemsFavoritesAddin = await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items()

        const userFound = allListItemsFavoritesAddin.find(
          user =>
            String(user?.Title?.trim() + '@Epstein.co.il').toLowerCase() ===
            this.currUser.Email?.trim().toLowerCase()
        );
        if (userFound) {
          await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items.getById(userFound.Id).update({
            Items: JSON.stringify(this.favoritesAddin)
          })
        } else {
          throw new Error('User not found in favoritesAddin list')
        }
      } catch (error) {
        console.error('Error in adding favoritesAddin', error)
      }
    }

    if (this.favorites.length > 0) {
      try {
        const item = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.filter(`email eq '${this.currUser.Email}'`)()

        if (item) {
          await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.getById(item[0].Id).update({
            favorites: JSON.stringify(this.favorites)
          }).then(() => {
            toast.success(`הקבצים נוספו למועדפים בהצלחה!`);
          })
        } else {
          throw new Error('User not found in favorites list')
        }

      } catch (error) {
        this.favorites = recoveryList
        console.error('Error in adding favorites', error)
        toast.error('הוספת הקבצים למועדפים נכשלה.')
      }
    } else {
      toast.success(`הקבצים קיימים במועדפים.`)
    }

  }

  private async selectedRowsDeleteFromFavorites(selectedRows: any[]): Promise<void> {
    // Initialize arrays to store file information
    const fileNames: string[] = [];
    const fileRefs: string[] = [];

    // Iterate through selected rows to gather file information
    selectedRows.forEach(row => {
      const fileName = row.getValueByName("FileLeafRef").toString();
      const fileRef = row.getValueByName("FileRef").toString();

      fileNames.push(fileName);
      fileRefs.push(fileRef);
    });

    try {
      this.favorites = await this.getFavorites()
    } catch (error) {
      console.error(error)
    }


    let isExist: boolean = false

    try {
      this.favoritesAddin = await this.getFavoritesAddin()
    } catch (error) {
      console.error('Error getFavoritesAddin', error)
    }

    const recoveryList = this.favorites

    for (let i = 0; i < selectedRows.length; i++) {
      if (this.favorites.find(fav => fav.fileRef === fileRefs[i])) {
        this.favorites = this.favorites.filter(fav => fav.fileRef !== fileRefs[i])
      }

      if (this.favoritesAddin.find(favAddin => favAddin.Path === fileRefs[i])) {
        isExist = true
        this.favoritesAddin = this.favoritesAddin.filter(favAddin => favAddin.Path !== fileRefs[i])
      }
    }

    if (isExist) {
      try {
        const allListItemsFavoritesAddin = await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items()

        const userFound = allListItemsFavoritesAddin.find(
          user =>
            String(user?.Title?.trim() + '@Epstein.co.il').toLowerCase() ===
            this.currUser.Email?.trim().toLowerCase()
        );

        await this.spPortal.web.lists.getById(FAVORITES_ADDIN_LIST_ID).items.getById(userFound.Id).update({
          Items: JSON.stringify(this.favoritesAddin)
        })
      } catch (error) {
        console.error('Error in deleting favoritesAddin', error)
      }
    }

    try {
      const item = await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.filter(`email eq '${this.currUser.Email}'`)()

      await this.spPortal.web.lists.getById(FAVORITES_LIST_ID).items.getById(item[0].Id).update({
        favorites: JSON.stringify(this.favorites)
      }).then(() => {
        toast.success('הקבצים הוסרו מהמועדפים בהצלחה!')
      })

    } catch (error) {
      this.favorites = recoveryList
      console.error(error)
      toast.error('הסרת הקבצים נכשלה.')
    }
  }

  private async selectedRowsToMergePDF(selectedRows: any[]): Promise<void> {
    const selectedFiles: { Title: string; RelativePath: string; SiteAdress: string }[] = [];

    for (const row of selectedRows) {
      try {
        const fileName = row.getValueByName("FileLeafRef").toString();
        const fileRef = row.getValueByName("FileRef").toString();
        const fileUrl = row.getValueByName("ServerRedirectedEmbedUrl").toString();

        // Step 1: Clean the URL to get the tenant and site information
        const cleanedUrl = fileUrl.split("/_layout")[0];

        // Create an object with file name, cleaned file reference, and cleaned URL
        const fileData = {
          Title: fileName,
          RelativePath: fileRef,
          SiteAdress: cleanedUrl
        };

        // Add the object to the selectedFiles array
        selectedFiles.push(fileData);

      } catch (error) {
        // Show error toast notification
        toast.error(`לא ניתן להוסיף את הקובץ לסל. אנא נסה שוב`);
      }
    }

    // Render the MergePDF component after processing all selected rows
    const element: React.ReactElement<MergePDFProps> = React.createElement(MergePDF, {
      context: this.context,
      selectedItems: selectedFiles, // Pass the array of file data objects
      unMountDialog: this._closeDialogContainer,
    });

    ReactDom.render(element, this.dialogContainer);
  }

  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    Log.info(LOG_SOURCE, "List view state changed");

    // Always on cmds: folderHierarchyCmd

    const approvalCmd: Command = this.tryGetCommand("Approval_Document");
    const convertToPdfCmd: Command = this.tryGetCommand("convertToPDF");
    const meetingInvCmd: Command = this.tryGetCommand('MeetingInv')
    const draftCmd: Command = this.tryGetCommand('draft')
    const shoppingCartCmd: Command = this.tryGetCommand('shoppingCart')
    const mergeToPDFCmd: Command = this.tryGetCommand('mergeToPDF')
    const addToFavoritesCmd: Command = this.tryGetCommand('addToFavorites')
    const deleteFromFavoritesCmd: Command = this.tryGetCommand('deleteFromFavorites');
    const linkToCategoryCmd: Command = this.tryGetCommand('linkToCategory');
    const exportToZipCmd: Command = this.tryGetCommand("ExportToZip");

    const selectedRows = event.selectedRows;

    if (approvalCmd) {
      approvalCmd.visible = selectedRows?.length === 1 && selectedRows[0]?.getValueByName('FSObjType') == 0
    }

    if (linkToCategoryCmd) {
      linkToCategoryCmd.visible = selectedRows?.length === 1;
    }
    // if there is only one selected item and its a file and its a file type that can be converted to pdf
    if (convertToPdfCmd) {
      convertToPdfCmd.visible = selectedRows?.length === 1
        && selectedRows[0]?.getValueByName('FSObjType') == 0
        && this.typeSet?.has(selectedRows[0]?.getValueByName(".fileType"));
    }

    // MeetingInv
    if (meetingInvCmd) {
      if (selectedRows?.length > 0) {
        const fileExt = selectedRows[0].getValueByName(".fileType")
        if (fileExt.toLowerCase() !== "" && selectedRows[0]?.getValueByName('FSObjType') == 0) meetingInvCmd.visible = true;
      } else meetingInvCmd.visible = false;
    }

    // Draft
    if (draftCmd) {
      if (selectedRows?.length > 0) {
        const fileExt = selectedRows[0].getValueByName(".fileType")
        if (fileExt.toLowerCase() !== "" && selectedRows[0]?.getValueByName('FSObjType') == 0) draftCmd.visible = true;
      } else draftCmd.visible = false;
    }

    // shoppingCart
    if (shoppingCartCmd) {
      if (selectedRows?.length > 0) {
        const fileExt = selectedRows[0].getValueByName(".fileType")
        if (fileExt.toLowerCase() !== "") shoppingCartCmd.visible = true;
      } else shoppingCartCmd.visible = false;
    }

    // addToFavorites
    if (addToFavoritesCmd) {
      if (selectedRows?.length > 0) {
        const allSelectedFiles = selectedRows.map(row => row.getValueByName('FileRef'));
        const allFavoritesFiles = [...this.favorites.map(fav => fav.fileRef), ...this.favoritesAddin.map(favAddin => favAddin.Path)];
        // Only show the button if every selected file is NOT in favorites.
        addToFavoritesCmd.visible = allSelectedFiles.every(file => !allFavoritesFiles.includes(file));
      } else {
        addToFavoritesCmd.visible = false;
      }
    }

    // deleteFromFavorites
    if (deleteFromFavoritesCmd) {
      if (selectedRows?.length > 0) {
        const allSelectedFiles = selectedRows.map(row => row.getValueByName('FileRef'));
        const allFavoritesFiles = [...this.favorites.map(fav => fav.fileRef), ...this.favoritesAddin.map(favAddin => favAddin.Path)];
        // Only show the button if every selected file IS in favorites.
        deleteFromFavoritesCmd.visible = allSelectedFiles.every(file => allFavoritesFiles.includes(file));
      } else {
        deleteFromFavoritesCmd.visible = false;
      }
    }

    // mergeToPDF
    if (mergeToPDFCmd) {
      if (selectedRows?.length > 0) {
        const fileExt = selectedRows[0].getValueByName(".fileType").toLowerCase();

        mergeToPDFCmd.visible = this.typeToConvert?.has(fileExt);
      } else {
        mergeToPDFCmd.visible = false;
      }
    }

    if (exportToZipCmd) {
      exportToZipCmd.visible = selectedRows?.length >= 1;
    }

    this.raiseOnChange();
  }
}