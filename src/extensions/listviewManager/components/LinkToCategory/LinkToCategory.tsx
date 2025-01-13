import React, { useEffect, useState } from "react";
import styles from "./LinkToCCategory.module.scss";
import { SPFI } from "@pnp/sp";
import { SelectedFile } from "../../models/global.model";
import modalStyles from "../../styles/modalStyles.module.scss";
import { FolderPicker, IFolder } from "@pnp/spfx-controls-react/lib/FolderPicker";
import { Button } from "@mui/material";
import { decimalToBinaryArray } from "../../service/util.service";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";

export interface LinkToCategoryProps {
  sp: SPFI;
  context: ListViewCommandSetContext;
  selectedRows: SelectedFile[]; // Accept an array of selected files
  unMountDialog: () => void;
}

export default function LinkToCategory({
  sp,
  context,
  selectedRows,
  unMountDialog,
}: LinkToCategoryProps) {
  const [errorMsg, setErrorMsg] = useState<string>("");
  const [unauthorizedFiles, setUnauthorizedFiles] = useState<SelectedFile[]>([]);

  useEffect(() => {
    initMoveFile();
  }, []);

  const initMoveFile = async () => {
    try {
      const unauthorized = await checkFilesPermission(
        context.pageContext.list.id["_guid"],
        selectedRows
      );

      if (unauthorized.length > 0) {
        setUnauthorizedFiles(unauthorized);
        const fileNames = unauthorized.map(file => file.FileLeafRef).join(", ");
        setErrorMsg(
          `You do not have permission to move the following files: ${fileNames}`
        );
        return;
      }

      // open folder picker automatically
      const attachForm = document.getElementById("FilePickerWrapper");
      (
        attachForm.children[0].children[1].children[1].children[0] as any
      ).click();
    } catch (err) {
      console.error("initMoveFile", err);
    }
  };

  const checkFilesPermission = async (
    listId: string,
    files: SelectedFile[]
  ): Promise<SelectedFile[]> => {
    try {
      const list = sp.web.lists.getById(listId);
      const unauthorizedFiles: SelectedFile[] = [];

      for (const file of files) {
        const itemPerm = await list.items
          .getById(file.ID)
          .getCurrentUserEffectivePermissions();
        const binaryArray = decimalToBinaryArray(Number(itemPerm.Low));

        if (binaryArray[binaryArray.length - 3] !== 1) {
          unauthorizedFiles.push(file);
        }
      }

      return unauthorizedFiles;
    } catch (err) {
      console.error("checkFilesPermission", err);
      return files; // Assume all are unauthorized in case of error
    }
  };

  const onFolderSelect = async (folder: IFolder) => {
    try {
      const libPath = context.pageContext.list.serverRelativeUrl;
      const item = selectedRows[0];
      // Define the target library or folder and link details
      const targetLibraryUrl = folder.ServerRelativeUrl; // Server relative path to your library
      const fileName = `${item.FileLeafRef}.url`; // Name of the link file
      // const linkUrl = "https://epstin100.sharepoint.com/sites/JerusalemDistrict/Shared%20Documents/%D7%97%D7%95%D7%91%D7%A8%D7%AA122.xlsx"; // URL of the document
      const linkUrl = `https://epstin100.sharepoint.com${item.FileRef}`
      const fileContent = `[InternetShortcut]\nURL=${linkUrl}`;

      // Convert file content to a Blob
      const fileBlob = new Blob([fileContent], { type: "text/plain" });

      // Upload the file
      const result = await sp.web.getFolderByServerRelativePath(targetLibraryUrl).files.addUsingPath(fileName, fileBlob, { Overwrite: true });
    } catch (error) {
      console.error("Error creating link file:", error);
    }
  };

  const serverRelativeUrl = context.pageContext.site.serverRelativeUrl; // to allow to use in multiple sites
  return (
    <div>
      <div
        className={`${modalStyles.modalScreen}`}
        style={errorMsg ? {} : { backgroundColor: "unset" }}
        onClick={() => unMountDialog()}
      >
        {errorMsg && (
          <div
            className={`${modalStyles.modal} ${styles.moveFileModal}`}
            onClick={(ev: any) => ev.stopPropagation()}
          >
            <span style={{ color: "red" }}>{errorMsg}</span>

            <Button onClick={unMountDialog}>Confirm</Button>
          </div>
        )}

        {!errorMsg && (
          <div
            className={`${modalStyles.modal} ${styles.hiddenMoveFileModal}`}
            onClick={(ev: any) => ev.stopPropagation()}
          >
            <div id="FilePickerWrapper">
              <FolderPicker
                context={context as any}
                label="Folder Picker"
                required={true}
                key={"folderPicker"}
                onSelect={onFolderSelect}
                canCreateFolders={false}
                defaultFolder={{
                  Name: "Libraries",
                  ServerRelativeUrl: context.pageContext.list.serverRelativeUrl,
                }}
                rootFolder={{
                  Name: "Libraries",
                  ServerRelativeUrl: serverRelativeUrl,
                }}
              />
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
