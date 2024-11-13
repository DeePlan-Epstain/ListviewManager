import React, { useEffect, useState } from "react";
import styles from "./MoveFile.module.scss";
import { SPFI } from "@pnp/sp";
import { SelectedFile } from "../../models/global.model";
import modalStyles from "../../styles/modalStyles.module.scss";
import { FolderPicker, IFolder } from "@pnp/spfx-controls-react/lib/FolderPicker";
import { Button } from "@mui/material";
import { decimalToBinaryArray } from "../../service/util.service";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";

export interface MoveFileProps {
  sp: SPFI;
  context: ListViewCommandSetContext;
  selectedRows: SelectedFile[]; // Accept an array of selected files
  unMountDialog: () => void;
}

export default function MoveFile({
  sp,
  context,
  selectedRows,
  unMountDialog,
}: MoveFileProps) {
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

  const onFolderSelect = async (folder: IFolder): Promise<void> => {
    try {
      const libPath = context.pageContext.list.serverRelativeUrl;

      for (const selectedRow of selectedRows) {
        const destUrl = folder.ServerRelativeUrl + "/" + selectedRow.FileLeafRef;

        if (destUrl === selectedRow.FileRef) {
          setErrorMsg("The selected folder is the same as the current folder.");
          return;
        }

        const destFolderProps: any = await sp.web
          .getFolderByServerRelativePath(folder.ServerRelativeUrl)
          .expand("properties")();
        const listId = destFolderProps.Properties["vti_x005f_listname"].slice(
          1,
          -1
        );

        const hasPermission = await checkFilesPermission(listId, selectedRows);

        if (hasPermission.length > 0) {
          const fileNames = hasPermission.map(file => file.FileLeafRef).join(", ");
          setErrorMsg(
            `You do not have permission to move the following files: ${fileNames}`
          );
          return;
        }

        await sp.web.lists
          .getById("e3aef93d-30b9-4f7d-ae41-6552d1346644")
          .items.add({
            Title: selectedRow.FileLeafRef,
            originalPath: selectedRow.FileRef.split("/").slice(3).join("/"),
            newPath: folder.ServerRelativeUrl.split("/").slice(3).join("/"),
            libraryId: context.pageContext.list.id["_guid"],
            FileId: selectedRow.ID,
          });
      }

      unMountDialog();
    } catch (err) {
      console.error("onFolderSelect", err);
      if (err.message.includes("Access denied"))
        setErrorMsg(
          "Your account does not have permission to move the files to the selected folder"
        );
      else if (err.message.includes("The destination file already exists"))
        setErrorMsg("The destination file already exists");
      else
        setErrorMsg(
          "Error occurred while moving the files, please contact administrator"
        );
    }
  };

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
                  ServerRelativeUrl: "/sites/METPRODdoc",
                }}
              />
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
