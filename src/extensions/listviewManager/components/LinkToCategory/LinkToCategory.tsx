import React, { useEffect, useState } from "react";
import styles from "./LinkToCCategory.module.scss";
import { SPFI } from "@pnp/sp";
import { SelectedFile } from "../../models/global.model";
import modalStyles from "../../styles/modalStyles.module.scss";
import { FolderPicker, IFolder } from "@pnp/spfx-controls-react/lib/FolderPicker";
import { Button } from "@mui/material";
import { decimalToBinaryArray } from "../../service/util.service";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { FolderExplorer } from "@pnp/spfx-controls-react";
import { getSiteCollections, createSitesOptions, createSitesMap } from "../../service/link.service";
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import { Site } from "../../models/LinkToCategory";
import Autocomplete from '@mui/material/Autocomplete';
import TextField from '@mui/material/TextField';
import { SPFx } from "@pnp/sp";
import { CacheProvider } from "@emotion/react";
import { cacheRtl } from "../../models/cacheRtl";
import { IconButton } from "@mui/material";
import CloseIcon from '@mui/icons-material/Close';
import CheckIcon from '@mui/icons-material/Check';
import Swal from "sweetalert2";



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
  // const [siteMap, setMap] = useState<Map<string, string>>(new Map());

  const [selectedSite, setSelectedSite] = useState<Site>(null);
  const [siteOptions, setOptions] = useState<Site[]>([]);
  const [selectedFolder, setSelectedFolder] = useState<IFolder>(null);
  const [isLoading, setIsLoading] = useState<Boolean>(true);
  // const [siteOptions, setOptions] = useState<string[]>([]);

  // const [selectedSite, setSite] = useState<string>("");
  // const [selectedSiteUrl, setSelectedSiteUrl] = useState<string>(context.pageContext.site.serverRelativeUrl);

  useEffect(() => {
    initMoveFile();
  }, []);

  const initMoveFile = async () => {
    try {
      const unauthorized = await checkFilesPermission(
        context.pageContext.list.id["_guid"],
        selectedRows
      );
      const sites = await getSiteCollections(context);
      // setMap(createSitesMap(sites));
      setOptions(sites);
      const serverRelativeUrl = context.pageContext.site.serverRelativeUrl;
      const defaultSite: Site = sites.filter(site => site.Url === serverRelativeUrl)[0];
      console.log("default site", defaultSite);

      setSelectedSite(defaultSite);
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
      console.log('attachForm', attachForm)
      // (
      //   attachForm.children[1].children[1].children[1].children[0] as any
      // ).click();
      setIsLoading(false)
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

      const sp: SPFI = new SPFI(`https://epstin100.sharepoint.com${selectedSite.Url}`).using(SPFx(context));
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
      Swal.fire({
        title: "הקישור נוצר בהצלחה",
        icon: "success",
      });
      unMountDialog();
    } catch (error) {
      console.error("Error creating link file:", error);
      // **Check if it's a 403 Forbidden Error**
      if (error?.isHttpRequestError && error?.status === 403) {
        Swal.fire({
          title: "אין לך הרשאות להעלות קובץ לתיקייה זו",
          text: "אנא בדוק את ההרשאות שלך או פנה למנהל המערכת.",
          icon: "error",
        });
      } else {
        Swal.fire({
          title: "שגיאה ביצירת הקובץ",
          text: "אירעה שגיאה בלתי צפויה. אנא פנה למנהל המערכת",
          icon: "error",
        });
      }
    }
  };

  const handleSiteChange = (newSite: string) => {
    const selected_site: Site = siteOptions.filter(site => site.Title === newSite)[0];
    console.log(`selected site ${selected_site}`);

    setSelectedSite(selected_site);
  }

  const serverRelativeUrl = context.pageContext.site.serverRelativeUrl; // to allow to use in multiple sites
  if (isLoading) return null;
  return (
    <>
      <CacheProvider value={cacheRtl}>
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

          <div
            className={`${modalStyles.modal} ${modalStyles.folderExplorerModal}`}
            onClick={(ev: any) => ev.stopPropagation()}
          >
            <div id="FilePickerWrapper">
              <Autocomplete
                options={siteOptions.map(option => option.Title)}
                defaultValue={selectedSite.Title}
                sx={{ width: 300 }}
                renderInput={(params) => <TextField {...params} label="אתר" />}
                onChange={(event, newSite) => handleSiteChange(newSite)}
              />
              <FolderExplorer
                context={context as any}
                key={"folderexplorer" + selectedSite?.Title || ''}
                onSelect={(folder) => setSelectedFolder(folder)}
                canCreateFolders={false}
                defaultFolder={{
                  Name: "ספריות",
                  ServerRelativeUrl: selectedSite?.Url || context.pageContext.list.serverRelativeUrl,
                }}
                rootFolder={{
                  Name: "ספריות",
                  ServerRelativeUrl: selectedSite?.Url || context.pageContext.site.serverRelativeUrl,
                }}
                siteAbsoluteUrl={`https://epstin100.sharepoint.com${selectedSite?.Url || context.pageContext.site.serverRelativeUrl}`}
              />
              <div className={`${modalStyles.modalFooter} ${modalStyles.modalFooterEnd}`}>
                <Button
                  color="error"
                  className={styles.button}
                  onClick={unMountDialog}

                  startIcon={<IconButton disableRipple style={{ color: "#f58383", paddingLeft: 0, margin: "0px !important" }}><CloseIcon /></IconButton>}
                  sx={{
                    "& .MuiButton-startIcon": {
                      margin: 0,
                    },
                  }}
                >סגירה</Button>
                <Button
                  disabled={selectedFolder === null}
                  onClick={() => onFolderSelect(selectedFolder)}
                  className={styles.button}
                  endIcon={<IconButton disableRipple style={{ color: "#1976d2", margin: "0px" }}><CheckIcon /></IconButton>}
                >שמירה</Button>
              </div>
            </div>
          </div>
          {/* )} */}
        </div>
      </CacheProvider>
    </>
  );
}
