import React, { useEffect, useState } from "react";
import styles from "./TreeFolders.module.scss";
import modalStyles from "../../styles/modalStyles.module.scss";
import { IFolder } from "@pnp/spfx-controls-react/lib/FolderPicker";
import { Button } from "@mui/material";
import { FolderExplorer } from "@pnp/spfx-controls-react/lib/FolderExplorer";
import Autocomplete from '@mui/material/Autocomplete';
import TextField from '@mui/material/TextField';
import { IconButton } from "@mui/material";
import CloseIcon from '@mui/icons-material/Close';
import CheckIcon from '@mui/icons-material/Check';
import { CacheProvider } from "@emotion/react";
import { cacheRtl } from "../../models/cacheRtl";
import { getSPByPath } from "../../../../pnpjs-config";
import toast from "react-hot-toast";
import { exportToZip } from "../../service/zip.service";
import { ConvertToPdf } from "../../service/mergePdf.service";
import { Site } from "../../models/LinkToCategory";
import { getSiteCollections } from "../../service/link.service";


export interface TreeFoldersProps {
  context: any;
  fileToSave: any;
  isClose: () => void;
  fileName: string;
}

export default function TreeFolders({ context, isClose, fileToSave, fileName }: TreeFoldersProps) {
  const [fileNameState, setFileName] = useState<string>(fileName);
  const [fileExtension, setFileExtension] = useState<string>("");
  const [selectedSite, setSelectedSite] = useState<Site>({ Title: "", Url: "" });
  const [siteOptions, setOptions] = useState<Site[]>([]);
  const [selectedFolder, setSelectedFolder] = useState<IFolder>(null as any);
  const [isLoading, setIsLoading] = useState<Boolean>(true);

  useEffect(() => {
    // get file type
    const match = fileName.match(/(.+)(\.(zip|pdf))$/i);
    if (match) {
      setFileName(match[1]); // file name without the type
      setFileExtension(match[2]); // save the file type
    } else {
      setFileName(fileName);
      setFileExtension("");
    }
    initMoveFile();
  }, []);

  const initMoveFile = async () => {
    try {
      if (!context.pageContext.list) {
        throw new Error("List context is undefined");
      }
      const sites = await getSiteCollections(context);
      setOptions(sites);
      const serverRelativeUrl = context.pageContext.site.serverRelativeUrl;
      const defaultSite: Site = sites.filter(site => site.Url === serverRelativeUrl)[0];

      setSelectedSite(defaultSite);
      setIsLoading(false)
    } catch (err) {
      console.error("initMoveFile", err);
    }
  };

  const handleSiteChange = (newSite: string) => {
    const selected_site: Site | undefined = siteOptions.find(site => site.Title === newSite);

    if (selected_site) {
      setSelectedSite(selected_site);
    } else {
      console.warn('Selected site not found:', newSite);
    }
  }

  const onFolderSelect = async (folder: IFolder) => {
    isClose();
    try {
      const sp = getSPByPath(`https://epstin100.sharepoint.com${selectedSite?.Url}`, context);
      const targetLibraryUrl = folder.ServerRelativeUrl;
      let finalFileName = `${fileNameState.trim()}${fileExtension || ".pdf"}`;

      // Check the openDialogType to determine file creation
      let fileBlob: any;

      fileBlob = await toast.promise(
        ConvertToPdf(context, fileToSave),
        {
          loading: 'מאחד קבצים ל-PDF...',
          success: 'ה-PDF נוצר בהצלחה!',
          error: 'אירעה שגיאה בעת יצירת ה-PDF. אנא נסה שוב',
        }
      );

      // Save the generated file to the specified folder
      sp.web.getFolderByServerRelativePath(targetLibraryUrl).files.addUsingPath(finalFileName, fileBlob, { EnsureUniqueFileName: true });

    } catch (error) {
      console.error("Error creating link file:", error);
    }
  };

  if (isLoading) return null;
  return (
    <>
      <CacheProvider value={cacheRtl}>
        <div
          className={`${modalStyles.modalScreen}`}
          style={{ backgroundColor: "unset" }}
        // onClick={() => isClose()}
        >
          <div
            className={`${modalStyles.modal} ${modalStyles.folderExplorerModal}`}
            onClick={(ev: any) => ev.stopPropagation()}
          >
            <Autocomplete
              options={siteOptions.map(option => option.Title)}
              defaultValue={selectedSite?.Title}
              sx={{ width: 300, paddingBottom: '7%' }}
              renderInput={(params) => <TextField {...params} label="אתר" />}
              onChange={(event, newSite: string) => handleSiteChange(newSite)}
            />
            <TextField
              label="שם הקובץ"
              fullWidth
              value={fileNameState}
              onChange={(event) => setFileName(event.target.value)}
            />
            {/* <div className={styles.folderPicker}> */}

            <FolderExplorer
              context={context}
              key={"folderexplorer" + (selectedSite?.Title || "")}
              onSelect={(folder) => setSelectedFolder(folder)}
              canCreateFolders={false}
              defaultFolder={{
                Name: "ספריות",
                ServerRelativeUrl: `${selectedSite?.Url}`,
              }}
              rootFolder={{
                Name: "ספריות",
                ServerRelativeUrl: `${selectedSite?.Url}`,
              }}
              siteAbsoluteUrl={`https://epstin100.sharepoint.com${selectedSite?.Url}`}
            />
            {/* </div> */}

            <div className={`${modalStyles.modalFooter} ${modalStyles.modalFooterEnd}`}>
              <Button
                color="error"
                className={styles.button}
                onClick={() => isClose()}
                startIcon={<IconButton disableRipple style={{ color: "#f58383", paddingLeft: 0, margin: "0px !important" }}><CloseIcon /></IconButton>}
                sx={{
                  "& .MuiButton-startIcon": {
                    margin: 0,
                  },
                }}
              >
                סגירה</Button>
              <Button
                disabled={selectedFolder && selectedFolder?.ServerRelativeUrl === context.pageContext.web.serverRelativeUrl}
                onClick={() => onFolderSelect(selectedFolder)}
                className={styles.button}
                endIcon={<IconButton disableRipple style={{ color: "#1976d2", margin: "0px" }}><CheckIcon /></IconButton>}
              >שמירה</Button>
            </div>
          </div>
        </div>
      </CacheProvider>

    </>
  );
}
