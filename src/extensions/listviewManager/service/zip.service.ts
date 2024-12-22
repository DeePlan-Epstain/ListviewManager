import { SelectedFile, File } from '../models/global.model';
import JSZip, { file } from 'jszip';
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { getSP } from '../../../pnpjs-config';
import { SPFI } from '@pnp/sp';
import { folderFromServerRelativePath } from "@pnp/sp/folders";
import { IFileInfo } from '@pnp/sp/files';


const DELAY_TIME = 5000;
const MAX_HTTP_CALLS = 200;
let HTTP_COUNTER = 0;


// a function that cause a delay
function delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

// a function that count the number of http calls made and delay if needed to prevent throttle
async function delayIfNeeded() {
    HTTP_COUNTER++;
    if (HTTP_COUNTER >= MAX_HTTP_CALLS) {
        HTTP_COUNTER = 0;
        await delay(DELAY_TIME);
    }
}

// main function that trigger the recursive calls
export async function exportToZip(selectedItems: SelectedFile[], context: ListViewCommandSetContext) {
    console.time("exportToZip")
    const archive = new JSZip();
    const rootFolderName = getRootFolder(selectedItems[0]);
    const sp = getSP(context);
    //handle the root folder
    for (const item of selectedItems) {
        if (item.FSObjType === "1") {// folder
            await handleFolder(rootFolderName, archive, item, sp);
        }
        else {// file
            await handleFile(rootFolderName, archive, item, sp);
        }
    }
    console.timeEnd("exportToZip")
    return archive;
}

// return the root folder of the zip
function getRootFolder(selectedFile: File) {
    const root = selectedFile.FileRef.split('/');
    return root[root.length - 2];
}

// Get the path to the folder from the root folder in the archive
function getFolderPath(selectedFile: SelectedFile | File, rootFolderName: string) {
    const pathArray = selectedFile.FileRef.split(rootFolderName);
    if (pathArray.length < 2 || !pathArray[1]) {
        return selectedFile.FileLeafRef;
    }
    return pathArray[1].replace(/^\/|\/$/g, ''); // Remove leading and trailing slashes
}

// This function get a selected item and add its content to the new zip archive
async function handleFile(rootFolder: string, archive: JSZip, selectedFile: File, sp: SPFI) {
    try {
        const fileContent = await sp.web.getFileByServerRelativePath(selectedFile.FileRef).getBuffer();
        delayIfNeeded();
        const filePath = getFolderPath(selectedFile, rootFolder);
        if (filePath) {
            archive.file(filePath, fileContent);
        }
    } catch (error) {
        console.error(`Error downloading file: ${selectedFile.FileRef} error: `, error)
    }
}

// This function get a selected item and add its content to the new zip archive
function handleFileBatched(rootFolder: string, archive: JSZip, selectedFile: File, sp: SPFI, fileContent: ArrayBuffer) {
    try {
        // const fileContent = await sp.web.getFileByServerRelativePath(selectedFile.FileRef).getBuffer();
        // delayIfNeeded();
        const filePath = getFolderPath(selectedFile, rootFolder);
        if (filePath) {
            archive.file(filePath, fileContent);
        }
    } catch (error) {
        console.error(`Error downloading file: ${selectedFile.FileRef} error: `, error)
    }
}

async function handleFolder(rootFolder: string, archive: JSZip, selectedFile: File, sp: SPFI) {
    try {
        archive.folder(getFolderPath(selectedFile, rootFolder));
        const selectedFolder = folderFromServerRelativePath(sp.web, selectedFile.FileRef);
        // const folderInfo = await selectedFolder();
        let batchcounter = 0;
        const files: IFileInfo[] = await selectedFolder.files();
        const folders = await selectedFolder.folders();
        let [batchedWeb, execute] = sp.web.batched();
        let res: ArrayBuffer[] = [];
        //handle files within the folder
        for (const file of files) {
            batchedWeb.getFileByServerRelativePath(file.ServerRelativeUrl).getBuffer().then(content => res.push(content));
            batchcounter++;
            if (batchcounter % 100 === 0) {
                await execute();
                batchcounter = 0;
                [batchedWeb, execute] = sp.web.batched();
            }
        }
        files.forEach(async (file, index) => {
            handleFileBatched(rootFolder, archive, { FileRef: files[index].ServerRelativeUrl, FileLeafRef: files[index].Name }, sp, res[index])
        })
        //handle sub folders
        for (const subfolder of folders) {
            await handleFolder(rootFolder, archive, { FileRef: subfolder.ServerRelativeUrl, FileLeafRef: subfolder.Name }, sp)
        }
    }
    catch (error) {
        console.error(error);
    }
}

// download the zip archive to my computer
export async function downloadToPC(archive: JSZip) {
    archive.generateAsync({ type: "blob" }).then((content) => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(content);
        link.download = `Files from Sharepoint.zip`;
        link.click();
    });
}

// Function to save the new zip file in the same folder as the selected items
export async function saveZipToSharePoint(archive: JSZip, selectedItems: SelectedFile[], sp: SPFI) {
    try {
        const rootFolderName = getRootFolder(selectedItems[0]);
        const zipBlob = await archive.generateAsync({ type: "blob" });
        const folderPath = selectedItems[0].FileRef.substring(0, selectedItems[0].FileRef.lastIndexOf('/'));
        const zipFileName = `${rootFolderName}_ExportedFiles.zip`;
        // Upload the zip file to SharePoint
        await sp.web.getFolderByServerRelativePath(folderPath).files.addChunked(zipFileName, zipBlob);
    } catch (error) {
        console.error("Error saving zip file to SharePoint:", error);
    }
}
