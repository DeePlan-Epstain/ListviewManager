import { SelectedFile, File } from '../models/global.model';
import JSZip, { file } from 'jszip';
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import getSP from '../../../pnpjs-config';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/folders";
import { folderFromServerRelativePath } from "@pnp/sp/folders";
import { IFileInfo } from '@pnp/sp/files';
import "@pnp/sp/webs";


const DELAY_TIME = 5000;
const MAX_HTTP_CALLS = 200;
let HTTP_COUNTER = 0;


// a function that cause a delay
function delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

// a function that count the number of http calls made and delay if needed to prevent throttle
async function delayIfNeeded(){
    HTTP_COUNTER++;
    console.log(`http call counter: ${HTTP_COUNTER}`);
    if(HTTP_COUNTER >= MAX_HTTP_CALLS){
        HTTP_COUNTER = 0;
        await delay(DELAY_TIME);
    }
}

// main function that trigger the recursive calls
export async function exportToZip(selectedItems: SelectedFile[], context: ListViewCommandSetContext ){
    console.log("context: ", context);
    
    const archive = new JSZip();
    console.log(selectedItems);
    const rootFolderName = getRootFolder(selectedItems[0]);
    const sp = getSP(context);
    //handle the root folder
    for(const item of selectedItems){
        if(item.FSObjType === "1"){// folder
            await handleFolder(rootFolderName, archive, item, sp);
        }
        else{// file
            await handleFile(rootFolderName, archive, item, sp);
        }
    }
    return archive;
}

// return the root folder of the zip
function getRootFolder(selectedFile: File){
    const root = selectedFile.FileRef.split('/');
    console.log("root: ", root[root.length - 2]);
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
async function handleFile(rootFolder: string, archive: JSZip, selectedFile: File, sp: SPFI){
    try{
        if (!sp || !selectedFile.FileRef) {
            console.error("Invalid SP object or FileRef in handleFile.");
            return;
        }
        const fileContent = await sp.web.getFileByServerRelativePath(selectedFile.FileRef).getBuffer();
        delayIfNeeded();
        const filePath = getFolderPath(selectedFile, rootFolder);
        if(filePath){
            archive.file(filePath, fileContent);
            console.log("archive:", archive);
            
        }
    } catch(error){
        console.error(`Error downloading file: ${selectedFile.FileRef} error: `, error)
    }
}

async function handleFolder(rootFolder: string, archive: JSZip, selectedFile: File, sp: SPFI){
    try{
        const folder = archive.folder(getFolderPath(selectedFile, rootFolder));
        const selectedFolder = folderFromServerRelativePath(sp.web, selectedFile.FileRef);
        const folderInfo = await selectedFolder();
        delayIfNeeded()
        const files: IFileInfo[] = await selectedFolder.files();
        delayIfNeeded()
        const folders = await selectedFolder.folders();
        delayIfNeeded()
        //handle files within the folder
        for(const file of files){
            await handleFile(rootFolder, archive, { FileRef: file.ServerRelativeUrl, FileLeafRef: file.Name}, sp)
        }
        //handle sub folders
        for(const subfolder of folders){
            await handleFolder(rootFolder, archive, { FileRef: subfolder.ServerRelativeUrl, FileLeafRef: subfolder.Name}, sp)
        }
    }
    catch(error){
        console.error(error);
    }
}

// download the zip archive to my computer
export async function downloadToPC(archive:JSZip) {
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

        // Generate the zip file as a Blob
        const zipBlob = await archive.generateAsync({ type: "blob" });

        // Determine the target folder where the zip file will be saved
        const folderPath = selectedItems[0].FileRef.substring(0, selectedItems[0].FileRef.lastIndexOf('/'));

        // Name the zip file
        const zipFileName = `${rootFolderName}_ExportedFiles.zip`;

        console.log("begin uploading the zip file");
        
        // Upload the zip file to SharePoint
        await sp.web.getFolderByServerRelativePath(folderPath).files.addChunked(zipFileName, zipBlob);

        console.log(`Zip file successfully uploaded to SharePoint folder: ${folderPath}`);
    } catch (error) {
        console.error("Error saving zip file to SharePoint:", error);
    }
}