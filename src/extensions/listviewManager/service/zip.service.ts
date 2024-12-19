import { SelectedFile, File } from '../models/global.model';
import JSZip, { file } from 'jszip';
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import getSP from '../../../pnpjs-config';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/folders";
import { folderFromServerRelativePath } from "@pnp/sp/folders";
import { IFileInfo } from '@pnp/sp/files';




export async function exportToZip(selectedItems: SelectedFile[], context: ListViewCommandSetContext ){
    const archive = new JSZip();
    console.log(selectedItems);
    const rootFolderName = getRootFolder(selectedItems[0]);
    const sp = getSP(context);
    //handle the root folder
    selectedItems.sort((a, b) => parseInt(a.FSObjType) - parseInt(b.FSObjType));
    // sp.web.getFolderByServerRelativePath('/sites/RailwaysDepartment/DocLib/שם לקוח_שם פרויקט_מספר פרויקט').
    for(const item of selectedItems){
        if(item.FSObjType === "1"){// folder
            await handleFolder(rootFolderName, archive, item, sp);
        }
        else{// file
            await handleFile(rootFolderName, archive, item, sp);
        }
    }

    //download for my computer
    archive.generateAsync({ type: "blob" }).then((content) => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(content);
        link.download = `${rootFolderName}.zip`;
        link.click();
    });
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
        const fileContent = await sp.web.getFileByServerRelativePath(selectedFile.FileRef).getBuffer();
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
        const files: IFileInfo[] = await selectedFolder.files();
        const folders = await selectedFolder.folders();
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
