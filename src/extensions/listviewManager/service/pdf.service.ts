import { SelectedFile } from "../models/global.model";
import axios from "axios";
import { SPFI, SPFx } from "@pnp/sp";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";

export async function getConvertibleTypes(context: ListViewCommandSetContext){
    const sp: SPFI = new SPFI('https://epstin100.sharepoint.com/sites/EpsteinPortal').using(SPFx(context));
    const listId = 'b748b7b9-6b49-44f9-b889-ef7e99ebdb47';
    const listItems = await sp.web.lists.getById(listId).items.getAll();
    const types = listItems.map(item => item.Title)
    const typeSet = new Set<string>();
    types.forEach(type => typeSet.add(type));
    return typeSet;
}

function getRelativeSite(fileRef: string) {
    let parts = fileRef.split('/');
    parts = parts.slice(0, 3);
    const output = parts.join('/');
    return output
}

// this function send a http to power automate to convert the file to pdf
export async function ConvertToPdf(context: ListViewCommandSetContext, selectedItem: SelectedFile) {


    let baseUrl = 'https://epstin100.sharepoint.com/';

    const siteAddress: string = baseUrl + getRelativeSite(selectedItem.FileRef);
    const libraryId: string = context.pageContext.list.id["_guid"];
    const itemId: number = selectedItem.ID;

    try {
        // let token = await getAccessToken(clientId, secret);
        const url = 'https://prod-48.westeurope.logic.azure.com:443/workflows/60282bd80e29428c9094b301317c665c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BuR0ghAmSlKrA0XyXtDcjfWF9KkCqAkViqYwWYI4JEg';
        const requestBody = {
            siteAddress: siteAddress,
            libraryId: libraryId,
            itemId: itemId
        }
        
        await axios.post(url,requestBody, {
            headers: {
                'Content-Type': 'application/json'
            }
        })
    }
    catch (error) {
        console.log("an error found: ", error);
    }
}

