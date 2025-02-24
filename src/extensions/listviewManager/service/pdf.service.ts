import { SelectedFile } from "../models/global.model";
import axios, { AxiosError } from "axios";
import { SPFI, SPFx } from "@pnp/sp";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import PAService from "./authService.service";
import Swal from "sweetalert2";
import { Errors } from "../../errorConfig";

export async function getConvertibleTypes(context: ListViewCommandSetContext) {
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
    const paService = new PAService(context, "https://prod-48.westeurope.logic.azure.com:443/workflows/60282bd80e29428c9094b301317c665c/triggers/manual/paths/invoke?api-version=2016-06-01");

    let baseUrl = 'https://epstin100.sharepoint.com';

    const siteAddress: string = baseUrl + getRelativeSite(selectedItem.FileRef);
    const libraryId: string = context.pageContext.list.id["_guid"];
    const itemId: number = selectedItem.ID;

    try {
        // let token = await getAccessToken(clientId, secret);
        const requestBody = {
            siteAddress: siteAddress,
            libraryId: libraryId,
            itemId: itemId
        }
        const data = await paService.post(paService.CONVERT_TO_PDF, requestBody);
        // if(data?.ok)
        Swal.fire({
            title: "הפעולה בוצעה בהצלחה!",
            text: "הקובץ הומר לPDF בהצלחה",
            icon: "success"
        });
    }
    catch (error) {
        if (error.response && error.response.status === 400) {
            Swal.fire({
                icon: "error",
                title: "שגיאה",
                text: Errors.CONVERT_TO_PDF_FAILED_EMPTY_EXCEL,
            });
        }
        else if (error.response && error.response.status === 401) {
            Swal.fire({
                icon: "error",
                title: "שגיאה",
                text: Errors.CONVERT_TO_PDF_FAILED_ALREADY_EXIST,
            });
        }
    }
}

