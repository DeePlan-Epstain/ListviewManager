import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { getGraph, getSP } from "../../../pnpjs-config";

export async function getUserAADId(context: ListViewCommandSetContext): Promise<string> {
    try {
        const user = await getGraph(context).me();
        return user.id || '';
    } catch (err) {
        console.error("Error fetching user AAD ID:", err);
        return '';
    }
}

export async function getFileGUID(fileUrl: string, context: ListViewCommandSetContext): Promise<string> {
    try {
        const sp = getSP(context);
        const file = await sp.web.getFileByUrl(fileUrl)();
        return file.UniqueId;
    } catch (err) {
        console.error("Error fetching file GUID:", err);
        return '';
    }
}

export async function clickEvent(context: ListViewCommandSetContext) {
    const userId = await getUserAADId(context);
    const listID = context.pageContext.list?.id.toString();
    const siteId = context.pageContext.site?.id.toString();
    const userEmail = context.pageContext.user?.email;
    const webUrl = context.pageContext.site?.absoluteUrl;

    window.addEventListener('click', (event) => {
        const target = event.target as HTMLElement;
        const fileName = target.innerText.trim();
        const isFieldRenderer = target.getAttribute("data-automationid") === "FieldRenderer-name";

        if (!isPdfOrDwg(fileName) || !isFieldRenderer) return;

        handleLinkClick(event, siteId, listID, userEmail, webUrl, userId, fileName, context)
    }, true);
    return Promise.resolve();
}

export async function handleLinkClick(event: MouseEvent, siteId: string, listId: string, userEmail: string, webUrl: string, userId: string, fileName: string, context: ListViewCommandSetContext): Promise<void> {
    event.preventDefault();
    event.stopPropagation();

    const fileUrl = await buildFileUrl(fileName, context) //builds a full link
    let fileId = ''

    try {
        fileId = await getFileGUID(fileUrl, context);
    }
    catch (err) {
        console.error('error fetching file GUID Id', err)
    }
    openFileInApp(siteId, listId, userEmail, webUrl, userId, fileId, fileName);
}

export function isPdfOrDwg(fileName: string): boolean {
    return fileName.toLowerCase().endsWith('.pdf') || fileName.toLowerCase().endsWith('.dwg');
}

export function openFileInApp(siteId: string, listId: string, userEmail: string, webUrl: string, userId: string, fileId: string, fileName: string,): void {

    const odopenUrl = `odopen://openFile/?fileId=${encodeURIComponent(fileId)}&siteId=${encodeURIComponent(siteId)}&listId=${encodeURIComponent(listId)}&userEmail=${encodeURIComponent(userEmail)}&userId=${encodeURIComponent(userId)}&webUrl=${encodeURIComponent(webUrl)}&fileName=${encodeURIComponent(fileName)}`;

    window.location.href = odopenUrl;
}

export async function buildFileUrl(fileName: string, context: ListViewCommandSetContext): Promise<string> {

    const webUrl = context.pageContext.web.absoluteUrl;  // https://Sname.sharepoint.com/sites/Wname
    const relativeUrl = context.pageContext.list?.serverRelativeUrl;  // /sites/Wname/Lib
    const baseUrl = webUrl.split('.com')[0] + '.com'; // https://Sname.sharepoint.com

    // full link
    const fullFileUrl = `${baseUrl}${relativeUrl}/${fileName}`;

    return fullFileUrl;
}