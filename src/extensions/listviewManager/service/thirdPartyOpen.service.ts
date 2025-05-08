import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SearchQueryBuilder } from "@pnp/sp/search";
import { getGraph } from "../../../pnpjs-config";
import { SPFI, SPFx } from "@pnp/sp";

export async function getUserAADId(context: ListViewCommandSetContext): Promise<string> {
    try {
        const user = await getGraph(context).me();
        return user.id || '';
    } catch (err) {
        console.error("Error fetching user AAD ID:", err);
        return '';
    }
}

export async function clickEvent(context: ListViewCommandSetContext) {
    const userId = await getUserAADId(context);
    const listID = context.pageContext.list?.id.toString();
    const siteId = context.pageContext.site?.id.toString();
    const userEmail = context.pageContext.user?.email;
    const webUrl = context.pageContext.site?.absoluteUrl;
    const sp: SPFI = new SPFI('https://epstin100.sharepoint.com/sites/EpsteinPortal').using(SPFx(context));
    const validTypesListId = 'd88944fd-fa4c-4d7e-8334-df52e5e5e247';
    let validTypes: string[] = []

    try {
        const typesArr = await sp.web.lists.getById(validTypesListId).items.select('Title')();
        validTypes = typesArr.map(type => type.Title);
    } catch (error) {
        console.error("Error fetching validTypes:", error);
    }

    // Event Listener for mouse clicks:
    window.addEventListener('click', (event) => {
        const target = event.target as HTMLElement;
        const fileName = target.innerText.trim();

        // const isFieldRenderer = target.getAttribute("data-automationid") === "FieldRenderer-name";  // no longer avilable
        const isheroField = target.getAttribute("data-id") === "heroField";
        const isValidType = checkValidType(fileName, validTypes);

        // get out if type is not valid (not: dpf, dwg, msg...)
        if (!isValidType || !isheroField) return;

        const dataActionsAttr = target.getAttribute("data-actions");
        let fileSpId: string | undefined;

        try {
            if (dataActionsAttr) {
                const dataActions = JSON.parse(dataActionsAttr);
                const heroFieldAction = dataActions.find((a: any) => a.key === "heroFieldHoverTarget");
                if (heroFieldAction?.data) {
                    const data = JSON.parse(heroFieldAction.data);
                    fileSpId = data.itemKey;
                }
            }
        } catch (err) {
            console.error("Failed to parse data-actions:", err);
        }

        handleLinkClick(event, siteId, listID, userEmail, webUrl, userId, fileName, sp, fileSpId)
    }, true);
    return Promise.resolve();
}

export async function handleLinkClick(event: MouseEvent, siteId: string, listId: string, userEmail: string, webUrl: string, userId: string, fileName: string, sp: SPFI, fileSpId: string): Promise<void> {
    event.preventDefault();
    event.stopPropagation();
    let fileId = ''
    try {
        const query = SearchQueryBuilder(`filename:"${fileName}"`)
            .selectProperties("UniqueId", "ListItemId")
            .refinementFilters(`SiteId:${siteId}`)
            .rowLimit(10);
        const results = await sp.search(query);
        if (results.PrimarySearchResults.length > 0) {
            let firstResult = results.PrimarySearchResults[0] as Record<string, string>;

            // if there is more the 1 file with the same name - set firstResult by id
            if (results.PrimarySearchResults.length > 1) {
                firstResult = results.PrimarySearchResults.find((res: Record<string, string>) => res["ListItemId"] === fileSpId) as Record<string, string>;
            }

            fileId = firstResult["IdentityListItemId"];
            if (!fileId) fileId = firstResult["UniqueId"];
        } else {
            console.warn("No results found for file:", fileName);
        }
    } catch (error) {
        console.error("Error searching fileId:", error);
    }

    openFileInApp(siteId, listId, userEmail, webUrl, userId, fileId, fileName);
}

export function checkValidType(fileName: string, validTypes: string[]): boolean {
    let isValidType = false;

    validTypes.forEach(type => {
        if (fileName.toLocaleLowerCase().endsWith(type)) {
            isValidType = true;
        }
    })
    return isValidType;
}

export function openFileInApp(siteId: string, listId: string, userEmail: string, webUrl: string, userId: string, fileId: string, fileName: string,): void {

    const odopenUrl = `odopen://openFile/?fileId=${encodeURIComponent(fileId)}&siteId=${encodeURIComponent(siteId)}&listId=${encodeURIComponent(listId)}&userEmail=${encodeURIComponent(userEmail)}&userId=${encodeURIComponent(userId)}&webUrl=${encodeURIComponent(webUrl)}&fileName=${encodeURIComponent(fileName)}`;

    window.location.href = odopenUrl;
}