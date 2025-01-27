import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPFx } from "@pnp/sp";
import { spfi, SPFI } from "@pnp/sp";
import { Site } from "../models/LinkToCategory";

export async function getSiteCollections(context: ListViewCommandSetContext): Promise<Site[]> {
    const sp: SPFI = new SPFI('https://epstin100.sharepoint.com/sites/EpsteinPortal').using(SPFx(context));
    const items = await sp.web.lists.getById('19b0d26b-14ae-4dc1-972d-d0ac441c4952').items.select('Title', 'Url')();
    return items;
}

export function createSitesOptions(items: { Title: string, Url: string }[]) {
    return items.map(item => item.Title);
}

export function createSitesMap(items: { Title: string, Url: string }[]) {
    const map: Map<string, string> = new Map();
    items.forEach(item => map.set(item.Title, item.Url));
    return map;

}