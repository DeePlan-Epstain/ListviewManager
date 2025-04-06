import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { getSP, getSPByPath } from "../../../pnpExt.config";
import { SPFI } from "@pnp/sp";
import {
    domain,
    epsteinSiteUrl,
    libraryNamesToSkip,
} from "../Models/constants";
import { CustomTreeItem, LoadAsyncLibraries, ExpandedFolder } from "../models/TreeView.mdl";
import { IFolderInfo } from "@pnp/sp/folders";
import { getLoadingItem, replaceHexadecimals, sortItems } from "./util.srv";

export class NavTreeService {
    private sp: SPFI;

    constructor(private context: ApplicationCustomizerContext) {
        this.sp = getSP(context);
    }

    public async getSiteLibraries(siteUrl: string): Promise<CustomTreeItem[]> {
        const siteSp = getSPByPath(siteUrl, this.context);

        try {
            const libraries = await siteSp.web.lists
                .filter(`BaseTemplate eq 101 and hidden eq false and EntityTypeName ne 'Shared_x0020_Documents' and ${this.getLibraryNamesToSkipStr()}`)
                .select("Id", "Title", "ParentWebUrl", "EntityTypeName")
                .orderBy("Title")
                ();

            const formattedLibraries: CustomTreeItem[] = libraries.map((lib) => ({
                id: crypto.randomUUID(),
                itemId: lib.Id,
                label: lib.Title,
                link: `${lib.ParentWebUrl}/${replaceHexadecimals(lib.EntityTypeName)}`,
                type: "library",
            }));

            const librariesRootFoldersPromise = formattedLibraries.map(
                async (lib): Promise<CustomTreeItem | null> => {
                    try {
                        const folders: ExpandedFolder[] = await siteSp.web.lists
                            .getById(lib.itemId).rootFolder.folders
                            .filter("Name ne 'Forms'")
                            .select("UniqueId", "Name", "ServerRelativeUrl")
                            .orderBy("Name")
                            .expand("properties")
                            ();

                        const fotmattedFolders: CustomTreeItem[] = folders.map((f) => ({
                            id: crypto.randomUUID(),
                            itemId: f.UniqueId,
                            label: f.Name,
                            link: f.ServerRelativeUrl,
                            type: "folder",
                            children: f.Properties.vti_x005f_foldersubfolderitemcount ? [getLoadingItem()] : []
                        }));

                        lib.children = fotmattedFolders;

                        return lib;
                    } catch (err) {
                        console.error("getSiteLibraries Error:", err);
                        return null;
                    }
                }
            );

            const librariesRootFolders: Array<CustomTreeItem | null> =
                await Promise.all(librariesRootFoldersPromise);

            return librariesRootFolders.filter((l) => !!l) as CustomTreeItem[];
        } catch (err) {
            console.error("getSiteLibraries Error:", err);
            return [];
        }
    }

    // public async getInitialState(): Promise<CustomTreeItem[]> {
    //     const portalSp = getSPByPath(epsteinSiteUrl, this.context);

    //     try {
    //         const sites = await portalSp.web.lists
    //             .getById("41d92fdd-1469-475b-8d19-9fe47cca24be")
    //             .items.orderBy("Order", true)
    //             .top(5000)();

    //         const sitesPromise = sites.map(async (l) => {
    //             const libraries = await this.getSiteLibraries(l.LinkUrl.Url);

    //             const item: CustomTreeItem = {
    //                 id: crypto.randomUUID(),
    //                 itemId: l.GUID,
    //                 label: l.Title,
    //                 link: l.LinkUrl.Url,
    //                 type: "site",
    //                 children: [...libraries],
    //             };

    //             // get subsites
    //             if (l.IsSubsitesExisted) {
    //                 const siteSp = getSPByPath(l.LinkUrl.Url, this.context);
    //                 const subsites = await siteSp.web
    //                     .getSubwebsFilteredForCurrentUser()
    //                     .orderBy("Title")();

    //                 const subsitesPromise = subsites.map(
    //                     async (s): Promise<CustomTreeItem> => {
    //                         const siteUrl = await this.getFullSiteUrl(s.ServerRelativeUrl);
    //                         const libraries = await this.getSiteLibraries(siteUrl);

    //                         return {
    //                             id: crypto.randomUUID(),
    //                             itemId: s.Id.toString(),
    //                             label: s.Title,
    //                             link: `${domain}/${s.ServerRelativeUrl}`,
    //                             type: "subsite",
    //                             children: [...libraries],
    //                         };
    //                     }
    //                 );

    //                 const subsitesChildren = await Promise.all(subsitesPromise);
    //                 item.children = [...libraries, ...subsitesChildren];
    //             }

    //             return item;
    //         });

    //         const LinksList = await Promise.all(sitesPromise);

    //         return LinksList;
    //     } catch (err) {
    //         console.error("getInitialState Error:", err);
    //         return [];
    //     }
    // }

    public async getInitialState(): Promise<CustomTreeItem[]> {
        const portalSp = getSPByPath(epsteinSiteUrl, this.context);

        try {
            const sites = await portalSp.web.lists
                .getById("41d92fdd-1469-475b-8d19-9fe47cca24be")
                .items.orderBy("Order", true)
                .top(5000)();

            const sitesPromise = sites.map(async (l) => {
                let libraries: CustomTreeItem[] = [];
                try {
                    libraries = await this.getSiteLibraries(l.LinkUrl.Url);
                } catch (err) {
                    console.warn(`🚨 Failed to fetch libraries for site: ${l.Title}`, err);
                }

                // Initialize the site item
                const item: CustomTreeItem = {
                    id: crypto.randomUUID(),
                    itemId: l.GUID,
                    label: l.Title,
                    link: l.LinkUrl.Url,
                    type: "site",
                    children: [], // Always initialize as an empty array
                };

                // Add libraries to children
                item.children = [...libraries];

                // Get subsites if they exist
                if (l.IsSubsitesExisted) {
                    try {
                        const siteSp = getSPByPath(l.LinkUrl.Url, this.context);
                        const subsites = await siteSp.web
                            .getSubwebsFilteredForCurrentUser()
                            .orderBy("Title")();

                        const subsitesPromise = subsites.map(async (s): Promise<CustomTreeItem> => {
                            try {
                                const siteUrl = await this.getFullSiteUrl(s.ServerRelativeUrl);
                                const subsiteLibraries = await this.getSiteLibraries(siteUrl);

                                return {
                                    id: crypto.randomUUID(),
                                    itemId: s.Id.toString(),
                                    label: s.Title,
                                    link: `${domain}/${s.ServerRelativeUrl}`,
                                    type: "subsite",
                                    children: [...subsiteLibraries],
                                };
                            } catch (err) {
                                console.warn(`🚨 Failed to fetch libraries for subsite: ${s.Title}`, err);
                                return {
                                    id: crypto.randomUUID(),
                                    itemId: s.Id.toString(),
                                    label: s.Title,
                                    link: `${domain}/${s.ServerRelativeUrl}`,
                                    type: "subsite",
                                    children: [],
                                };
                            }
                        });

                        const subsitesChildren = await Promise.all(subsitesPromise);
                        item.children = [...item.children, ...subsitesChildren];
                    } catch (err) {
                        console.warn(`🚨 Failed to fetch subsites for site: ${l.Title}`, err);
                    }
                }

                // Filter out the item if it has no children (no libraries and no subsites)
                if (item.children.length === 0) {
                    console.warn(`Skipping site ${l.Title} because it has no libraries or subsites.`);
                    return null;
                }

                return item;
            });

            // Filter out null results after mapping
            const LinksList = (await Promise.all(sitesPromise)).filter((item): item is CustomTreeItem => item !== null);

            return LinksList;
        } catch (err) {
            console.error("getInitialState Error:", err);
            return [];
        }
    }

    public async loadLibrariesAsync(TreeItems: CustomTreeItem[], expandedNodes: string[]): Promise<{
        TreeItems: CustomTreeItem[];
        LoadAsyncLibraries: LoadAsyncLibraries[];
        ExpandedNodes: string[];
    } | null> {
        let TreeItemsCopy: CustomTreeItem[] = [...TreeItems];
        let currentLoadAsyncLibraries: LoadAsyncLibraries[] = [];
        let currentExpandedNodes = [...expandedNodes];

        try {
            // Getting all libraries that require to be loaded async
            currentLoadAsyncLibraries = JSON.parse(
                sessionStorage.getItem("LoadAsyncLibraries") || ""
            );

            for (let j = 0; j < currentLoadAsyncLibraries.length; j++) {
                const CurrLibrary = currentLoadAsyncLibraries[j];
                var ListOfFolders: CustomTreeItem[] = [];

                // Fetch all library folders
                const LibraryFolders =
                    (await this.GetLibraryFoldersWhenFailed(
                        CurrLibrary.Id,
                        CurrLibrary.LibraryLink
                    )) || [];

                for (let i = 0; i < LibraryFolders.length; i++) {
                    const CurrFolder = LibraryFolders[i];

                    // Get folder's props
                    const CurrFolderProps: ExpandedFolder = await this.sp.web
                        .getFolderByServerRelativePath(CurrFolder.FileRef)
                        .expand("properties")();

                    // Skip folders with the name "Forms"
                    if (CurrFolder.Name === "Forms") continue;

                    // Determines whether folder has subFolders
                    const SubFoldersCount =
                        CurrFolderProps["Properties"]["vti_x005f_foldersubfolderitemcount"];

                    // Insert current folder's object
                    ListOfFolders.push({
                        id: crypto.randomUUID(),
                        itemId: CurrFolderProps.UniqueId,
                        label: CurrFolderProps.Name,
                        link: CurrFolderProps.ServerRelativeUrl,
                        type: "folder",
                        children: SubFoldersCount === 0 ? undefined : [getLoadingItem()],
                    });

                    // Replace the new children folders array with the existing one
                    TreeItemsCopy = this.GetUpdatedTreeItemsRecursively(
                        TreeItemsCopy,
                        CurrLibrary.Id,
                        ListOfFolders
                    );

                    // Remove loaded library from LoadLibrariesAsync
                    currentLoadAsyncLibraries = currentLoadAsyncLibraries.filter(
                        (l: any) => l.Id !== CurrLibrary.Id
                    );

                    // Remove fetched library from expanded nodes array when done
                    currentExpandedNodes = currentExpandedNodes.filter(
                        (Id) => Id !== CurrLibrary.Id
                    );
                }
            }

            return {
                TreeItems: TreeItemsCopy,
                LoadAsyncLibraries: currentLoadAsyncLibraries,
                ExpandedNodes: currentExpandedNodes,
            };
        } catch (error) {
            console.error("loadLibrariesAsync Error:", error);
            return null;
        }
    }

    private async GetLibraryFoldersWhenFailed(
        LibraryId: string,
        LibraryLink: string
    ) {
        try {
            // Get all library items.
            const AllFolders = await this.sp.web.lists
                .getById(LibraryId)
                .items.select("FileLeafRef", "FileRef")
                .expand("ContentType")
                .getAll();

            // Filter folders only.
            const FilteredRootFolderFolders = AllFolders.filter(
                (CurrItem: any, Index: number) => {
                    if (
                        (CurrItem.ContentType.Name === "תיקיה" ||
                            CurrItem.ContentType.Name === "CustomTreeItem") &&
                        CurrItem.FileRef === `${LibraryLink}/${CurrItem.FileLeafRef}`
                    ) {
                        return CurrItem;
                    }
                }
            );

            return sortItems.text(FilteredRootFolderFolders, "FileLeafRef");
        } catch (error) { }
    }

    public GetUpdatedTreeItemsRecursively(UpdatedTreeItems: CustomTreeItem[], TreeItemId: string, NewChildrens: string | CustomTreeItem[]): CustomTreeItem[] {
        for (let i = 0; i < UpdatedTreeItems.length; i++) {
            const CurrItem = UpdatedTreeItems[i];
            const Childrens = CurrItem?.children || [];

            if (TreeItemId === CurrItem.id) {
                // If the wanted node is found.
                if (Array.isArray(NewChildrens))
                    CurrItem.children = sortItems.text(NewChildrens, "label");

                return UpdatedTreeItems;
            } else if (Childrens?.length) {
                // If wanted node wasn't found, continue looking in it's children array.

                const NewChildrenTreeItems = this.GetUpdatedTreeItemsRecursively(Childrens, TreeItemId, NewChildrens);

                if (NewChildrenTreeItems) {
                    CurrItem.children = NewChildrenTreeItems;
                    // return UpdatedTreeItems;
                }
            }
        }

        return UpdatedTreeItems;
    }

    public async getFullSiteUrl(serverRelativeUrl: string): Promise<string> {
        const splittedUrl = serverRelativeUrl.split('/');
        const relativeUrl = splittedUrl.length > 4 ? splittedUrl.slice(0, 4).join('/') : serverRelativeUrl;

        try {
            // Get the web information for the given server-relative URL
            return await this.sp.site.getWebUrlFromPageUrl(domain + relativeUrl);
        } catch (error) {
            console.error("Error fetching site URL:", error);
            console.log('domain + serverRelativeUrl:', domain + serverRelativeUrl)
            throw new Error(`Unable to fetch site URL for: ${serverRelativeUrl}`);
        }
    }

    public GetFolderChildrenFolders(ServerRelativeUrl: string): Promise<CustomTreeItem[]> {
        return new Promise(async (resolve) => {
            try {
                const siteUrl = await this.getFullSiteUrl(ServerRelativeUrl);
                const siteSp = getSPByPath(siteUrl, this.context);
                const folders = await siteSp.web.getFolderByServerRelativePath(ServerRelativeUrl).folders
                    .select("UniqueId", "Name", "ServerRelativeUrl")
                    .filter("Name ne 'Forms'")
                    .expand("properties")();

                const res: CustomTreeItem[] = folders.map((f: IFolderInfo & { Properties: { vti_x005f_foldersubfolderitemcount: number }; }) => {
                    const subFoldersCount: number = f["Properties"]["vti_x005f_foldersubfolderitemcount"];

                    const newFolder: CustomTreeItem = {
                        id: crypto.randomUUID(),
                        itemId: f.UniqueId,
                        label: f.Name,
                        type: "folder",
                        link: f.ServerRelativeUrl,
                        children: subFoldersCount ? [getLoadingItem()] : [],
                    };

                    return newFolder;
                }
                );

                resolve(res);
            } catch (error) {
                console.log("GetFolderChildrenFolders:", error);
                resolve([]);
            }
        });
    }

    public GetUpdatedSavedNodesServerRelativeUrl(
        NodeId: string,
        ServerRelativeUrl: string,
        NodesServerRelativeUrls: any[]
    ) {
        let NodesUrlsCopy = [...NodesServerRelativeUrls];

        let IsNodeUrlSaved = false;
        // Validate node wasn't saved already.
        NodesUrlsCopy.forEach((N) => {
            if (NodeId === N.NodeId) {
                IsNodeUrlSaved = true;
            }
        });

        // Save node if wasn't saved yet.
        if (IsNodeUrlSaved === false) {
            NodesUrlsCopy.push({
                NodeId,
                ServerRelativeUrl,
            });
        }

        return NodesUrlsCopy;
    }

    public filterTreeItems(treeItems: CustomTreeItem[], filter: string, expandedNodes: Set<string>): CustomTreeItem[] {
        if (!filter) return treeItems;

        const filteredTreeItems: (CustomTreeItem | null)[] = treeItems.map((item): CustomTreeItem | null => {
            const children = item.children || [];

            // Filter children recursively.
            // If filteredChildrens is empty, it means that this item does not contain any items
            // that match the filter, so we should not show the arrow to open it.
            const filteredChildrens = this.filterTreeItems(children, filter, expandedNodes);

            const isItemMatched = item.label.toLowerCase().includes(filter.toLowerCase());

            if (isItemMatched && !children.length) {

                return {
                    ...item,
                    children: filteredChildrens.length ? filteredChildrens : children
                };
            }

            // If the item found or has children, return it.
            if (isItemMatched || filteredChildrens.length) {

                if (filteredChildrens.length) expandedNodes.add(item.id);

                return {
                    ...item,
                    children: filteredChildrens.length ? filteredChildrens : children,
                };
            }

            return null;
        }
        );

        return filteredTreeItems.filter((i) => !!i) as CustomTreeItem[];
    }

    public flattenTreeItems(
        treeItems: CustomTreeItem[],
        allItems: CustomTreeItem[] = []
    ): CustomTreeItem[] {
        treeItems.forEach((item) => {
            allItems.push(item);

            if (item.children && item.children.length > 0) {
                this.flattenTreeItems(item.children, allItems);
            }
        });

        return allItems;
    }

    public getLibraryNamesToSkipStr(key?: string): string {
        return libraryNamesToSkip
            .map((l) => `${key || "EntityTypeName"} ne '${l}'`)
            .join(" and ");
    }
}
