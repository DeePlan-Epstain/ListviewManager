import { IFolderInfo } from "@pnp/sp/folders";

export type TreeItemTypes =
    | "site"
    | "subsite"
    | "library"
    | "folder"
    | "loading";

export interface CustomTreeItem {
    id: string;
    itemId: string;
    label: string;
    link: string;
    type: TreeItemTypes;
    children?: CustomTreeItem[];
}

export interface LoadAsyncLibraries {
    Id: string;
    LibraryLink: string;
}

export interface ExpandedFolder extends IFolderInfo {
    Properties: {
        vti_x005f_foldersubfolderitemcount: number;
    };
}
