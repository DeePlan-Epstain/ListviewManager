
export interface ModalExtStates {
    open: boolean,
    error:boolean,
    isSave:boolean
    FoldersHierarchy:Array<any>
    FolderHierarchy:any,
    FolderHierarchyValidate:boolean
    NewFolderName:string,
    NewFolderNameValidate:boolean
    success:boolean
    FoldersHierarchyAfterChoosingDivision:Array<any>
    DivisionValidate:boolean
    Division:string
}