import JSZip from "jszip";
import { SelectedFile } from "../../models/global.model";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";

export interface ExportZipModalProps{
    status: boolean;
    unMountDialog: () => void;
    selectedItems: SelectedFile[];
    context: ListViewCommandSetContext;
    sp: SPFI;
}