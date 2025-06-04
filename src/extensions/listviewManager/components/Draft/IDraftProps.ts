import { SPFI } from "@pnp/sp";
import { DraftProperties } from "../../models/global.model";
import { IService } from "./models/IService";

export interface IDraftProps {
    close: () => void;
    submit: (eMailProperties: DraftProperties) => void;
    draftProperties: DraftProperties;
    createDraft: IService;
    sp: SPFI;
}


