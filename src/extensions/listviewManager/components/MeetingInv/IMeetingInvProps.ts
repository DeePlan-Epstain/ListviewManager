import { EventProperties } from "../../models/global.model";
import { IService } from "./models/IService";


export interface IMeetingInvProps {
    close: () => void;
    submit: (eMailProperties: EventProperties) => void;
    eventProperties: EventProperties;
    createEvent: IService;
}


