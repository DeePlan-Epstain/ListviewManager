import { ICommandDefinition } from "@microsoft/sp-module-interfaces";

export type CMD = ICommandDefinition & {
    id: string;
    visible?: boolean;
};