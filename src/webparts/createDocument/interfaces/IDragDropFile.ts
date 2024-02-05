import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDragDropFileProps {
    context: WebPartContext;
    returnFileData(fileData: any): any;
}
export interface IDragDropFileState {
    files: any;
}