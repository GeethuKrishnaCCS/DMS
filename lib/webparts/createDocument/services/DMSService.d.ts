import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
export declare class DMSService extends BaseService {
    private _spfi;
    private currentContext;
    private spQdms;
    constructor(context: WebPartContext, qdmsURL: string);
    getItems(siteUrl: string, listname: string): Promise<any>;
    getCurrentUser(): Promise<any>;
    getUserIdByEmail(email: string): Promise<any>;
    createNewItem(siteUrl: string, listname: string, metadata: any): Promise<any>;
    getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any>;
    getItemsByID(siteUrl: string, listname: string, id: number): Promise<any>;
    getItemsFromDepartments(siteUrl: string, listname: string): Promise<any>;
    uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any>;
    itemUpdate(siteUrl: string, listname: string, id: number, metadata: any): Promise<any>;
    itemFromTemplate(siteUrl: string, listname: string): Promise<any>;
    getBuffer(siteUrl: string): Promise<any>;
    documnetPath(uniqueId: any): Promise<any>;
    itemFromLibrary(siteUrl: string, listname: string): Promise<any>;
    itemFromLibraryUpdate(siteUrl: string, listname: string, id: number, metadata: any): Promise<any>;
    getQDMSPermissionWebpart(siteUrl: string, listname: string): Promise<any>;
    DocumentPermission(siteUrl: string, listname: string): Promise<any>;
    DocumentSendForReview(siteUrl: string, listname: string): Promise<any>;
    DocumentPublish(siteUrl: string, listname: string): Promise<any>;
    itemFromLibraryByID(siteUrl: string, listname: string, id: number): Promise<any>;
    itemFromPrefernce(siteUrl: string, listname: string, emailUser: string): Promise<any>;
    sendMail(emailPostBody: any): Promise<any>;
    getItemFromSIbFunction(siteUrl: string, listname: string, id: number): Promise<any>;
    getItemOnSelectCategory(siteUrl: string, listname: string, id: number): Promise<any>;
    getselectLibraryItems(url: string, listname: string): Promise<any>;
    getqdmsLibraryItems(url: string, listname: string): Promise<any>;
    getqdmsselectLibraryItems(url: string, listname: string): Promise<any>;
    getPathOfSelectedTemplate(fileName: string, listname: string): Promise<any>;
    getQDMSBuffer(siteUrl: string): Promise<any>;
}
//# sourceMappingURL=DMSService.d.ts.map