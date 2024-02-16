import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class DMSService extends BaseService {
    private _spfi: SPFI;
    private currentContext: WebPartContext;
    //private spQdms: SPFI;
    constructor(context: WebPartContext, qdmsURL?: string) {
        super(context);
        this.currentContext = context;
        this._spfi = getSP(this.currentContext);
        //this.spQdms = new SPFI(qdmsURL).using(SPFx(context));
    }
    public getItems(siteUrl: string, listname: string,): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items();
    }
    public getCurrentUser(): Promise<any> {
        return this._spfi.web.currentUser();
    }
    public getUserIdByEmail(email: string): Promise<any> {
        return this._spfi.web.siteUsers.getByEmail(email)();
    }
    public getUserById(reviewID: number): Promise<any> {
        return this._spfi.web.siteUsers.getById(reviewID)()
    }
    public getItemById(siteUrl: string, listname: string, itemid: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid)();
    }
    public getItemByIdSelect(siteUrl: string, listname: string, itemid: any, select: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).select(select)();
    }
    public createNewItem(siteUrl: string, listname: string, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(metadata);
    }
    public updateItemById(siteUrl: string, listname: string, itemid: number, dataitem: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).update(dataitem);
    }
    public getItemTitleFilter(siteUrl: string, listname: string, title: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq '" + title + "'")();
    }
    public deleteItemById(siteUrl: string, listname: string, itemid: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).delete();
    }
    public validateUpdateListItem(siteUrl: string, listname: string, itemid: number, arrayData: any[]): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.getById(itemid).validateUpdateListItem(arrayData);
    }
    public getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'DocumentIndex'")();
    }
    public getItemsFromDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")();
    }
    public async uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any> {
        const file = await this._spfi.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });
        return file;
    }
    public getDocumentIndexID(siteUrl: string, listname: string, headerId: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(headerId).select("DocumentIndexID")();
    }
    public getUserMessageForReview(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'Review'")();
    }
    public getApproverData(siteUrl: string, listname: string, headerid: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(headerid).select("Approver/ID,Approver/EMail,DocumentIndexID").expand("Approver")()
    }
    public getSelectFilter(siteUrl: string, listname: string, select: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname)
            .items
            .select(select)
            .filter(filter)()
    }
    public getItemFilter(siteUrl: string, listname: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname)
            .items
            .filter(filter)()
    }
    public getByIdSelect(siteUrl: string, listname: string, itemid: number, select: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname)
            .items
            .getById(itemid)
            .select(select)()
    }
    public getByIdSelectExpand(siteUrl: string, listname: string, itemid: number, select: string, expand: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname)
            .items
            .getById(itemid)
            .select(select)
            .expand(expand)()
    }
    public getByIdSelectFilterExpand(siteUrl: string, listname: string, select: string, filter: string, expand: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname)
            .items
            .select(select)
            .filter(filter)
            .expand(expand)()
    }
    //MS Graph service
    public sendMail(emailPostBody: any): Promise<any> {
        return this.currentContext.msGraphClientFactory
            .getClient("3")
            .then((client: MSGraphClientV3): void => {
                client
                    .api('/me/sendMail')
                    .post(emailPostBody);
            });
    }
    public getGroupMembers(groupId: string): Promise<any> {
        return this.currentContext.msGraphClientFactory
            .getClient("3")
            .then((client: MSGraphClientV3): void => {
                client
                    .api(`/groups/${groupId}/members`)
                    .version('v1.0')
                    .select(['mail', 'displayName'])
                    .get()
            });
    }
}
