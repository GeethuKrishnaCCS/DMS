import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class DMSService extends BaseService {
    private _spfi: SPFI;
    private currentContext: WebPartContext;
    constructor(context: WebPartContext) {
        super(context);
        this.currentContext = context;
        this._spfi = getSP(this.currentContext);
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
    public createNewItem(siteUrl: string, listname: string, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(metadata);
    }
    public getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'DocumentIndex'")();
    }
    public getItemsFromUserMsgSettingsNP(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Title,Message")
            .orderBy("ID")
            .filter("PageName eq 'NotificationPreference'")();
    }
    public getItemsByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id)();
    }
    public getItemsFromDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")();
    }
    public updateItemById(siteUrl: string, listname: string, itemid: number, dataitem: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).update(dataitem);
    }
    public getNotificationPref(siteUrl: string, listname: string, mail: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("ID,Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail")
            .expand("EmailUser")
            .filter("EmailUser/EMail eq '" + mail + "'")();
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
}
