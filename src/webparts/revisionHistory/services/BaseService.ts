import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as Constant from "../shared/constants";
import { SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export class BaseService {
    private _paplSP: SPFI;

    constructor(context: WebPartContext,) {
        this._paplSP = new SPFI(Constant.hubsiteurl).using(SPFx(context));
    }


    public getListItems(listname: string): Promise<any> {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items();
    }
    public getListItemsById(listname: string, id: number): Promise<any> {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items.select("ID,Code,Reviewers/Title,Reviewers/EMail,Approver/Title,Approver/EMail").expand("Reviewers,Approver").filter("ID eq '" + id + "'")();
    }
    public getHubListItems(listname: string): Promise<any> {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items();
    }
    public createNewProcess(data: any, listname: string): Promise<any> {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname)
            .items.add(data);
    }

    public getNotificationPreference(listName: string, email: string): Promise<any> {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listName).items.filter("EmailUser/EMail eq'" + email + "'")()
    }
    public getEmailNotificationListItems(listname: string, filter: string): Promise<any> {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items.filter("Title eq '" + filter + "'")();
    }
} 