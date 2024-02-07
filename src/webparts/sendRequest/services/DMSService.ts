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
    public getWorkflowStatus(siteUrl: string, documentIndexList: string, documentIndexID: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + documentIndexList).items
            .getById(documentIndexID)
            .select("WorkflowStatus,SourceDocument,DocumentStatus")();
    }
    public createNewItem(siteUrl: string, listname: string, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(metadata);

    }
    public getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'SendRequest'")();
    }
    public getItemsByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id)();
    }
    public getDocumentIndexItem(siteUrl: string, listname: string, documentIndexID: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(documentIndexID)
            .select("DocumentID,DocumentName,DepartmentID,BusinessUnitID,Owner/ID,Owner/Title,Owner/EMail,Approver/ID,Approver/Title,Approver/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID,Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Owner,Approver,Reviewers")();

    }
    public getSourceDocumentItem(siteUrl: string, listname: string, documentIndexID: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .filter('DocumentIndexId eq ' + documentIndexID)();

    }
    public getBusinessUnit(siteUrl: string, businessUnitList: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + businessUnitList).items
            .select("ID,Title,Approver/Title,Approver/ID,Approver/EMail")
            .expand("Approver")();
    }
    public getSiteUserByEmail(EMail: string): Promise<any> {
        return this._spfi.web.siteUsers.getByEmail(EMail)();
    }
    public getDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("ID,Title,Approver/Title,Approver/ID,Approver/EMail")
            .expand("Approver")();
    }
    public getPreviousHeaderItems(siteUrl: string, workflowHeaderList: string, documentIndexID: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + workflowHeaderList).items
            .select("ID").filter("DocumentIndex eq '" + documentIndexID + "' and(WorkflowStatus eq 'Returned with comments')")();
    }
    public addToWorkflowHeaderList(siteUrl: string, listname: string, itemtobeadded: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(itemtobeadded);
    }
    public addToDocumentRevision(siteUrl: string, listname: string, revisionitem: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(revisionitem);
    }
    public getSiteUserById(itemid: number): Promise<any> {
        return this._spfi.web.siteUsers.getById(itemid)();
    }
    public getTaskDelegation(siteUrl: string, taskDelegationSettings: string, userId: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + taskDelegationSettings).items
            .select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate")
            .expand("DelegatedFor,DelegatedTo")
            .filter("DelegatedFor/ID eq '" + userId + "' and(Status eq 'Active')")();
    }
    public addToWorkflowDetail(siteUrl: string, workflowDetailsList: string, item: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + workflowDetailsList).items.add(item);
    }
    public updateWorkflowDetailsList(siteUrl: string, listname: string, itemid: number, item: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).update(item);
    }
    public addToWorkflowTasksList(siteUrl: string, listname: string, item: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(item);
    }
    public updateItemById(siteUrl: string, listname: string, itemid: number, dataitem: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).update(dataitem);
    }
    public getUnderReview(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_DocumentPermission_UnderReview'")();
    }
    public getUnderApproval(siteUrl: string, requestList: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + requestList).items.filter("Title eq 'QDMS_DocumentPermission_UnderApproval'")();
    }
    public getEmailNotification(siteUrl: string, listname: string, type: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq '" + type + "'")();
    }
    public getMailPreference(siteUrl: string, listname: string, emailuser: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference")();
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
