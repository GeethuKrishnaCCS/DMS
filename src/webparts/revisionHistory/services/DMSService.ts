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
    public getItemById(siteUrl: string, listname: string, itemid: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid)();
    }
    public createNewItem(siteUrl: string, listname: string, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(metadata);
    }
    public updateItemById(siteUrl: string, listname: string, itemid: number, dataitem: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).update(dataitem);
    }
    public deleteItemById(siteUrl: string, listname: string, itemid: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).delete();
    }
    public getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'RevisionHistory'")();
    }
    /* public getItemsFromDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")();
    } */
    public async uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any> {
        const file = await this._spfi.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });
        return file;
    }
    public getProject_CancelWF(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("AccessGroups,AccessFields")
            .filter("Title eq 'Project_CancelWF'")();
    }
    public getProject_SendReminderWFTasks(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("AccessGroups,AccessFields")
            .filter("Title eq 'Project_SendReminderWFTasks'")();
    }
    public getProject_DelegateWFTask(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("AccessGroups,AccessFields").filter("Title eq 'Project_DelegateWFTask'")();
    }
    public getQDMS_SendReminderWFTasks(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendReminderWFTasks'")();
    }
    public getQDMS_CancelWF(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_CancelWF'")();
    }
    public getQDMS_DelegateWFTask(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_DelegateWFTask'")();
    }
    public getDocumentIndexItem(siteUrl: string, listname: string, itemid: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).select("DepartmentID,BusinessUnitID")();
    }
    public getAccessGroupID(siteUrl: string, listname: string, AG: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq '" + AG + "'")();
    }
    public getQDMS_DocumentPermission_UnderApproval(siteUrl: string, listname: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_DocumentPermission_UnderApproval'")();
    }
    public getLogItems(siteUrl: string, listname: string, itemid: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Title,Status,Modified,Created,Author/ID,Author/Title,Editor/ID,Editor/Title,LogDate,WorkflowID,Revision,DocumentIndex/ID,DocumentIndex/Title,DueDate,Workflow,ID")
            .expand("Author,Editor,DocumentIndex")
            .filter("DocumentIndex eq '" + itemid + "'")
            .getAll(5000);
    }
    public getIndexItemsWithOwnerApprover(siteUrl: string, listname: string, itemid: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Owner/Title,Owner/ID,Owner/EMail,DocumentName,SourceDocumentID,CriticalDocument,Revision,DocumentID,Approver/Title,Approver/ID")
            .expand("Owner,Approver")
            .filter("ID eq '" + itemid + "'")();
    }
    public getQDMS_DocumentPermission_UnderReview(siteUrl: string, listname: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPermission_UnderReview'")();
    }
    public getQDMS_DocumentPermission_Delegate(siteUrl: string, listname: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPermission_Delegate'")();
    }
    public getDocumentRevisionLog(siteUrl: string, listname: string, documentIndexID: string, ID: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("WorkflowID")
            .filter("DocumentIndex eq '" + documentIndexID + "'and (ID eq '" + ID + "')")();
    }
    public getWorkflowHeaderWithApproverRequester(siteUrl: string, listname: string, WorkflowID: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(WorkflowID)
            .select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,RequestedDate,DueDate")
            .expand("Approver,Requester")()
    }
    public getWorkflowHeaderItem(siteUrl: string, listname: string, WorkflowID: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(WorkflowID)
            .select("Requester/ID,Requester/Title,Approver/ID,Approver/Title,Reviewers/ID,Reviewers/Title,RequestedDate,DueDate")
            .expand("Approver,Requester,Reviewers")()
    }
    public getFlowDataInDocumentRevisionLog(siteUrl: string, listname: string, documentIndexID: string, ID: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("WorkflowID,ID,DueDate")
            .filter("DocumentIndex eq '" + documentIndexID + "' and (ID eq '" + ID + "')")();
    }
    public getDetailsWorkflow_DCCReview(siteUrl: string, listname: string, WorkflowID: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Workflow").expand("Responsible,Editor")
            .filter("HeaderID eq '" + WorkflowID + "' and (Workflow eq 'DCC Review') ")();
    }
    public getDetailsWorkflow_Review(siteUrl: string, listname: string, WorkflowID: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Workflow").expand("Responsible,Editor")
            .filter("HeaderID eq '" + WorkflowID + "' and (Workflow eq 'Review') ")();
    }
    public getDetailsWorkflow_Void(siteUrl: string, listname: string, WorkflowID: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Workflow").expand("Responsible,Editor")
            .filter("HeaderID eq '" + WorkflowID + "' and (Workflow eq 'Void') ")();
    }
    public getWorkflowApproval(siteUrl: string, listname: string, WorkflowID: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Link,Workflow")
            .expand("Responsible,Editor")
            .filter("HeaderID eq '" + WorkflowID + "' and (Workflow eq 'Approval') ")();
    }
    public getReviewersResponseStatus(siteUrl: string, listname: string, headerId: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("ResponseStatus")
            .filter("HeaderID eq " + headerId + " and (Workflow eq 'Review')")();
    }
    public validateUpdateListItem(siteUrl: string, listname: string, itemid: number, arrayData: any[]): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.getById(itemid).validateUpdateListItem(arrayData);
    }
    public getTaskDelegationData(siteUrl: string, listname: string, itemid: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate")
            .expand("DelegatedFor,DelegatedTo")
            .filter("DelegatedFor/ID eq '" + itemid + "'")();
    }
    public getNotificationPref(siteUrl: string, listname: string, email: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail")
            .expand("EmailUser")
            .filter("EmailUser/EMail eq '" + email + "'")();
    }
    public getEmailNoficationSettings(siteUrl: string, listname: string, type: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq '" + type + "'")();
    }
    public getRevisionHistory(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("PageName eq 'RevisionHistory'")();
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
