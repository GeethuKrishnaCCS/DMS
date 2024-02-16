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
    public getQDMS_SendReviewWF(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("AccessGroups,AccessFields")
            .filter("Title eq 'QDMS_SendReviewWF'")();
    }
    public getBusinessDepartment(siteUrl: string, listname: string, itemid: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid)
            .select("DepartmentID,BusinessUnitID")();
    }
    public getWorkflowReviewDCCReview(siteUrl: string, listname: string, headerid: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,SourceDocument,Workflow")
            .expand("Responsible")
            .filter("HeaderID eq '" + headerid + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review') ")()
    }

    public getWFDetailWithResponsible(siteUrl: string, listname: string, detailID: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,TaskID")
            .expand("Responsible").filter("ID eq '" + detailID + "'")()
    }
    public getDetailWorkflowReview(siteUrl: string, listname: string, headerId: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
            .expand("Responsible")
            .filter("HeaderID eq '" + headerId + "' and (Workflow eq 'Review') ")()
    }
    public getHeaderItemsDocumentController(siteUrl: string, listname: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,Owner/Title,Owner/ID,Owner/EMail,Revision,WorkflowStatus,Title,DocumentIndexID,RequesterComment,RequestedDate,DueDate,PreviousReviewHeader,DocumentID,SourceDocumentID,DocumentController/Title,DocumentController/EMail,DocumentController/Id,DCCCompletionDate,Workflow")
            .expand("Owner,Approver,Requester,DocumentController")()
    }
    public getWFHOwnerApproverRequester(siteUrl: string, listname: string, headerId: any) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(headerId)
            .select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,Owner/Title,Owner/ID,Owner/EMail,Revision,WorkflowStatus,Title,DocumentIndexID,RequesterComment,RequestedDate,DueDate,PreviousReviewHeader,DocumentID,SourceDocumentID,Workflow")
            .expand("Owner,Approver,Requester")()
    }
    public getQDMS_DocumentPermission_UnderApproval(siteUrl: string, listname: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPermission_UnderApproval'")()
    }
    public getQDMS_DocumentPermission_UnderReview(siteUrl: string, listname: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPermission_UnderReview'")();
    }
    public getDCCReviewUnderReview(siteUrl: string, listname: string, headerId: string, documentIndexId: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("WorkflowID eq '" + headerId + "' and (DocumentIndexId eq '" + documentIndexId + "') and (Workflow eq 'DCC Review') and (Status eq 'Under Review')")()
    }
    public getWorkflowReviewUnderReview(siteUrl: string, listname: string, headerId: string, documentIndexId: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("WorkflowID eq '" + headerId + "' and (DocumentIndexId eq '" + documentIndexId + "') and (Workflow eq 'Review') and (Status eq 'Under Review')")()
    }
    public getResponsibleWithWFReview(siteUrl: string, listname: string, headerid: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate")
            .expand("Responsible")
            .filter("HeaderID eq '" + headerid + "' and (Workflow eq 'Review')  ")();
    }
    public getWorkflowDCCReview(siteUrl: string, listname: string, headerId: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate,ResponsibleComment,ResponseDate").expand("Responsible").filter("HeaderID eq '" + headerId + "' and (Workflow eq 'DCC Review')  ")()
    }
    public getWFDetailResponseStatus(siteUrl: string, listname: string, headerId: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("ResponseStatus")
            .filter("HeaderID eq " + headerId + " and (Workflow eq 'Review')")()
    }
    public getDelegateAndActive(siteUrl: string, listname: string, userid: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate")
            .expand("DelegatedFor,DelegatedTo")
            .filter("DelegatedFor/ID eq '" + userid + "' and(Status eq 'Active')")();
    }
    public getReviewersData(siteUrl: string, listname: string, headerId: number) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Reviewers/ID,Reviewers/Title,Reviewers/EMail")
            .expand("Reviewers").getById(headerId)();
    }
    public getEmailUserandPreference(siteUrl: string, listname: string, email: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail")
            .expand("EmailUser")
            .filter("EmailUser/EMail eq '" + email + "'")();
    }
    public getResponseStatusNeUnderReview(siteUrl: string, listname: string, headerId: string) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
            .expand("Responsible")
            .filter("HeaderID eq '" + headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review') and (ResponseStatus ne 'Under Review') ")()
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
