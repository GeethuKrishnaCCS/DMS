import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class cdmsEditService extends BaseService {
    private _spfi: SPFI;
    private currentContext: WebPartContext;
    private spQdms: SPFI;
    constructor(context: WebPartContext, qdmsURL: string) {
        super(context);
        this.currentContext = context;
        this._spfi = getSP(this.currentContext);
        this._spfi = getSP(this.currentContext);
        this.spQdms = new SPFI(qdmsURL).using(SPFx(context));
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
    public getItemsByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id)();
    }
    public getItemsFromDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")();
    }
    public async uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any> {
        const file = await this._spfi.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });

        return file;
    }
    public async itemUpdate(siteUrl: string, listname: string, id: number, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id).update(metadata);
    }
    /* public async itemFromTemplate(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.select("LinkFilename,ID").getAll();
    } */
    public async getBuffer(siteUrl: string): Promise<any> {
        return this._spfi.web.getFileByServerRelativePath(siteUrl).getBuffer()
    }
    /* public async documnetPath(uniqueId: any): Promise<any> {
        return this._spfi.web.getFileById(uniqueId)
    } */
    public itemFromLibrary(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items();
    }
    public itemFromLibraryUpdate(siteUrl: string, listname: string, id: number, metadata): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.getById(id).update(metadata);
    }
    public getQDMSPermissionWebpart(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_PermissionWebpart'")();
    }
    public DocumentPermission(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_DocumentPermission-Create Document'")();
    }
    public DocumentPublish(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_DocumentPublish'")();
    }
    public itemFromLibraryByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.getById(id)();
    }
    public itemFromPrefernce(siteUrl: string, listname: string, emailUser: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("EmailUser/EMail eq '" + emailUser + "'").select("Preference")();
    }
    public libraryByTitle(listname: string): Promise<any> {
        return this._spfi.web.lists.getByTitle(listname).select("Id")();
    }
    //MS Graph service
    /* public sendMail(emailPostBody: any): Promise<any> {
        return this.currentContext.msGraphClientFactory
            .getClient("3")
            .then((client: MSGraphClientV3): void => {
                client
                    .api('/me/sendMail')
                    .post(emailPostBody);
            });
    } */

    /* public getItemFromSIbFunction(siteUrl: string, listname: string, id: number,): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,ID,SFDepartment/Title,SFDepartment/ID,Reviewer/EMail,Reviewer/Title,Reviewer/ID,Approvers/Title,Approvers/ID").expand("SFDepartment,Approvers,Reviewer").filter("SFDepartment/ID eq '" + id + "'")();
    } */
    /* public getItemOnSelectCategory(siteUrl: string, listname: string, id: number,): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,ID,SFDepartment/Title,SFDepartment/ID,Reviewer/EMail,Reviewer/Title,Reviewer/ID,Approvers/Title,Approvers/ID").expand("SFDepartment,Approvers,Reviewer").filter("ID eq '" + id + "'")();
    } */

    public getselectLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items.select("LinkFilename,ID,Template,DocumentName,Category")();
    }
    public itemsFromIndex(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id).select("DocumentStatus,SourceDocumentID,WorkflowStatus")();
    }

    public itemsFromIndexExpanded(siteUrl: string, listname: string, id: number): Promise<any> {
        const items = "Title,Owner/Title,Owner/ID,Owner/EMail,SubCategoryID,WorkflowStatus,SourceDocument,SubCategory,Approver/Title,Approver/ID,ApprovedDate,BusinessUnit,BusinessUnitID,Category,CategoryID,DepartmentName,DepartmentID,DocumentID,DocumentName,ExpiryDate,Reviewers/ID,Reviewers/Title,ExpiryLeadPeriod,CategoryID,CriticalDocument,Template,PublishFormat,ApprovedDate,DirectPublish,CreateDocument,LegalEntity";
        const expand = "Owner,Approver,Reviewers";
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id).select(items).expand(expand)();
    }

    public itemIDFromPublish(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.select("ID").filter("DocumentIndex/ID eq '" + id + "'")();
    }
    public getSourceLink(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id).select("SourceDocument")();
    }
    public DocumentSendForReview(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'Send For Review New DMS'")();
    }
    public getqdmsLibraryItems(url: string, listname: string): Promise<any> {
        return this.spQdms.web.getList(url + "/" + listname).items();
    }
    public getqdmsselectLibraryItems(url: string, listname: string): Promise<any> {
        return this.spQdms.web.getList(url + "/" + listname).items.select("LinkFilename,ID")();
    }
}
