import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class cdmsEditService extends BaseService {
    private _spfi: SPFI;
    private currentContext: WebPartContext;
    // private spQdms: SPFI;
    constructor(context: WebPartContext, qdmsURL: string) {
        super(context);
        this.currentContext = context;
        this._spfi = getSP(this.currentContext);
        this._spfi = getSP(this.currentContext);
        // this.spQdms = new SPFI(qdmsURL).using(SPFx(context));
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
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .add(metadata);
    }
    public async getBuffer(siteUrl: string): Promise<any> {
        return this._spfi.web.getFileByServerRelativePath(siteUrl)
            .getBuffer()
    }
    public async uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any> {
        const file = await this._spfi.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });

        return file;
    }
    public getSelectLibraryItems(url: string, listname: string, select: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items
            .select(select)();
    }
    public getItemsFromLibrary(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname)
            .items();
    }
    public getByIdSelectExpand(siteUrl: string, listname: string, id: number, select: string, expand: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .select(select)
            .expand(expand)();
    }
    public getFilter(siteUrl: string, listname: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter(filter)();
    }
    public getByIdSelect(siteUrl: string, listname: string, id: number, select: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .select(select)();
    }
    public getSelectFilter(siteUrl: string, listname: string, select: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .select(select)
            .filter(filter)();
    }

    public getSelectExpand(siteUrl: string, listname: string, select: string, expand: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select(select)
            .expand(expand)();
    }




    public getItemsByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)();
    }
    public itemFromLibraryByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .getById(id)();
    }
    public async itemUpdate(siteUrl: string, listname: string, id: number, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .update(metadata);
    }
    public itemFromLibraryUpdate(siteUrl: string, listname: string, id: number, metadata): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .getById(id)
            .update(metadata);
    }

    public libraryByTitle(listname: string): Promise<any> {
        return this._spfi.web.lists.getByTitle(listname).select("Id")();
    }


    public getSelectFilterList(siteUrl: string, listname: string, select: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select(select)
            .filter(filter)();
    }









    /* public getItemsFromDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("ID,Title,Approver/Title,Approver/ID,Approver/EMail")
            .expand("Approver")();
    }    */

    /* public async itemFromTemplate(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.select("LinkFilename,ID").getAll();
    } */

    /* public async documnetPath(uniqueId: any): Promise<any> {
        return this._spfi.web.getFileById(uniqueId)
    } */


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





    /* public itemIDFromPublish(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .select("ID")
            .filter("DocumentIndex/ID eq '" + id + "'")();
    } */
    /* public getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Title,Message")
            .filter("PageName eq 'DocumentIndex'")();
    } */
    /*  public itemFromPrefernce(siteUrl: string, listname: string, emailUser: string): Promise<any> {
         return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
             .filter("EmailUser/EMail eq '" + emailUser + "'")
             .select("Preference")();
     } */



    /* public getSourceLink(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .select("SourceDocument")();
    } */
    /* public itemsFromIndex(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .select("DocumentStatus,SourceDocumentID,WorkflowStatus")();
    } */



    /* public DocumentSendForReview(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'Send For Review New DMS'")();
    } */
    /* public getQDMSPermissionWebpart(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_PermissionWebpart'")();
    } */
    /* public DocumentPermission(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPermission-Create Document'")();
    } */
    /* public DocumentPublish(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPublish'")();
    } */


    /* public itemsFromIndexExpanded(siteUrl: string, listname: string, id: number): Promise<any> {
        const items = "Title,Owner/Title,Owner/ID,Owner/EMail,SubCategoryID,WorkflowStatus,SourceDocument,SubCategory,Approver/Title,Approver/ID,ApprovedDate,BusinessUnit,BusinessUnitID,Category,CategoryID,DepartmentName,DepartmentID,DocumentID,DocumentName,ExpiryDate,Reviewers/ID,Reviewers/Title,ExpiryLeadPeriod,CategoryID,CriticalDocument,Template,PublishFormat,ApprovedDate,DirectPublish,CreateDocument,LegalEntity";
        const expand = "Owner,Approver,Reviewers";
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .select(items)
            .expand(expand)();
    } */



    /* public getqdmsLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname)
            .items();
    }
    public itemFromLibrary(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname)
            .items();
    } */

    /* public getqdmsselectLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items
            .select("LinkFilename,ID")();
    }
    public getselectLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items
            .select("LinkFilename,ID,Template,DocumentName,Category")();
    } */

}
