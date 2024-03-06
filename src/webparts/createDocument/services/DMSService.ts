import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/pnp/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class DMSService extends BaseService {
    private _spfi: SPFI;
    private currentContext: WebPartContext;
    // private spQdms: SPFI;
    constructor(context: WebPartContext, qdmsURL: string) {
        super(context);
        this.currentContext = context;
        this._spfi = getSP(this.currentContext);
        // this.spQdms = new SPFI(qdmsURL).using(SPFx(context));
    }
    public getCurrentUser(): Promise<any> {
        return this._spfi.web.currentUser();
    }
    public getUserIdByEmail(email: string): Promise<any> {
        return this._spfi.web.siteUsers.getByEmail(email)();
    }
    public async getBuffer(siteUrl: string): Promise<any> {
        return this._spfi.web.getFileByServerRelativePath(siteUrl).getBuffer()
    }
    public async uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any> {
        const file = await this._spfi.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });
            console.log('file: ', file);
        return file;
    }
    public getSelectLibraryItems(url: string, listname: string, select: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items
            .select(select)();
    }
    public getItemsFilter(siteUrl: string, listname: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter(filter)();
    }
    public getSelectFilter(siteUrl: string, listname: string, select: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select(select)
            .filter(filter)();
    }
    public async getByIdUpdate(siteUrl: string, listname: string, id: number, data: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .update(data);
    }
    public async getByIdLibraryUpdate(siteUrl: string, listname: string, id: number, data: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .getById(id)
            .update(data);
    }
    public getSelectExpand(siteUrl: string, listname: string, select: string, expand: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select(select)
            .expand(expand)();
    }
    public getItemFromLibrary(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname)
            .items();
    }









    public getItems(siteUrl: string, listname: string,): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items();
    }
    public createNewItem(siteUrl: string, listname: string, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .add(metadata);

    }


    /* public getqdmsLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname)
        .items();
    } */
    /* public itemFromLibrary(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname)
        .items();
    } */



    public getItemsByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)();
    }
    public itemFromLibraryByID(siteUrl: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .getById(id)();
    }


    /* public getItemsFromDepartments(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("ID,Title,Approver/Title,Approver/ID,Approver/EMail")
            .expand("Approver")();
    }
 */



    /* public async itemFromTemplate(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.select("LinkFilename,ID").getAll();
    } */

    /* public async documnetPath(uniqueId: any): Promise<any> {
        return this._spfi.web.getFileById(uniqueId)
    } */


    //MS Graph service
    /*  public sendMail(emailPostBody: any): Promise<any> {
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

    /* public async itemUpdate(siteUrl: string, listname: string, id: number, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .getById(id)
            .update(metadata);
    } */
    /* public itemFromLibraryUpdate(siteUrl: string, listname: string, id: number, metadata: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/" + listname).items
            .getById(id)
            .update(metadata);
    } */



    /* public getItemsFromUserMsgSettings(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select("Title,Message")
            .filter("PageName eq 'DocumentIndex'")();
    } */
    /* public itemFromPrefernce(siteUrl: string, listname: string, emailUser: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("EmailUser/EMail eq '" + emailUser + "'")
            .select("Preference")();
    } */
    public getPathOfSelectedTemplate(fileName: string, listname: string): Promise<any> {
        return this._spfi.web.lists.getByTitle(listname).items
            .select("FileDirRef,FileLeafRef")
            .filter(`FileLeafRef eq '${fileName}'`)()
    }



    /* public getQDMSPermissionWebpart(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_PermissionWebpart'")();
    } */
    /* public DocumentPermission(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPermission-Create Document'")();
    } */
    /* public DocumentSendForReview(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'Send For Review New DMS'")();
    } */
    /* public DocumentPublish(siteUrl: string, listname: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .filter("Title eq 'QDMS_DocumentPublish'")();
    } */




    /* public getqdmsselectLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items
            .select("LinkFilename,ID,FileLeafRef,DocumentName")();
    } */
    /* public getselectLibraryItems(url: string, listname: string): Promise<any> {
        return this._spfi.web.getList(url + "/" + listname).items
            .select("LinkFilename,ID,Template,DocumentName")();
    } */

    /* public async getQDMSBuffer(siteUrl: string): Promise<any> {
        return this.spQdms.web.getFileByServerRelativePath(siteUrl).getBuffer()
    } */
}
