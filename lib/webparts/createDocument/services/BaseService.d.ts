import { WebPartContext } from '@microsoft/sp-webpart-base';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export declare class BaseService {
    private _paplSP;
    constructor(context: WebPartContext);
    getListItems(listname: string): Promise<any>;
    getListItemsById(listname: string, id: number): Promise<any>;
    getHubListItems(listname: string): Promise<any>;
    createNewProcess(data: any, listname: string): Promise<any>;
    getNotificationPreference(listName: string, email: string): Promise<any>;
    getEmailNotificationListItems(listname: string, filter: string): Promise<any>;
}
//# sourceMappingURL=BaseService.d.ts.map