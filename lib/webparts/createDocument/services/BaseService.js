import * as Constant from "../shared/constants";
import { SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
var BaseService = /** @class */ (function () {
    function BaseService(context) {
        this._paplSP = new SPFI(Constant.hubsiteurl).using(SPFx(context));
    }
    BaseService.prototype.getListItems = function (listname) {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items();
    };
    BaseService.prototype.getListItemsById = function (listname, id) {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items.select("ID,Code,Reviewers/Title,Reviewers/EMail,Approver/Title,Approver/EMail").expand("Reviewers,Approver").filter("ID eq '" + id + "'")();
    };
    BaseService.prototype.getHubListItems = function (listname) {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items();
    };
    BaseService.prototype.createNewProcess = function (data, listname) {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname)
            .items.add(data);
    };
    BaseService.prototype.getNotificationPreference = function (listName, email) {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listName).items.filter("EmailUser/EMail eq'" + email + "'")();
    };
    BaseService.prototype.getEmailNotificationListItems = function (listname, filter) {
        return this._paplSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items.filter("Title eq '" + filter + "'")();
    };
    return BaseService;
}());
export { BaseService };
//# sourceMappingURL=BaseService.js.map