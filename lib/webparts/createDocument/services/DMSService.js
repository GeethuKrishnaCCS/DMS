var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { BaseService } from "./BaseService";
import { getSP } from "../shared/pnp/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
var DMSService = /** @class */ (function (_super) {
    __extends(DMSService, _super);
    function DMSService(context, qdmsURL) {
        var _this = _super.call(this, context) || this;
        _this.currentContext = context;
        _this._spfi = getSP(_this.currentContext);
        _this.spQdms = new SPFI(qdmsURL).using(SPFx(context));
        return _this;
    }
    DMSService.prototype.getItems = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items();
    };
    DMSService.prototype.getCurrentUser = function () {
        return this._spfi.web.currentUser();
    };
    DMSService.prototype.getUserIdByEmail = function (email) {
        return this._spfi.web.siteUsers.getByEmail(email)();
    };
    DMSService.prototype.createNewItem = function (siteUrl, listname, metadata) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.add(metadata);
    };
    DMSService.prototype.getItemsFromUserMsgSettings = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'DocumentIndex'")();
    };
    DMSService.prototype.getItemsByID = function (siteUrl, listname, id) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id)();
    };
    DMSService.prototype.getItemsFromDepartments = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")();
    };
    DMSService.prototype.uploadDocument = function (filename, filedata, libraryname) {
        return __awaiter(this, void 0, void 0, function () {
            var file;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._spfi.web.getFolderByServerRelativePath(libraryname)
                            .files.addUsingPath(filename, filedata, { Overwrite: true })];
                    case 1:
                        file = _a.sent();
                        return [2 /*return*/, file];
                }
            });
        });
    };
    DMSService.prototype.itemUpdate = function (siteUrl, listname, id, metadata) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(id).update(metadata)];
            });
        });
    };
    DMSService.prototype.itemFromTemplate = function (siteUrl, listname) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, this._spfi.web.getList(siteUrl + "/" + listname).items.select("LinkFilename,ID").getAll()];
            });
        });
    };
    DMSService.prototype.getBuffer = function (siteUrl) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, this._spfi.web.getFileByServerRelativePath(siteUrl).getBuffer()];
            });
        });
    };
    DMSService.prototype.documnetPath = function (uniqueId) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, this._spfi.web.getFileById(uniqueId)];
            });
        });
    };
    DMSService.prototype.itemFromLibrary = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/" + listname).items();
    };
    DMSService.prototype.itemFromLibraryUpdate = function (siteUrl, listname, id, metadata) {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.getById(id).update(metadata);
    };
    DMSService.prototype.getQDMSPermissionWebpart = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_PermissionWebpart'")();
    };
    DMSService.prototype.DocumentPermission = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_DocumentPermission-Create Document'")();
    };
    DMSService.prototype.DocumentSendForReview = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'Send For Review New DMS'")();
    };
    DMSService.prototype.DocumentPublish = function (siteUrl, listname) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("Title eq 'QDMS_DocumentPublish'")();
    };
    DMSService.prototype.itemFromLibraryByID = function (siteUrl, listname, id) {
        return this._spfi.web.getList(siteUrl + "/" + listname).items.getById(id)();
    };
    DMSService.prototype.itemFromPrefernce = function (siteUrl, listname, emailUser) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.filter("EmailUser/EMail eq '" + emailUser + "'").select("Preference")();
    };
    //MS Graph service
    DMSService.prototype.sendMail = function (emailPostBody) {
        return this.currentContext.msGraphClientFactory
            .getClient("3")
            .then(function (client) {
            client
                .api('/me/sendMail')
                .post(emailPostBody);
        });
    };
    DMSService.prototype.getItemFromSIbFunction = function (siteUrl, listname, id) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,ID,SFDepartment/Title,SFDepartment/ID,Reviewer/EMail,Reviewer/Title,Reviewer/ID,Approvers/Title,Approvers/ID").expand("SFDepartment,Approvers,Reviewer").filter("SFDepartment/ID eq '" + id + "'")();
    };
    DMSService.prototype.getItemOnSelectCategory = function (siteUrl, listname, id) {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.select("Title,ID,SFDepartment/Title,SFDepartment/ID,Reviewer/EMail,Reviewer/Title,Reviewer/ID,Approvers/Title,Approvers/ID").expand("SFDepartment,Approvers,Reviewer").filter("ID eq '" + id + "'")();
    };
    DMSService.prototype.getselectLibraryItems = function (url, listname) {
        return this._spfi.web.getList(url + "/" + listname).items.select("LinkFilename,ID,Template,DocumentName")();
    };
    DMSService.prototype.getqdmsLibraryItems = function (url, listname) {
        return this.spQdms.web.getList(url + "/" + listname).items();
    };
    DMSService.prototype.getqdmsselectLibraryItems = function (url, listname) {
        return this.spQdms.web.getList(url + "/" + listname).items.select("LinkFilename,ID,FileLeafRef,DocumentName")();
    };
    DMSService.prototype.getPathOfSelectedTemplate = function (fileName, listname) {
        return this.spQdms.web.lists.getByTitle(listname).items.select("FileDirRef,FileLeafRef").filter("FileLeafRef eq '".concat(fileName, "'"))();
    };
    DMSService.prototype.getQDMSBuffer = function (siteUrl) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, this.spQdms.web.getFileByServerRelativePath(siteUrl).getBuffer()];
            });
        });
    };
    return DMSService;
}(BaseService));
export { DMSService };
//# sourceMappingURL=DMSService.js.map