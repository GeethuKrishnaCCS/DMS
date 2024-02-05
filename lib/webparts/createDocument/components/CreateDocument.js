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
import * as React from 'react';
import styles from './CreateDocument.module.scss';
import * as moment from 'moment';
import { HttpClient } from '@microsoft/sp-http';
import SimpleReactValidator from 'simple-react-validator';
import * as _ from 'lodash';
import replaceString from 'replace-string';
import { DMSService } from '../services';
import { Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, getTheme, IconButton, Label, mergeStyleSets, MessageBar, Modal, PrimaryButton, Spinner, TextField, TooltipHost } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
var CreateDocument = /** @class */ (function (_super) {
    __extends(CreateDocument, _super);
    function CreateDocument(props) {
        var _this = _super.call(this, props) || this;
        _this.getSelectedReviewers = [];
        _this.Timeout = 5000;
        // Validator
        _this.componentWillMount = function () {
            _this.validator = new SimpleReactValidator({
                messages: { required: "This field is mandatory" }
            });
        };
        //Title Change
        _this._titleChange = function (ev, title) {
            _this.setState({ title: title || '', saveDisable: false });
        };
        //Owner Change
        _this._selectedOwner = function (items) {
            var ownerEmail;
            var ownerName;
            var getSelectedOwner = [];
            for (var item in items) {
                ownerEmail = items[item].secondaryText,
                    ownerName = items[item].text,
                    getSelectedOwner.push(items[item].id);
            }
            _this.setState({ owner: getSelectedOwner[0], ownerEmail: ownerEmail, ownerName: ownerName, saveDisable: false });
        };
        //Reviewer Change
        _this._selectedReviewers = function (items) {
            _this.getSelectedReviewers = [];
            for (var item in items) {
                _this.getSelectedReviewers.push(items[item].id);
            }
            _this.setState({ reviewers: _this.getSelectedReviewers });
        };
        //Approver Change
        _this._selectedApprover = function (items) { return __awaiter(_this, void 0, void 0, function () {
            var approverEmail, approverName, getSelectedApprover, departments, i, deptapprove;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        getSelectedApprover = [];
                        this.setState({ validApprover: "", approver: null, approverEmail: "", approverName: "", });
                        if (!(this.state.businessUnitCode != "")) return [3 /*break*/, 1];
                        return [3 /*break*/, 6];
                    case 1: return [4 /*yield*/, this._Service.getItemsFromDepartments(this.props.siteUrl, this.props.department)];
                    case 2:
                        departments = _a.sent();
                        i = 0;
                        _a.label = 3;
                    case 3:
                        if (!(i < departments.length)) return [3 /*break*/, 6];
                        if (!(departments[i].ID == this.state.departmentId)) return [3 /*break*/, 5];
                        return [4 /*yield*/, this._Service.getUserIdByEmail(departments[i].Approver.EMail)];
                    case 4:
                        deptapprove = _a.sent();
                        approverEmail = departments[i].Approver.EMail;
                        approverName = departments[i].Approver.Title;
                        getSelectedApprover.push(deptapprove.Id);
                        _a.label = 5;
                    case 5:
                        i++;
                        return [3 /*break*/, 3];
                    case 6:
                        this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
                        setTimeout(function () {
                            _this.setState({ validApprover: "none" });
                        }, 5000);
                        return [2 /*return*/];
                }
            });
        }); };
        //Create Document Change
        _this._onCreateDocChecked = function (ev, isChecked) { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                if (isChecked) {
                    this.setState({
                        hideDoc: "",
                        createDocument: true,
                        hideDirect: ""
                    });
                }
                else if (!isChecked) {
                    if (this.state.upload == true) {
                        this.myfile.value = "";
                    }
                    this.setState({ hideDirect: "", checkdirect: "none", insertdocument: "none", hideDoc: "", createDocument: false, hidePublish: "none", directPublishCheck: false });
                }
                return [2 /*return*/];
            });
        }); };
        _this.onUploadOrTemplateRadioBtnChange = function (ev, option) { return __awaiter(_this, void 0, void 0, function () {
            var publishedDocumentArray, sorted_PublishedDocument, publishedDocument, i, publishedDocumentdata;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({
                            uploadOrTemplateRadioBtn: option.key,
                            createDocument: true
                        });
                        if (option.key == "Upload") {
                            this.setState({ upload: true, hideupload: "", template: false, hidesource: "none", hidetemplate: "none" });
                        }
                        if (!(option.key == "Template")) return [3 /*break*/, 2];
                        publishedDocumentArray = [];
                        sorted_PublishedDocument = void 0;
                        this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
                        return [4 /*yield*/, this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary)];
                    case 1:
                        publishedDocument = _a.sent();
                        for (i = 0; i < publishedDocument.length; i++) {
                            if (publishedDocument[i].Template === true && publishedDocument[i].Category === this.state.category) {
                                publishedDocumentdata = {
                                    key: publishedDocument[i].ID,
                                    text: publishedDocument[i].DocumentName,
                                };
                                publishedDocumentArray.push(publishedDocumentdata);
                            }
                        }
                        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
                        this.setState({ templateDocuments: sorted_PublishedDocument });
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        }); };
        _this._onUploadCheck = function (ev, isChecked) {
            if (isChecked) {
                _this.setState({ upload: true, hideupload: "", template: false, hidesource: "none", hidetemplate: "none" });
            }
            else if (!isChecked) {
                // this.myfile.value = "";
                _this.setState({ upload: false, hideupload: "none" });
            }
        };
        _this._onTemplateCheck = function (ev, isChecked) { return __awaiter(_this, void 0, void 0, function () {
            var publishedDocumentArray, sorted_PublishedDocument, publishedDocument, i, publishedDocumentdata;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        publishedDocumentArray = [];
                        if (!isChecked) return [3 /*break*/, 2];
                        this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
                        return [4 /*yield*/, this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary)];
                    case 1:
                        publishedDocument = _a.sent();
                        for (i = 0; i < publishedDocument.length; i++) {
                            if (publishedDocument[i].Template === true) {
                                publishedDocumentdata = {
                                    key: publishedDocument[i].ID,
                                    text: publishedDocument[i].DocumentName,
                                };
                                publishedDocumentArray.push(publishedDocumentdata);
                            }
                        }
                        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
                        this.setState({ templateDocuments: sorted_PublishedDocument });
                        return [3 /*break*/, 3];
                    case 2:
                        if (!isChecked) {
                            this.setState({ template: false, hidesource: "none", hidetemplate: "none" });
                        }
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        //Direct Publish change
        _this._onDirectPublishChecked = function (ev, isChecked) {
            if (isChecked) {
                // this.setState({ checkdirect: "", });
                // this._checkdirectPublish('QDMS_DirectPublish');
                _this.setState({ hidePublish: "", directPublishCheck: true, approvalDate: new Date() });
            }
            else if (!isChecked) {
                _this.setState({ hidePublish: "none", directPublishCheck: false, approvalDate: new Date(), publishOption: "" });
            }
        };
        //Approval Date Change
        _this._onApprovalDatePickerChange = function (date) {
            _this.setState({
                approvalDate: date,
                approvalDateEdit: date
            });
        };
        //Expiry Change
        _this._onExpiryDateChecked = function (ev, isChecked) {
            if (isChecked) {
                _this.setState({ hideExpiry: "", expiryCheck: true, dateValid: "" });
            }
            else if (!isChecked) {
                _this.setState({ hideExpiry: "", expiryCheck: false, expiryDate: null, expiryLeadPeriod: "" });
            }
        };
        //Expiry Date Change
        _this._onExpDatePickerChange = function (date) {
            _this.setState({ expiryDate: date });
        };
        //Expiry Lead Period Change
        _this._expLeadPeriodChange = function (ev, expiryLeadPeriod) {
            var LeadPeriodformat = /^[0-9]*$/;
            if (expiryLeadPeriod.match(LeadPeriodformat)) {
                if (Number(expiryLeadPeriod) < 101) {
                    _this.setState({ expiryLeadPeriod: expiryLeadPeriod || '', leadmsg: "none" });
                }
                else {
                    _this.setState({ leadmsg: "" });
                }
            }
            else {
                _this.setState({ leadmsg: "" });
            }
        };
        //Critical Document Change
        _this._onCriticalChecked = function (ev, isChecked) {
            if (isChecked) {
                _this.setState({ criticalDocument: true });
            }
            else if (!isChecked) {
                _this.setState({ criticalDocument: false });
            }
        };
        // Template Change
        _this._onTemplateChecked = function (ev, isChecked) {
            if (isChecked) {
                _this.setState({ templateDocument: true });
            }
            else if (!isChecked) {
                _this.setState({ templateDocument: false });
            }
        };
        // qdms revision
        _this._revisionCoding = function () { return __awaiter(_this, void 0, void 0, function () {
            var revision, rev;
            return __generator(this, function (_a) {
                revision = parseInt("0");
                rev = revision + 1;
                this.setState({ newRevision: rev.toString() });
                return [2 /*return*/];
            });
        }); };
        //Send Mail
        _this._sendMail = function (emailuser, type, name) { return __awaiter(_this, void 0, void 0, function () {
            var formatday, day, mailSend, Subject, Body, link, notificationPreference, emailNotification, k, linkValue, replacedSubject, replaceRequester, replaceDate, replaceApprover, replaceBody, replacelink, FinalBody, emailPostBody_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        formatday = moment(this.today).format('DD/MMM/YYYY');
                        day = formatday.toString();
                        mailSend = "No";
                        console.log(this.state.criticalDocument);
                        return [4 /*yield*/, this._Service.itemFromPrefernce(this.props.siteUrl, this.props.notificationPreference, emailuser)];
                    case 1:
                        notificationPreference = _a.sent();
                        console.log(notificationPreference[0].Preference);
                        if (notificationPreference.length > 0) {
                            if (notificationPreference[0].Preference === "Send all emails") {
                                mailSend = "Yes";
                            }
                            else if (notificationPreference[0].Preference === "Send mail for critical document" && this.state.criticalDocument === true) {
                                mailSend = "Yes";
                            }
                            else {
                                mailSend = "No";
                            }
                        }
                        else if (this.state.criticalDocument === true) {
                            mailSend = "Yes";
                        }
                        if (!(mailSend === "Yes")) return [3 /*break*/, 3];
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.emailNotification)];
                    case 2:
                        emailNotification = _a.sent();
                        console.log(emailNotification);
                        for (k in emailNotification) {
                            if (emailNotification[k].Title === type) {
                                Subject = emailNotification[k].Subject;
                                Body = emailNotification[k].Body;
                            }
                        }
                        linkValue = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2";
                        // link = `<a href=${window.location.protocol + "//" + window.location.hostname+this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"}>`+this.state.documentName+`</a>`;
                        link = "<a href=".concat(linkValue, ">") + this.state.documentName + "</a>";
                        replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
                        replaceRequester = replaceString(Body, '[Sir/Madam]', name);
                        replaceDate = replaceString(replaceRequester, '[PublishedDate]', day);
                        replaceApprover = replaceString(replaceDate, '[Approver]', this.state.approverName);
                        replaceBody = replaceString(replaceApprover, '[DocumentName]', this.state.documentName);
                        replacelink = replaceString(replaceBody, '[DocumentLink]', link);
                        FinalBody = replacelink;
                        emailPostBody_1 = {
                            "message": {
                                "subject": replacedSubject,
                                "body": {
                                    "contentType": "HTML",
                                    "content": FinalBody
                                },
                                "toRecipients": [
                                    {
                                        "emailAddress": {
                                            "address": emailuser
                                        }
                                    }
                                ],
                            }
                        };
                        //Send Email uisng MS Graph  
                        this.props.context.msGraphClientFactory
                            .getClient("3")
                            .then(function (client) {
                            client
                                .api('/me/sendMail')
                                .post(emailPostBody_1);
                        });
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        //Cancel Document
        _this._onCancel = function () {
            _this.setState({
                cancelConfirmMsg: "",
                confirmDialog: false,
            });
        };
        //Cancel confirm
        _this._confirmYesCancel = function () {
            _this.setState({
                cancelConfirmMsg: "none",
                confirmDialog: true,
            });
            _this.validator.hideMessages();
            window.location.replace(_this.siteUrl);
        };
        //Not Cancel
        _this._confirmNoCancel = function () {
            _this.setState({
                cancelConfirmMsg: "none",
                confirmDialog: true,
            });
        };
        //For dialog box of cancel
        _this._dialogCloseButton = function () {
            _this.setState({
                cancelConfirmMsg: "none",
                confirmDialog: true,
            });
        };
        _this.dialogStyles = { main: { maxWidth: 500 } };
        _this.dialogContentProps = {
            type: DialogType.normal,
            closeButtonAriaLabel: 'none',
            title: 'Do you want to cancel?',
        };
        _this.modalProps = {
            isBlocking: true,
        };
        // On format date
        _this._onFormatDate = function (date) {
            console.log(moment(date).format("DD/MM/YYYY"));
            var selectd = moment(date).format("DD/MM/YYYY");
            return selectd;
        };
        _this._DueDateChange = function (date) {
            _this.setState({ DueDate: date, dueDateMadatory: "" });
        };
        _this._commentChange = function (ev, comments) {
            _this.setState({ comments: comments || '', });
        };
        _this.onConfirmReview = function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(this.state.DueDate !== null)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.setState({
                                sendForReview: true,
                                showReviewModal: false,
                                dueDateMadatory: "",
                                saveDisable: true, hideCreateLoading: " ",
                                norefresh: " "
                            })];
                    case 1:
                        _a.sent();
                        this._documentidgeneration();
                        return [3 /*break*/, 3];
                    case 2:
                        this.setState({ dueDateMadatory: "Yes" });
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this.state = {
            statusMessage: {
                isShowMessage: false,
                message: "",
                messageType: 90000,
            },
            title: "",
            approvalDate: "",
            loaderDisplay: "",
            legalEntityOption: [],
            businessUnitOption: [],
            departmentOption: [],
            categoryOption: [],
            owner: "",
            ownerEmail: "",
            ownerName: "",
            documentName: "",
            saveDisable: false,
            businessUnitID: null,
            departmentId: null,
            categoryId: null,
            businessUnit: "",
            businessUnitCode: "",
            departmentCode: "",
            department: "",
            subCategoryArray: [],
            subCategoryId: null,
            category: "",
            subCategory: "",
            categoryCode: "",
            legalEntityId: null,
            legalEntity: "",
            approver: null,
            approverEmail: "",
            approverName: "",
            reviewers: [],
            validApprover: "none",
            hideDoc: "",
            createDocument: false,
            hideDirect: "none",
            upload: false,
            checkdirect: "none",
            insertdocument: "none",
            hidePublish: "none",
            directPublishCheck: false,
            hideupload: "none",
            template: false,
            hidesource: "none",
            hidetemplate: "none",
            templateDocuments: "",
            isdocx: "none",
            nodocx: "",
            sourceId: "",
            templateId: "",
            templateKey: "",
            approvalDateEdit: new Date(),
            publishOption: "",
            hideExpiry: "",
            expiryCheck: false,
            expiryDate: null,
            expiryLeadPeriod: "",
            leadmsg: "none",
            criticalDocument: true,
            templateDocument: false,
            hideLoading: true,
            hideCreateLoading: "none",
            norefresh: "none",
            cancelConfirmMsg: "none",
            confirmDialog: true,
            hideloader: true,
            documentid: "",
            incrementSequenceNumber: "",
            sourceDocumentId: "",
            newDocumentId: "",
            newRevision: "",
            messageBar: "none",
            dateValid: "none",
            uploadOrTemplateRadioBtn: "",
            showReviewModal: false,
            DueDate: new Date(),
            sendForReview: false,
            dueDateMadatory: "",
            comments: ""
        };
        _this._Service = new DMSService(_this.props.context, window.location.protocol + "//" + window.location.hostname + "/" + _this.props.QDMSUrl);
        _this._bindData = _this._bindData.bind(_this);
        _this._departmentChange = _this._departmentChange.bind(_this);
        _this._categoryChange = _this._categoryChange.bind(_this);
        _this._subCategoryChange = _this._subCategoryChange.bind(_this);
        _this._selectedOwner = _this._selectedOwner.bind(_this);
        _this._selectedReviewers = _this._selectedReviewers.bind(_this);
        _this._selectedApprover = _this._selectedApprover.bind(_this);
        _this._onCreateDocChecked = _this._onCreateDocChecked.bind(_this);
        _this._sourcechange = _this._sourcechange.bind(_this);
        _this._templatechange = _this._templatechange.bind(_this);
        _this._onDirectPublishChecked = _this._onDirectPublishChecked.bind(_this);
        _this._onApprovalDatePickerChange = _this._onApprovalDatePickerChange.bind(_this);
        _this._publishOptionChange = _this._publishOptionChange.bind(_this);
        _this._onExpiryDateChecked = _this._onExpiryDateChecked.bind(_this);
        _this._onExpDatePickerChange = _this._onExpDatePickerChange.bind(_this);
        _this._expLeadPeriodChange = _this._expLeadPeriodChange.bind(_this);
        _this._onCriticalChecked = _this._onCriticalChecked.bind(_this);
        _this._onTemplateChecked = _this._onTemplateChecked.bind(_this);
        _this._onCreateDocument = _this._onCreateDocument.bind(_this);
        _this._documentidgeneration = _this._documentidgeneration.bind(_this);
        _this._incrementSequenceNumber = _this._incrementSequenceNumber.bind(_this);
        _this._documentCreation = _this._documentCreation.bind(_this);
        _this._addSourceDocument = _this._addSourceDocument.bind(_this);
        _this._createDocumentIndex = _this._createDocumentIndex.bind(_this);
        _this._revisionCoding = _this._revisionCoding.bind(_this);
        _this._legalEntityChange = _this._legalEntityChange.bind(_this);
        _this._add = _this._add.bind(_this);
        _this._checkdirectPublish = _this._checkdirectPublish.bind(_this);
        _this._onUploadCheck = _this._onUploadCheck.bind(_this);
        _this._onTemplateCheck = _this._onTemplateCheck.bind(_this);
        _this._onSendForReview = _this._onSendForReview.bind(_this);
        _this.onConfirmReview = _this.onConfirmReview.bind(_this);
        _this._dialogCloseButton = _this._dialogCloseButton.bind(_this);
        _this._closeModal = _this._closeModal.bind(_this);
        return _this;
    }
    // On load
    CreateDocument.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var user;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        //Huburl
                        this.siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
                        return [4 /*yield*/, this._Service.getCurrentUser()];
                    case 1:
                        user = _a.sent();
                        this.currentEmail = user.Email;
                        this.currentId = user.Id;
                        this.currentUser = user.Title;
                        //Get Today
                        this.today = new Date();
                        this.setState({ approvalDate: this.today });
                        this.setState({ loaderDisplay: "none" });
                        this._bindData();
                        this._checkdirectPublish('QDMS_DirectPublish');
                        return [2 /*return*/];
                }
            });
        });
    };
    //Bind dropdown in create
    CreateDocument.prototype._bindData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var businessUnitArray, sorted_BusinessUnit, departmentArray, sorted_Department, categoryArray, sorted_Category, legalEntityArray, sorted_LegalEntity, businessUnit, i, businessUnitdata, department, i, departmentdata, i, departmentdata, category, categorydata, i, legalEntity, i, legalEntityItemdata;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        businessUnitArray = [];
                        departmentArray = [];
                        categoryArray = [];
                        legalEntityArray = [];
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.businessUnit)];
                    case 1:
                        businessUnit = _a.sent();
                        for (i = 0; i < businessUnit.length; i++) {
                            businessUnitdata = {
                                key: businessUnit[i].ID,
                                text: businessUnit[i].BusinessUnitName,
                            };
                            businessUnitArray.push(businessUnitdata);
                        }
                        sorted_BusinessUnit = _.orderBy(businessUnitArray, 'text', ['asc']);
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.department)];
                    case 2:
                        department = _a.sent();
                        if (this.props.siteUrl === "/sites/Quality" || "/sites/PropertyManagement") {
                            for (i = 0; i < department.length; i++) {
                                departmentdata = {
                                    key: department[i].ID,
                                    text: department[i].Department,
                                };
                                departmentArray.push(departmentdata);
                            }
                        }
                        else {
                            for (i = 0; i < department.length; i++) {
                                if (this.props.siteUrl === "/sites/" + department[i].Title) {
                                    this.setState({
                                        departmentId: department[i].ID,
                                    });
                                    departmentdata = {
                                        key: department[i].ID,
                                        text: department[i].Department,
                                    };
                                    departmentArray.push(departmentdata);
                                    this._departmentChange(departmentdata);
                                }
                            }
                        }
                        sorted_Department = _.orderBy(departmentArray, 'text', ['asc']);
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.category)];
                    case 3:
                        category = _a.sent();
                        for (i = 0; i < category.length; i++) {
                            if (category[i].QDMS == true) {
                                categorydata = {
                                    key: category[i].ID,
                                    text: category[i].Category,
                                };
                                categoryArray.push(categorydata);
                            }
                        }
                        sorted_Category = _.orderBy(categoryArray, 'text', ['asc']);
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.legalEntity)];
                    case 4:
                        legalEntity = _a.sent();
                        for (i = 0; i < legalEntity.length; i++) {
                            legalEntityItemdata = {
                                key: legalEntity[i].ID,
                                text: legalEntity[i].Title
                            };
                            legalEntityArray.push(legalEntityItemdata);
                        }
                        sorted_LegalEntity = _.orderBy(legalEntityArray, 'text', ['asc']);
                        this.setState({
                            businessUnitOption: sorted_BusinessUnit,
                            departmentOption: sorted_Department,
                            categoryOption: sorted_Category,
                            legalEntityOption: sorted_LegalEntity,
                            owner: this.currentId,
                            ownerEmail: this.currentEmail,
                            ownerName: this.currentUser
                        });
                        this._userMessageSettings();
                        return [2 /*return*/];
                }
            });
        });
    };
    //Messages
    CreateDocument.prototype._userMessageSettings = function () {
        return __awaiter(this, void 0, void 0, function () {
            var userMessageSettings, i, successmsg, publishmsg;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._Service.getItemsFromUserMsgSettings(this.props.siteUrl, this.props.userMessageSettings)];
                    case 1:
                        userMessageSettings = _a.sent();
                        console.log(userMessageSettings);
                        for (i in userMessageSettings) {
                            if (userMessageSettings[i].Title == "CreateDocumentSuccess") {
                                successmsg = userMessageSettings[i].Message;
                                this.createDocument = replaceString(successmsg, '[DocumentName]', this.state.documentName);
                            }
                            if (userMessageSettings[i].Title == "DirectPublishSuccess") {
                                publishmsg = userMessageSettings[i].Message;
                                this.directPublish = replaceString(publishmsg, '[DocumentName]', this.state.documentName);
                            }
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    //Department Change
    CreateDocument.prototype._departmentChange = function (option) {
        return __awaiter(this, void 0, void 0, function () {
            var getApprover, approverEmail, approverName, department, departmentCode, departments, i, deptapprove;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        getApprover = [];
                        return [4 /*yield*/, this._Service.getItemsByID(this.props.siteUrl, this.props.department, option.key)];
                    case 1:
                        department = _a.sent();
                        departmentCode = department.Title;
                        this.setState({ departmentId: option.key, departmentCode: departmentCode, department: option.text, saveDisable: false });
                        if (!(this.state.businessUnitCode == "")) return [3 /*break*/, 7];
                        return [4 /*yield*/, this._Service.getItemsFromDepartments(this.props.siteUrl, this.props.department)];
                    case 2:
                        departments = _a.sent();
                        i = 0;
                        _a.label = 3;
                    case 3:
                        if (!(i < departments.length)) return [3 /*break*/, 6];
                        if (!(departments[i].ID == option.key)) return [3 /*break*/, 5];
                        return [4 /*yield*/, this._Service.getUserIdByEmail(departments[i].Approver.EMail)];
                    case 4:
                        deptapprove = _a.sent();
                        approverEmail = departments[i].Approver.EMail;
                        approverName = departments[i].Approver.Title;
                        getApprover.push(deptapprove.Id);
                        _a.label = 5;
                    case 5:
                        i++;
                        return [3 /*break*/, 3];
                    case 6:
                        this.setState({ approver: getApprover[0], approverEmail: approverEmail, approverName: approverName });
                        _a.label = 7;
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    //Category Change
    CreateDocument.prototype._categoryChange = function (option) {
        return __awaiter(this, void 0, void 0, function () {
            var subcategoryArray, sorted_subcategory, category, categoryCode, publishedDocumentArray, sorted_PublishedDocument, publishedDocument, i, publishedDocumentdata;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        subcategoryArray = [];
                        return [4 /*yield*/, this._Service.getItemsByID(this.props.siteUrl, this.props.category, option.key)];
                    case 1:
                        category = _a.sent();
                        categoryCode = category.Title;
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.subCategory).then(function (subcategory) {
                                for (var i = 0; i < subcategory.length; i++) {
                                    if (subcategory[i].CategoryId == option.key) {
                                        var subcategorydata = {
                                            key: subcategory[i].ID,
                                            text: subcategory[i].SubCategory,
                                        };
                                        subcategoryArray.push(subcategorydata);
                                    }
                                }
                                sorted_subcategory = _.orderBy(subcategoryArray, 'text', ['asc']);
                                _this.setState({
                                    categoryId: option.key,
                                    subCategoryArray: sorted_subcategory,
                                    category: option.text,
                                    categoryCode: categoryCode,
                                    saveDisable: false
                                });
                            })];
                    case 2:
                        _a.sent();
                        publishedDocumentArray = [];
                        return [4 /*yield*/, this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary)];
                    case 3:
                        publishedDocument = _a.sent();
                        for (i = 0; i < publishedDocument.length; i++) {
                            if (publishedDocument[i].Template === true && publishedDocument[i].Category === this.state.category) {
                                publishedDocumentdata = {
                                    key: publishedDocument[i].ID,
                                    text: publishedDocument[i].DocumentName,
                                };
                                publishedDocumentArray.push(publishedDocumentdata);
                            }
                        }
                        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
                        this.setState({ templateDocuments: sorted_PublishedDocument });
                        return [2 /*return*/];
                }
            });
        });
    };
    //SubCategory Change
    CreateDocument.prototype._subCategoryChange = function (option) {
        this.setState({ subCategoryId: option.key, subCategory: option.text });
    };
    // Legal Entity Change
    CreateDocument.prototype._legalEntityChange = function (option) {
        this.setState({ legalEntityId: option.key, legalEntity: option.text });
    };
    // On upload
    CreateDocument.prototype._add = function (e) {
        this.setState({ insertdocument: "none" });
        this.myfile = e.target.value;
        var type;
        var myfile;
        this.isDocument = "Yes";
        // @ts-ignore: Object is possibly 'null'.
        myfile = document.querySelector("#addqdms").files[0];
        console.log(myfile);
        this.isDocument = "Yes";
        var splitted = myfile.name.split(".");
        // let docsplit =splitted.slice(0, -1).join('.')+"."+splitted[splitted.length - 1];
        // alert(docsplit);
        type = splitted[splitted.length - 1];
        if (type === "docx") {
            this.setState({ isdocx: "", nodocx: "none" });
        }
        else {
            this.setState({ isdocx: "none", nodocx: "" });
        }
    };
    CreateDocument.prototype._sourcechange = function (option) {
        return __awaiter(this, void 0, void 0, function () {
            var publishedDocumentArray, sorted_PublishedDocument, publishedDocument, i, publishedDocumentdata, publishedDocument, i, publishedDocumentdata;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ hidetemplate: "", sourceId: option.key });
                        publishedDocumentArray = [];
                        if (!(option.key === "Quality")) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._Service.getqdmsLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary)];
                    case 1:
                        publishedDocument = _a.sent();
                        for (i = 0; i < publishedDocument.length; i++) {
                            if (publishedDocument[i].Template === true) {
                                publishedDocumentdata = {
                                    key: publishedDocument[i].ID,
                                    text: publishedDocument[i].DocumentName,
                                };
                                publishedDocumentArray.push(publishedDocumentdata);
                            }
                        }
                        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
                        this.setState({ templateDocuments: sorted_PublishedDocument, sourceId: option.key });
                        return [3 /*break*/, 4];
                    case 2: return [4 /*yield*/, this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary)];
                    case 3:
                        publishedDocument = _a.sent();
                        for (i = 0; i < publishedDocument.length; i++) {
                            if (publishedDocument[i].Template === true) {
                                publishedDocumentdata = {
                                    key: publishedDocument[i].ID,
                                    text: publishedDocument[i].DocumentName,
                                };
                                publishedDocumentArray.push(publishedDocumentdata);
                            }
                        }
                        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
                        this.setState({ templateDocuments: sorted_PublishedDocument, sourceId: option.key });
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    //Template change
    CreateDocument.prototype._templatechange = function (option) {
        return __awaiter(this, void 0, void 0, function () {
            var type, publishName;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ insertdocument: "none" });
                        this.setState({ templateId: option.key, templateKey: option.text });
                        this.isDocument = "Yes";
                        if (!(this.state.sourceId === "Quality")) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary).then(function (publishdoc) {
                                console.log(publishdoc);
                                for (var i = 0; i < publishdoc.length; i++) {
                                    if (publishdoc[i].Id === _this.state.templateId) {
                                        publishName = publishdoc[i].LinkFilename;
                                    }
                                }
                                var split = publishName.split(".", 2);
                                type = split[1];
                                if (type === "docx") {
                                    _this.setState({ isdocx: "", nodocx: "none" });
                                }
                                else {
                                    _this.setState({ isdocx: "none", nodocx: "" });
                                }
                            })];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 2: return [4 /*yield*/, this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary).then(function (publishdoc) {
                            console.log(publishdoc);
                            for (var i = 0; i < publishdoc.length; i++) {
                                if (publishdoc[i].Id === _this.state.templateId) {
                                    publishName = publishdoc[i].LinkFilename;
                                }
                            }
                            var split = publishName.split(".", 2);
                            type = split[1];
                            if (type === "docx") {
                                _this.setState({ isdocx: "", nodocx: "none" });
                            }
                            else {
                                _this.setState({ isdocx: "none", nodocx: "" });
                            }
                        })];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // Direct publish change
    CreateDocument.prototype._checkdirectPublish = function (type) {
        return __awaiter(this, void 0, void 0, function () {
            var laUrl, siteUrl, postURL, requestHeaders, body, postOptions, response, responseJSON;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._Service.getQDMSPermissionWebpart(this.props.siteUrl, this.props.requestList)];
                    case 1:
                        laUrl = _a.sent();
                        console.log("Posturl", laUrl[0].PostUrl);
                        this.permissionpostUrl = laUrl[0].PostUrl;
                        siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
                        postURL = this.permissionpostUrl;
                        requestHeaders = new Headers();
                        requestHeaders.append("Content-type", "application/json");
                        body = JSON.stringify({
                            'PermissionTitle': type,
                            'SiteUrl': siteUrl,
                            'CurrentUserEmail': this.currentEmail
                        });
                        postOptions = {
                            headers: requestHeaders,
                            body: body
                        };
                        return [4 /*yield*/, this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions)];
                    case 2:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 3:
                        responseJSON = _a.sent();
                        console.log(responseJSON);
                        if (response.ok) {
                            console.log(responseJSON['Status']);
                            if (responseJSON['Status'] === "Valid") {
                                if (this.props.directPublish === true) {
                                    this.setState({ checkdirect: "none", hideDirect: "", hidePublish: "none" });
                                }
                            }
                            else {
                                this.setState({ checkdirect: "none", hideDirect: "none", hidePublish: "none" });
                            }
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    //PublishOption Change
    CreateDocument.prototype._publishOptionChange = function (option) {
        this.setState({ publishOption: option.key, saveDisable: false });
    };
    //On create button click
    CreateDocument.prototype._onCreateDocument = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(this.state.createDocument === true && this.isDocument === "Yes" || this.state.createDocument === false)) return [3 /*break*/, 12];
                        if (!(this.state.expiryCheck === true)) return [3 /*break*/, 6];
                        if (!(this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver'))) return [3 /*break*/, 2];
                        this.setState({
                            saveDisable: true, hideCreateLoading: " ",
                            norefresh: " "
                        });
                        return [4 /*yield*/, this._documentidgeneration()];
                    case 1:
                        _a.sent();
                        this.validator.hideMessages();
                        return [3 /*break*/, 5];
                    case 2:
                        if (!(this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver'))) return [3 /*break*/, 4];
                        this.setState({
                            saveDisable: true, hideloader: false, hideCreateLoading: " ",
                            norefresh: " "
                        });
                        return [4 /*yield*/, this._documentidgeneration()];
                    case 3:
                        _a.sent();
                        this.validator.hideMessages();
                        return [3 /*break*/, 5];
                    case 4:
                        this.validator.showMessages();
                        this.forceUpdate();
                        _a.label = 5;
                    case 5: return [3 /*break*/, 11];
                    case 6:
                        if (!(this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver'))) return [3 /*break*/, 8];
                        this.setState({
                            saveDisable: true, hideCreateLoading: " ",
                            norefresh: " "
                        });
                        return [4 /*yield*/, this._documentidgeneration()];
                    case 7:
                        _a.sent();
                        this.validator.hideMessages();
                        return [3 /*break*/, 11];
                    case 8:
                        if (!(this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver'))) return [3 /*break*/, 10];
                        this.setState({
                            saveDisable: true, hideloader: false, hideCreateLoading: " ",
                            norefresh: " "
                        });
                        return [4 /*yield*/, this._documentidgeneration()];
                    case 9:
                        _a.sent();
                        this.validator.hideMessages();
                        return [3 /*break*/, 11];
                    case 10:
                        this.validator.showMessages();
                        this.forceUpdate();
                        _a.label = 11;
                    case 11: return [3 /*break*/, 13];
                    case 12:
                        this.setState({ insertdocument: "" });
                        _a.label = 13;
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    //Documentid generation
    CreateDocument.prototype._documentidgeneration = function () {
        return __awaiter(this, void 0, void 0, function () {
            var separator, sequenceNumber, idcode, counter, incrementstring, increment, documentid, isValue, settingsid, documentname, documentIdSettings, documentIdSequenceSettings, k, idsettings, addidseq, idItems, afterCounter;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        isValue = "false";
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.documentIdSettings)];
                    case 1:
                        documentIdSettings = _a.sent();
                        console.log("documentIdSettings", documentIdSettings);
                        separator = documentIdSettings[0].Separator;
                        sequenceNumber = documentIdSettings[0].SequenceDigit;
                        idcode = this.state.departmentCode + separator + this.state.categoryCode;
                        if (!documentIdSettings) return [3 /*break*/, 11];
                        return [4 /*yield*/, this._Service.getItems(this.props.siteUrl, this.props.documentIdSequenceSettings)];
                    case 2:
                        documentIdSequenceSettings = _a.sent();
                        console.log("documentIdSequenceSettings", documentIdSequenceSettings);
                        for (k in documentIdSequenceSettings) {
                            if (documentIdSequenceSettings[k].Title === idcode) {
                                counter = documentIdSequenceSettings[k].Sequence;
                                settingsid = documentIdSequenceSettings[k].ID;
                                isValue = "true";
                            }
                        }
                        if (!documentIdSequenceSettings) return [3 /*break*/, 11];
                        if (!(isValue === "false")) return [3 /*break*/, 7];
                        increment = 1;
                        incrementstring = increment.toString();
                        idsettings = {
                            Title: idcode,
                            Sequence: incrementstring
                        };
                        return [4 /*yield*/, this._Service.createNewItem(this.props.siteUrl, this.props.documentIdSequenceSettings, idsettings)];
                    case 3:
                        addidseq = _a.sent();
                        if (!addidseq) return [3 /*break*/, 6];
                        return [4 /*yield*/, this._incrementSequenceNumber(incrementstring, sequenceNumber)];
                    case 4:
                        _a.sent();
                        if (this.state.departmentCode != "") {
                            documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
                        }
                        else {
                            documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
                        }
                        documentname = documentid + " " + this.state.title;
                        this.setState({ documentid: documentid, documentName: documentname });
                        return [4 /*yield*/, this._documentCreation()];
                    case 5:
                        _a.sent();
                        _a.label = 6;
                    case 6: return [3 /*break*/, 11];
                    case 7:
                        increment = parseInt(counter) + 1;
                        incrementstring = increment.toString();
                        idItems = {
                            Title: idcode,
                            Sequence: incrementstring
                        };
                        return [4 /*yield*/, this._Service.itemUpdate(this.props.siteUrl, this.props.documentIdSequenceSettings, settingsid, idItems)];
                    case 8:
                        afterCounter = _a.sent();
                        if (!afterCounter) return [3 /*break*/, 11];
                        return [4 /*yield*/, this._incrementSequenceNumber(incrementstring, sequenceNumber)];
                    case 9:
                        _a.sent();
                        if (this.state.departmentCode != "") {
                            documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
                        }
                        else {
                            documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
                        }
                        documentname = documentid + " " + this.state.title;
                        this.setState({ documentid: documentid, documentName: documentname });
                        return [4 /*yield*/, this._documentCreation()];
                    case 10:
                        _a.sent();
                        _a.label = 11;
                    case 11: return [2 /*return*/];
                }
            });
        });
    };
    // Append sequence to the count
    CreateDocument.prototype._incrementSequenceNumber = function (incrementvalue, sequenceNumber) {
        var incrementSequenceNumber = incrementvalue;
        while (incrementSequenceNumber.length < sequenceNumber)
            incrementSequenceNumber = "0" + incrementSequenceNumber;
        console.log(incrementSequenceNumber);
        this.setState({
            incrementSequenceNumber: incrementSequenceNumber,
        });
    }; // Create item with id
    CreateDocument.prototype._documentCreation = function () {
        return __awaiter(this, void 0, void 0, function () {
            var documentNameExtension, sourceDocumentId, upload, docinsertname, myfile, splitted, fileUploaded, filePath, item, revision, logItems, indexItems, indexItems, publishName_1, extension_1, newDocumentName_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._userMessageSettings()];
                    case 1:
                        _a.sent();
                        upload = "#addqdms";
                        if (!(this.state.createDocument === true)) return [3 /*break*/, 17];
                        // Create document index item
                        return [4 /*yield*/, this._createDocumentIndex()];
                    case 2:
                        // Create document index item
                        _a.sent();
                        if (!(document.querySelector(upload).files[0] != null)) return [3 /*break*/, 15];
                        myfile = document.querySelector(upload).files[0];
                        console.log(myfile);
                        splitted = myfile.name.split(".");
                        documentNameExtension = this.state.documentName + '.' + splitted[splitted.length - 1];
                        this.documentNameExtension = documentNameExtension;
                        docinsertname = this.state.documentid + '.' + splitted[splitted.length - 1];
                        if (!myfile.size) return [3 /*break*/, 14];
                        return [4 /*yield*/, this._Service.uploadDocument(docinsertname, myfile, this.props.sourceDocumentLibrary)];
                    case 3:
                        fileUploaded = _a.sent();
                        if (!fileUploaded) return [3 /*break*/, 14];
                        filePath = window.location.protocol + "//" + window.location.host + fileUploaded.data.ServerRelativeUrl;
                        return [4 /*yield*/, fileUploaded.file.getItem()];
                    case 4:
                        item = _a.sent();
                        console.log(item);
                        sourceDocumentId = item["ID"];
                        this.setState({ sourceDocumentId: sourceDocumentId });
                        // update metadata
                        return [4 /*yield*/, this._addSourceDocument()];
                    case 5:
                        // update metadata
                        _a.sent();
                        if (!item) return [3 /*break*/, 14];
                        revision = void 0;
                        revision = "0";
                        logItems = {
                            Title: this.state.documentid,
                            Status: "Document Created",
                            LogDate: this.today,
                            Revision: revision,
                            DocumentIndexId: parseInt(this.state.newDocumentId),
                        };
                        return [4 /*yield*/, this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems)];
                    case 6:
                        _a.sent();
                        if (!(this.state.directPublishCheck === false)) return [3 /*break*/, 8];
                        indexItems = {
                            SourceDocumentID: parseInt(this.state.sourceDocumentId),
                            DocumentName: this.documentNameExtension,
                            SourceDocument: {
                                Description: this.documentNameExtension,
                                Url: (fileUploaded.data.LinkingUrl !== "") ? fileUploaded.data.LinkingUrl : filePath,
                            },
                            RevokeExpiry: {
                                Description: "Revoke",
                                Url: this.revokeUrl
                            },
                        };
                        return [4 /*yield*/, this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexItems)];
                    case 7:
                        _a.sent();
                        return [3 /*break*/, 10];
                    case 8:
                        indexItems = {
                            SourceDocumentID: parseInt(this.state.sourceDocumentId),
                            ApprovedDate: this.state.approvalDate,
                            DocumentName: this.documentNameExtension,
                            SourceDocument: {
                                Description: this.documentNameExtension,
                                Url: (fileUploaded.data.LinkingUrl !== "") ? fileUploaded.data.LinkingUrl : filePath,
                            },
                            RevokeExpiry: {
                                Description: "Revoke",
                                Url: this.revokeUrl
                            },
                        };
                        return [4 /*yield*/, this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexItems)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [4 /*yield*/, this._triggerPermission(sourceDocumentId)];
                    case 11:
                        _a.sent();
                        if (!(this.state.directPublishCheck === true)) return [3 /*break*/, 13];
                        this.setState({ hideLoading: false, hideCreateLoading: "none" });
                        return [4 /*yield*/, this._publish()];
                    case 12:
                        _a.sent();
                        return [3 /*break*/, 14];
                    case 13:
                        if (this.state.sendForReview === true) {
                            this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                            this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                            setTimeout(function () {
                                window.location.replace(_this.siteUrl);
                            }, 5000);
                        }
                        else {
                            this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                            setTimeout(function () {
                                window.location.replace(_this.siteUrl);
                            }, 5000);
                        }
                        _a.label = 14;
                    case 14: return [3 /*break*/, 16];
                    case 15:
                        if (this.state.templateId != "") {
                            // Get template
                            if (this.state.sourceId === "Quality") {
                                this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary)
                                    .then(function (publishdoc) {
                                    console.log(publishdoc);
                                    for (var i = 0; i < publishdoc.length; i++) {
                                        if (publishdoc[i].Id === _this.state.templateId) {
                                            publishName_1 = publishdoc[i].DocumentName;
                                        }
                                    }
                                    var split = publishName_1.split(".", 2);
                                    extension_1 = split[1];
                                }).then(function (cpysrc) {
                                    // Add template document to source document
                                    newDocumentName_1 = _this.state.documentName + "." + extension_1;
                                    _this.documentNameExtension = newDocumentName_1;
                                    docinsertname = _this.state.documentid + '.' + extension_1;
                                    var filePath;
                                    _this._Service.getPathOfSelectedTemplate(publishName_1, "SourceDocuments").then(function (items) {
                                        if (items.length > 0) {
                                            // Get the first item (assuming the file names are unique)
                                            var fileItem = items[0];
                                            // Access the server-relative URL of the file
                                            filePath = fileItem.FileDirRef + '/' + publishName_1;
                                            console.log(filePath);
                                        }
                                    }).then(function (afterPath) {
                                        _this._Service.getBuffer(filePath).then(function (templateData) {
                                            return _this._Service.uploadDocument(docinsertname, templateData, _this.props.sourceDocumentLibrary);
                                        }).then(function (fileUploaded) {
                                            var filePath = window.location.protocol + "//" + window.location.host + fileUploaded.data.ServerRelativeUrl;
                                            console.log("File Uploaded");
                                            fileUploaded.file.getItem().then(function (item) { return __awaiter(_this, void 0, void 0, function () {
                                                return __generator(this, function (_a) {
                                                    switch (_a.label) {
                                                        case 0:
                                                            console.log(item);
                                                            sourceDocumentId = item["ID"];
                                                            this.setState({ sourceDocumentId: sourceDocumentId });
                                                            return [4 /*yield*/, this._addSourceDocument()];
                                                        case 1:
                                                            _a.sent();
                                                            return [2 /*return*/];
                                                    }
                                                });
                                            }); }).then(function (updateDocumentIndex) { return __awaiter(_this, void 0, void 0, function () {
                                                var revision, logItems, indexUpdateItems, indexUpdateItems;
                                                var _this = this;
                                                return __generator(this, function (_a) {
                                                    switch (_a.label) {
                                                        case 0:
                                                            revision = "0";
                                                            logItems = {
                                                                Title: this.state.documentid,
                                                                Status: "Document Created",
                                                                LogDate: this.today,
                                                                Revision: revision,
                                                                DocumentIndexId: parseInt(this.state.newDocumentId),
                                                            };
                                                            return [4 /*yield*/, this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems)];
                                                        case 1:
                                                            _a.sent();
                                                            if (this.state.directPublishCheck === false) {
                                                                indexUpdateItems = {
                                                                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                                                                    DocumentName: this.documentNameExtension,
                                                                    SourceDocument: {
                                                                        Description: this.documentNameExtension,
                                                                        Url: (fileUploaded.data.LinkingUrl !== "") ? fileUploaded.data.LinkingUrl : filePath,
                                                                    },
                                                                    RevokeExpiry: {
                                                                        Description: "Revoke",
                                                                        Url: this.revokeUrl
                                                                    }
                                                                };
                                                                this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                                                            }
                                                            else {
                                                                indexUpdateItems = {
                                                                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                                                                    DocumentName: this.documentNameExtension,
                                                                    ApprovedDate: this.state.approvalDate,
                                                                    SourceDocument: {
                                                                        Description: this.documentNameExtension,
                                                                        Url: (fileUploaded.data.LinkingUrl !== "") ? fileUploaded.data.LinkingUrl : filePath,
                                                                    },
                                                                    RevokeExpiry: {
                                                                        Description: "Revoke",
                                                                        Url: this.revokeUrl
                                                                    },
                                                                };
                                                                this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                                                            }
                                                            return [4 /*yield*/, this._triggerPermission(sourceDocumentId)];
                                                        case 2:
                                                            _a.sent();
                                                            if (!(this.state.directPublishCheck === true)) return [3 /*break*/, 4];
                                                            this.setState({ hideLoading: false, hideCreateLoading: "none" });
                                                            return [4 /*yield*/, this._publish()];
                                                        case 3:
                                                            _a.sent();
                                                            return [3 /*break*/, 5];
                                                        case 4:
                                                            if (this.state.sendForReview === true) {
                                                                this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                                                                this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                                                                setTimeout(function () {
                                                                    window.location.replace(_this.siteUrl);
                                                                }, 5000);
                                                            }
                                                            else {
                                                                this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                                                                setTimeout(function () {
                                                                    window.location.replace(_this.siteUrl);
                                                                }, 5000);
                                                            }
                                                            _a.label = 5;
                                                        case 5: return [2 /*return*/];
                                                    }
                                                });
                                            }); });
                                        });
                                    });
                                    // let siteUrl = this.props.QDMSUrl + "/" + this.props.publisheddocumentLibrary + "/" + publishName;
                                });
                            }
                            else {
                                this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary)
                                    .then(function (publishdoc) {
                                    console.log(publishdoc);
                                    for (var i = 0; i < publishdoc.length; i++) {
                                        if (publishdoc[i].Id === _this.state.templateId) {
                                            publishName_1 = publishdoc[i].LinkFilename;
                                        }
                                    }
                                    var split = publishName_1.split(".", 2);
                                    extension_1 = split[1];
                                }).then(function (cpysrc) {
                                    // Add template document to source document
                                    newDocumentName_1 = _this.state.documentName + "." + extension_1;
                                    _this.documentNameExtension = newDocumentName_1;
                                    docinsertname = _this.state.documentid + '.' + extension_1;
                                    var siteUrl = _this.props.siteUrl + "/" + _this.props.publisheddocumentLibrary + "/" + _this.state.category + "/" + publishName_1;
                                    _this._Service.getBuffer(siteUrl).then(function (templateData) {
                                        return _this._Service.uploadDocument(docinsertname, templateData, _this.props.sourceDocumentLibrary);
                                    }).then(function (fileUploaded) {
                                        var filePath = window.location.protocol + "//" + window.location.host + fileUploaded.data.ServerRelativeUrl;
                                        console.log("File Uploaded");
                                        fileUploaded.file.getItem().then(function (item) { return __awaiter(_this, void 0, void 0, function () {
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        console.log(item);
                                                        sourceDocumentId = item["ID"];
                                                        this.setState({ sourceDocumentId: sourceDocumentId });
                                                        return [4 /*yield*/, this._addSourceDocument()];
                                                    case 1:
                                                        _a.sent();
                                                        return [2 /*return*/];
                                                }
                                            });
                                        }); }).then(function (updateDocumentIndex) { return __awaiter(_this, void 0, void 0, function () {
                                            var revision, logItems, indexUpdateItems, indexUpdateItems;
                                            var _this = this;
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        revision = "0";
                                                        logItems = {
                                                            Title: this.state.documentid,
                                                            Status: "Document Created",
                                                            LogDate: this.today,
                                                            Revision: revision,
                                                            DocumentIndexId: parseInt(this.state.newDocumentId),
                                                        };
                                                        return [4 /*yield*/, this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems)];
                                                    case 1:
                                                        _a.sent();
                                                        if (this.state.directPublishCheck === false) {
                                                            indexUpdateItems = {
                                                                SourceDocumentID: parseInt(this.state.sourceDocumentId),
                                                                DocumentName: this.documentNameExtension,
                                                                SourceDocument: {
                                                                    Description: this.documentNameExtension,
                                                                    Url: (fileUploaded.data.LinkingUrl !== "") ? fileUploaded.data.LinkingUrl : filePath,
                                                                },
                                                                RevokeExpiry: {
                                                                    Description: "Revoke",
                                                                    Url: this.revokeUrl
                                                                }
                                                            };
                                                            this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                                                        }
                                                        else {
                                                            indexUpdateItems = {
                                                                SourceDocumentID: parseInt(this.state.sourceDocumentId),
                                                                DocumentName: this.documentNameExtension,
                                                                ApprovedDate: this.state.approvalDate,
                                                                SourceDocument: {
                                                                    Description: this.documentNameExtension,
                                                                    Url: (fileUploaded.data.LinkingUrl !== "") ? fileUploaded.data.LinkingUrl : filePath,
                                                                },
                                                                RevokeExpiry: {
                                                                    Description: "Revoke",
                                                                    Url: this.revokeUrl
                                                                },
                                                            };
                                                            this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                                                        }
                                                        return [4 /*yield*/, this._triggerPermission(sourceDocumentId)];
                                                    case 2:
                                                        _a.sent();
                                                        if (!(this.state.directPublishCheck === true)) return [3 /*break*/, 4];
                                                        this.setState({ hideLoading: false, hideCreateLoading: "none" });
                                                        return [4 /*yield*/, this._publish()];
                                                    case 3:
                                                        _a.sent();
                                                        return [3 /*break*/, 5];
                                                    case 4:
                                                        if (this.state.sendForReview === true) {
                                                            this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                                                            this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                                                            setTimeout(function () {
                                                                window.location.replace(_this.siteUrl);
                                                            }, 5000);
                                                        }
                                                        else {
                                                            this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                                                            setTimeout(function () {
                                                                window.location.replace(_this.siteUrl);
                                                            }, 5000);
                                                        }
                                                        _a.label = 5;
                                                    case 5: return [2 /*return*/];
                                                }
                                            });
                                        }); });
                                    });
                                });
                            }
                        }
                        else { }
                        _a.label = 16;
                    case 16: return [3 /*break*/, 19];
                    case 17: return [4 /*yield*/, this._createDocumentIndex()];
                    case 18:
                        _a.sent();
                        this.setState({ statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 }, norefresh: "none", hideCreateLoading: "none" });
                        setTimeout(function () {
                            window.location.replace(_this.siteUrl);
                        }, this.Timeout);
                        _a.label = 19;
                    case 19: return [2 /*return*/];
                }
            });
        });
    };
    // Create Document Index
    CreateDocument.prototype._createDocumentIndex = function () {
        var _this = this;
        var documentIndexId;
        // Without Expiry date
        if (this.state.expiryCheck === false) {
            var indexItems = {
                Title: this.state.title,
                DocumentID: this.state.documentid,
                ReviewersId: this.state.reviewers,
                DocumentName: this.state.documentName,
                BusinessUnitID: this.state.businessUnitID,
                BusinessUnit: this.state.businessUnit,
                CategoryID: this.state.categoryId,
                Category: this.state.category,
                SubCategoryID: this.state.subCategoryId,
                SubCategory: this.state.subCategory,
                ApproverId: this.state.approver,
                Revision: "0",
                WorkflowStatus: this.state.sendForReview === true ? "Under Review" : "Draft",
                DocumentStatus: "Active",
                Template: this.state.templateDocument,
                CriticalDocument: this.state.criticalDocument,
                CreateDocument: this.state.createDocument,
                DirectPublish: this.state.directPublishCheck,
                OwnerId: this.state.owner,
                DepartmentName: this.state.department,
                DepartmentID: this.state.departmentId,
                PublishFormat: this.state.publishOption
            };
            this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, indexItems).then(function (newdocid) { return __awaiter(_this, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    console.log(newdocid);
                    this.documentIndexID = newdocid.data.ID;
                    documentIndexId = newdocid.data.ID;
                    this.setState({ newDocumentId: documentIndexId });
                    this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
                    this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
                    return [2 /*return*/];
                });
            }); });
        }
        // With Expiry date
        else {
            var indexItems = {
                Title: this.state.title,
                DocumentID: this.state.documentid,
                ReviewersId: this.state.reviewers,
                DocumentName: this.state.documentName,
                BusinessUnitID: this.state.businessUnitID,
                BusinessUnit: this.state.businessUnit,
                CategoryID: this.state.categoryId,
                Category: this.state.category,
                SubCategoryID: this.state.subCategoryId,
                SubCategory: this.state.subCategory,
                ApproverId: this.state.approver,
                ExpiryDate: this.state.expiryDate,
                DirectPublish: this.state.directPublishCheck,
                ExpiryLeadPeriod: this.state.expiryLeadPeriod,
                Revision: "0",
                WorkflowStatus: this.state.sendForReview === true ? "Under Review" : "Draft",
                DocumentStatus: "Active",
                Template: this.state.templateDocument,
                CriticalDocument: this.state.criticalDocument,
                CreateDocument: this.state.createDocument,
                OwnerId: this.state.owner,
                DepartmentName: this.state.department,
                DepartmentID: this.state.departmentId,
                PublishFormat: this.state.publishOption,
            };
            this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, indexItems).then(function (newdocid) { return __awaiter(_this, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    console.log(newdocid);
                    this.documentIndexID = newdocid.data.ID;
                    documentIndexId = newdocid.data.ID;
                    this.setState({ newDocumentId: documentIndexId });
                    this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
                    this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
                    return [2 /*return*/];
                });
            }); });
        }
    };
    // Add Source Document metadata
    CreateDocument.prototype._addSourceDocument = function () {
        return __awaiter(this, void 0, void 0, function () {
            var sourceUpdate, sourceUpdate;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(this.state.expiryCheck === false)) return [3 /*break*/, 2];
                        sourceUpdate = {
                            Title: this.state.title,
                            DocumentID: this.state.documentid,
                            ReviewersId: this.state.reviewers,
                            DocumentName: this.documentNameExtension,
                            BusinessUnit: this.state.businessUnit,
                            Category: this.state.category,
                            SubCategory: this.state.subCategory,
                            ApproverId: this.state.approver,
                            Revision: "0",
                            WorkflowStatus: this.state.sendForReview === true ? "Under Review" : "Draft",
                            DocumentStatus: "Active",
                            DocumentIndexId: parseInt(this.state.newDocumentId),
                            PublishFormat: this.state.publishOption,
                            CriticalDocument: this.state.criticalDocument,
                            Template: this.state.templateDocument,
                            OwnerId: this.state.owner,
                            DepartmentName: this.state.department,
                            RevisionHistory: {
                                Description: "Revision History",
                                Url: this.revisionHistoryUrl
                            }
                        };
                        return [4 /*yield*/, this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId, sourceUpdate)];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 2:
                        sourceUpdate = {
                            DocumentID: this.state.documentid,
                            Title: this.state.title,
                            ReviewersId: this.state.reviewers,
                            DocumentName: this.documentNameExtension,
                            BusinessUnit: this.state.businessUnit,
                            Category: this.state.category,
                            SubCategory: this.state.subCategory,
                            ApproverId: this.state.approver,
                            ExpiryDate: this.state.expiryDate,
                            ExpiryLeadPeriod: this.state.expiryLeadPeriod,
                            Revision: "0",
                            WorkflowStatus: this.state.sendForReview !== true ? "Draft" : "Under Review",
                            DocumentStatus: "Active",
                            CriticalDocument: this.state.criticalDocument,
                            DocumentIndexId: parseInt(this.state.newDocumentId),
                            PublishFormat: this.state.publishOption,
                            Template: this.state.templateDocument,
                            OwnerId: this.state.owner,
                            DepartmentName: this.state.department,
                            RevisionHistory: {
                                Description: "Revision History",
                                Url: this.revisionHistoryUrl
                            }
                        };
                        return [4 /*yield*/, this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId, sourceUpdate)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // Set permission for document
    CreateDocument.prototype._triggerPermission = function (sourceDocumentID) {
        return __awaiter(this, void 0, void 0, function () {
            var laUrl, siteUrl, postURL, requestHeaders, body, postOptions;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._Service.DocumentPermission(this.props.siteUrl, this.props.requestList)];
                    case 1:
                        laUrl = _a.sent();
                        console.log("Posturl", laUrl[0].PostUrl);
                        this.postUrl = laUrl[0].PostUrl;
                        siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
                        postURL = this.postUrl;
                        requestHeaders = new Headers();
                        requestHeaders.append("Content-type", "application/json");
                        body = JSON.stringify({
                            'SiteURL': siteUrl,
                            'ItemId': sourceDocumentID
                        });
                        postOptions = {
                            headers: requestHeaders,
                            body: body
                        };
                        return [4 /*yield*/, this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions)];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    //trigger sendForRevew
    CreateDocument.prototype._triggerSendForReview = function (sourceDocumentID, documentIndexId) {
        return __awaiter(this, void 0, void 0, function () {
            var laUrl, siteUrl, postURL, requestHeaders, body, postOptions;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._Service.DocumentSendForReview(this.props.siteUrl, this.props.requestList)];
                    case 1:
                        laUrl = _a.sent();
                        console.log("Posturl", laUrl[0].PostUrl);
                        this.postUrl = laUrl[0].PostUrl;
                        siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
                        postURL = this.postUrl;
                        requestHeaders = new Headers();
                        requestHeaders.append("Content-type", "application/json");
                        body = JSON.stringify({
                            'SiteURL': siteUrl,
                            'ItemId': sourceDocumentID,
                            'IndexId': documentIndexId,
                            'Title': this.state.title,
                            'DueDate': this.state.DueDate,
                            'Comments': this.state.comments,
                        });
                        postOptions = {
                            headers: requestHeaders,
                            body: body
                        };
                        return [4 /*yield*/, this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions)];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    //Document Published
    CreateDocument.prototype._publish = function () {
        return __awaiter(this, void 0, void 0, function () {
            var laUrl, siteUrl, postURL, requestHeaders, body, postOptions, response, responseJSON;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._revisionCoding()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this._Service.DocumentPublish(this.props.siteUrl, this.props.requestList)];
                    case 2:
                        laUrl = _a.sent();
                        console.log("Posturl", laUrl[0].PostUrl);
                        this.postUrl = laUrl[0].PostUrl;
                        siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
                        postURL = this.postUrl;
                        requestHeaders = new Headers();
                        requestHeaders.append("Content-type", "application/json");
                        body = JSON.stringify({
                            'Status': 'Published',
                            'PublishFormat': this.state.publishOption,
                            'SourceDocumentID': this.state.sourceDocumentId,
                            'SiteURL': siteUrl,
                            'PublishedDate': this.today,
                            'DocumentName': this.state.documentName,
                            'Revision': this.state.newRevision,
                            'SourceDocumentLibrary': this.props.sourceDocumentLibrary,
                            'WorkflowStatus': "Published",
                            'RevisionUrl': this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.state.newDocumentId,
                        });
                        postOptions = {
                            headers: requestHeaders,
                            body: body
                        };
                        return [4 /*yield*/, this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions)];
                    case 3:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 4:
                        responseJSON = _a.sent();
                        console.log(responseJSON);
                        if (response.ok) {
                            this._publishUpdate();
                        }
                        else { }
                        return [2 /*return*/];
                }
            });
        });
    };
    // Published Document Metadata update
    CreateDocument.prototype._publishUpdate = function () {
        return __awaiter(this, void 0, void 0, function () {
            var itemToUpdate, itemToLog;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._Service.itemFromLibraryByID(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId)];
                    case 1:
                        _a.sent();
                        itemToUpdate = {
                            PublishFormat: this.state.publishOption,
                            WorkflowStatus: "Published",
                            Revision: this.state.newRevision,
                            ApprovedDate: new Date()
                        };
                        return [4 /*yield*/, this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemToUpdate)];
                    case 2:
                        _a.sent();
                        if (this.state.owner != this.currentId) {
                            this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
                        }
                        itemToLog = {
                            Title: this.state.documentid,
                            Status: "Published",
                            LogDate: this.today,
                            Revision: this.state.newRevision,
                            DocumentIndexId: this.documentIndexID,
                        };
                        return [4 /*yield*/, this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, itemToLog)];
                    case 3:
                        _a.sent();
                        this.setState({ hideLoading: true, norefresh: "none", hideCreateLoading: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.directPublish, messageType: 4 } });
                        setTimeout(function () {
                            window.location.replace(_this.siteUrl);
                        }, 5000);
                        return [2 /*return*/];
                }
            });
        });
    };
    CreateDocument.prototype._closeModal = function () {
        this.setState({ showReviewModal: false });
    };
    CreateDocument.prototype._onSendForReview = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                if (this.state.createDocument === true && this.isDocument === "Yes" || this.state.createDocument === false) {
                    if (this.state.expiryCheck === true) {
                        //Validation without direct publish
                        if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
                            if (this.isDocument === "Yes") {
                                this.setState({
                                    showReviewModal: true,
                                });
                            }
                            else {
                                this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
                                setTimeout(function () {
                                    _this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
                                }, 5000);
                            }
                            this.validator.hideMessages();
                        }
                        //Validation with direct publish
                        else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
                            if (this.isDocument === "Yes") {
                                this.setState({
                                    showReviewModal: true,
                                });
                            }
                            else {
                                this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
                                setTimeout(function () {
                                    _this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
                                }, 5000);
                            }
                            this.validator.hideMessages();
                        }
                        else {
                            this.validator.showMessages();
                            this.forceUpdate();
                        }
                    }
                    else {
                        //Validation without direct publish
                        if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
                            if (this.isDocument === "Yes") {
                                this.setState({
                                    showReviewModal: true,
                                });
                            }
                            else {
                                this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
                                setTimeout(function () {
                                    _this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
                                }, 5000);
                            }
                            this.validator.hideMessages();
                        }
                        //Validation with direct publish
                        else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
                            if (this.isDocument === "Yes") {
                                this.setState({
                                    showReviewModal: true,
                                });
                            }
                            else {
                                this.setState({ statusMessage: { isShowMessage: true, message: "Please select document", messageType: 1 } });
                                setTimeout(function () {
                                    _this.setState({ statusMessage: { isShowMessage: false, message: "Please select document", messageType: 4 } });
                                }, 5000);
                            }
                            this.validator.hideMessages();
                        }
                        else {
                            this.validator.showMessages();
                            this.forceUpdate();
                        }
                    }
                }
                else {
                    this.setState({ insertdocument: "" });
                }
                return [2 /*return*/];
            });
        });
    };
    CreateDocument.prototype.render = function () {
        var _this = this;
        var publishOptions = [
            { key: 'PDF', text: 'PDF' },
            { key: 'Native', text: 'Native' },
        ];
        var publishOption = [
            { key: 'Native', text: 'Native' },
        ];
        var Source = [
            { key: 'Quality', text: 'Quality' },
            { key: 'Current Site', text: 'Current Site' }
        ];
        var calloutProps = { gapSpace: 0 };
        var hostStyles = { root: { display: 'inline-block' } };
        var uploadOrTemplateRadioBtnOptions = [
            { key: 'Upload', text: 'Upload existing file' },
            { key: 'Template', text: 'Create document using existing template', styles: { field: { marginLeft: "35px" } } },
        ];
        var choiceGroupStyles = { root: { display: 'flex' }, flexContainer: { display: "flex", justifyContent: 'space-between' } };
        var cancelIcon = { iconName: 'Cancel' };
        var theme = getTheme();
        var contentStyles = mergeStyleSets({
            container: {
                width: "40%",
                marginLeft: "8%",
                borderRadius: "12px"
            }
        });
        var iconButtonStyles = {
            root: {
                color: theme.palette.neutralPrimary,
                marginTop: '4px',
                marginRight: '4px',
                width: '25px',
                height: '25px',
                float: "right",
                cursor: "pointer"
            },
            rootHovered: {
                color: theme.palette.neutralDark,
            },
        };
        return (React.createElement("section", { className: "".concat(styles.createDocument) },
            React.createElement("div", { className: styles.border },
                React.createElement("div", { className: styles.alignCenter }, this.props.webpartHeader),
                React.createElement("div", null,
                    React.createElement(TextField, { required: true, id: "t1", label: "Title", onChange: this._titleChange, value: this.state.title }),
                    React.createElement("div", { style: { color: "#dc3545" } },
                        this.validator.message("Title", this.state.title, "required|alpha_num_dash_space|max:200"),
                        " ")),
                React.createElement("div", { className: styles.divrow },
                    React.createElement("div", { className: styles.divColumn1 },
                        React.createElement(Dropdown, { id: "t3", label: "Category", selectedKey: this.state.departmentId, placeholder: "Select an option", defaultSelectedKey: this.state.departmentId, required: true, 
                            //disabled={this.state.departmentId !== ""}
                            options: this.state.departmentOption, onChanged: this._departmentChange }),
                        React.createElement("div", { style: { color: "#dc3545", textAlign: "center" } },
                            this.validator.message("BU/Dep", this.state.businessUnitID || this.state.departmentId, "required"),
                            "")),
                    React.createElement("div", { className: styles.divColumn2 },
                        React.createElement(Dropdown, { id: "t2", required: true, label: "Doc Category", placeholder: "Select an option", selectedKey: this.state.categoryId, options: this.state.categoryOption, onChanged: this._categoryChange }),
                        React.createElement("div", { style: { color: "#dc3545" } },
                            this.validator.message("category", this.state.categoryId, "required"),
                            " ")),
                    React.createElement("div", { className: styles.divColumn2 },
                        React.createElement(Dropdown, { id: "t2", required: true, label: "Doc Type", placeholder: "Select an option", selectedKey: this.state.subCategoryId, options: this.state.subCategoryArray, onChanged: this._subCategoryChange }),
                        React.createElement("div", { style: { color: "#dc3545" } },
                            this.validator.message("subCategory", this.state.subCategoryId, "required"),
                            " "))),
                React.createElement("div", { className: styles.documentMainDiv },
                    React.createElement("div", { className: styles.radioDiv, style: { display: this.state.hideDoc } },
                        React.createElement(ChoiceGroup, { selectedKey: this.state.uploadOrTemplateRadioBtn, onChange: this.onUploadOrTemplateRadioBtnChange, options: uploadOrTemplateRadioBtnOptions, styles: choiceGroupStyles })),
                    React.createElement("div", { className: styles.uploadDiv, style: { display: this.state.hideupload } },
                        React.createElement("div", null,
                            React.createElement("input", { type: "file", name: "myFile", id: "addqdms", onChange: this._add })),
                        React.createElement("div", { style: { display: this.state.insertdocument, color: "#dc3545" } }, "Please select  document ")),
                    React.createElement("div", { className: styles.templateDiv, style: { display: this.state.hidetemplate } },
                        React.createElement("div", { className: styles.divColumn2, style: { display: "flex" } },
                            this.props.siteUrl !== "/sites/Quality" &&
                                React.createElement("div", { className: styles.divColumn2 },
                                    React.createElement(Dropdown, { id: "t7", label: "Source", placeholder: "Select an option", selectedKey: this.state.sourceId, options: Source, onChanged: this._sourcechange })),
                            React.createElement("div", { className: styles.divColumn2, style: { maxWidth: (this.props.siteUrl === "/sites/Quality") ? "26.8rem" : "163.8rem" } },
                                React.createElement(Dropdown, { id: "t7", label: "Select a Template", placeholder: "Select an option", selectedKey: this.state.templateId, options: this.state.templateDocuments, onChanged: this._templatechange, style: { width: "150%", } }))))),
                React.createElement("div", { className: styles.divrow },
                    React.createElement("div", { style: { width: "77%" } },
                        React.createElement(PeoplePicker, { context: this.props.context, titleText: "Owner", personSelectionLimit: 1, groupName: "", showtooltip: true, required: true, disabled: false, ensureUser: true, onChange: this._selectedOwner, defaultSelectedUsers: [this.props.context.pageContext.user.email], showHiddenInUI: false, principalTypes: [PrincipalType.User], resolveDelay: 1000 })),
                    React.createElement("div", { style: { width: "75%", marginLeft: "10px" } },
                        React.createElement(PeoplePicker, { context: this.props.context, titleText: "Reviewer(s)", personSelectionLimit: 10, groupName: "", showtooltip: true, required: false, disabled: false, ensureUser: true, showHiddenInUI: false, onChange: function (items) { return _this._selectedReviewers(items); }, principalTypes: [PrincipalType.User], resolveDelay: 1000, peoplePickerCntrlclassName: "testClass" })),
                    React.createElement("div", { className: styles.divApprover },
                        React.createElement(PeoplePicker, { context: this.props.context, titleText: "Approver", personSelectionLimit: 1, groupName: "", showtooltip: true, required: true, disabled: false, ensureUser: true, onChange: this._selectedApprover, showHiddenInUI: false, defaultSelectedUsers: [this.state.approverName], principalTypes: [PrincipalType.User], resolveDelay: 1000 }))),
                React.createElement("div", { className: styles.divrow },
                    React.createElement("div", { style: { width: "77%" } },
                        React.createElement("div", { style: { color: "#dc3545" } },
                            this.validator.message("Owner", this.state.owner, "required"),
                            " ")),
                    React.createElement("div", { style: { width: "75%", marginLeft: "10px" } },
                        React.createElement("div", { style: { color: "#dc3545" } })),
                    React.createElement("div", { className: styles.divApprover },
                        React.createElement("div", { style: { display: this.state.validApprover, color: "#dc3545" } }, "Not able to change approver"),
                        React.createElement("div", { style: { color: "#dc3545" } },
                            this.validator.message("Approver", this.state.approver, "required"),
                            " "))),
                React.createElement("div", { className: styles.divrow },
                    React.createElement("div", { className: styles.divDate, style: { display: this.state.hideExpiry } },
                        React.createElement(DatePicker, { label: "Expiry Date", value: this.state.expiryDate, onSelectDate: this._onExpDatePickerChange, placeholder: "Select a date...", ariaLabel: "Select a date", minDate: new Date(), formatDate: this._onFormatDate }),
                        React.createElement("div", { style: { display: this.state.dateValid } },
                            React.createElement("div", { style: { color: "#dc3545" } },
                                this.validator.message("expiryDate", this.state.expiryDate, "required"),
                                ""))),
                    React.createElement("div", { className: styles.wdthmid, style: { display: this.state.hideExpiry, width: "14.5%", } },
                        React.createElement(TextField, { id: "Expiry Reminder", name: "Expiry Reminder (Days)", label: "Expiry Reminder(Days)", onChange: this._expLeadPeriodChange, value: this.state.expiryLeadPeriod }),
                        React.createElement("div", { style: { display: this.state.dateValid } },
                            React.createElement("div", { style: { color: "#dc3545" } },
                                this.validator.message("ExpiryLeadPeriod", this.state.expiryLeadPeriod, "required"),
                                "")),
                        React.createElement("div", { style: { color: "#dc3545", display: this.state.leadmsg } }, "Enter only numbers less than 100")),
                    React.createElement("div", { style: { marginTop: "35px", marginLeft: "11px" } },
                        " ",
                        React.createElement(TooltipHost, { content: "Do you want to make this as a template?", 
                            //id={tooltipId}
                            calloutProps: calloutProps, styles: hostStyles },
                            React.createElement(Checkbox, { label: "Save as template ", boxSide: "start", onChange: this._onTemplateChecked, checked: this.state.templateDocument }))),
                    React.createElement("div", { style: { display: this.state.hideDirect, marginTop: "36px", marginLeft: "15px" } },
                        React.createElement(TooltipHost, { content: "Without review or approval, the document will be published.", 
                            //id={tooltipId}
                            calloutProps: calloutProps, styles: hostStyles },
                            React.createElement(Checkbox, { label: "Direct Publish?", boxSide: "start", onChange: this._onDirectPublishChecked, checked: this.state.directPublishCheck }))),
                    React.createElement("div", { style: { marginLeft: "31px", display: this.state.hidePublish } },
                        React.createElement(Dropdown, { id: "t2", required: true, label: "Publish Option", selectedKey: this.state.publishOption, placeholder: "Select an option", options: this.state.isdocx === "" ? publishOptions : publishOption, onChanged: this._publishOptionChange }),
                        React.createElement("div", { style: { color: "#dc3545" } },
                            this.validator.message("publish", this.state.publishOption, "required"),
                            ""))),
                React.createElement("div", { className: styles.divrow },
                    React.createElement("div", { className: styles.wdthmid, style: { display: this.state.checkdirect } },
                        React.createElement(Spinner, { label: 'Please Wait...' }))),
                React.createElement("div", null,
                    " ",
                    this.state.statusMessage.isShowMessage ?
                        React.createElement(MessageBar, { messageBarType: this.state.statusMessage.messageType, isMultiline: false, dismissButtonAriaLabel: "Close" }, this.state.statusMessage.message)
                        : '',
                    " "),
                React.createElement("div", { className: styles.mt },
                    React.createElement("div", { hidden: this.state.hideLoading },
                        React.createElement(Spinner, { label: 'Publishing...' }))),
                React.createElement("div", { className: styles.mt },
                    React.createElement("div", { style: { display: this.state.hideCreateLoading } },
                        React.createElement(Spinner, { label: 'Creating...' }))),
                React.createElement("div", { className: styles.mt },
                    React.createElement("div", { style: { display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" } },
                        React.createElement(Label, null, "***PLEASE DON'T REFRESH***"))),
                React.createElement("div", { className: styles.divrow },
                    React.createElement("div", { style: { fontStyle: "italic", fontSize: "12px", position: "absolute" } },
                        React.createElement("span", { style: { color: "red", fontSize: "23px" } }, "*"),
                        "fields are mandatory "),
                    React.createElement("div", { className: styles.rgtalign },
                        React.createElement(PrimaryButton, { id: "b2", className: styles.btn, disabled: this.state.saveDisable, onClick: this._onSendForReview }, "Send for review and submit"),
                        React.createElement(PrimaryButton, { id: "b2", className: styles.btn, disabled: this.state.saveDisable, onClick: this._onCreateDocument }, "Submit"),
                        React.createElement(PrimaryButton, { id: "b1", className: styles.btn, onClick: this._onCancel }, "Cancel"))),
                React.createElement("div", { style: { display: this.state.cancelConfirmMsg } },
                    React.createElement("div", null,
                        React.createElement(Dialog, { hidden: this.state.confirmDialog, dialogContentProps: this.dialogContentProps, onDismiss: this._dialogCloseButton, styles: this.dialogStyles, modalProps: this.modalProps },
                            React.createElement(DialogFooter, null,
                                React.createElement(PrimaryButton, { onClick: this._confirmYesCancel, text: "Yes" }),
                                React.createElement(DefaultButton, { onClick: this._confirmNoCancel, text: "No" }))))),
                React.createElement("div", { style: { padding: "18px" } },
                    React.createElement(Modal, { isOpen: this.state.showReviewModal, isModeless: true, containerClassName: contentStyles.container },
                        React.createElement("div", { style: { padding: "18px" } },
                            React.createElement("div", { className: styles.modalHeading, style: { display: "flex" } },
                                React.createElement("span", { style: { textAlign: "center", display: "flex", justifyContent: "center", flexGrow: "1" } },
                                    React.createElement("b", null, "Send For Review")),
                                React.createElement(IconButton, { iconProps: cancelIcon, ariaLabel: "Close popup modal", onClick: this._closeModal, styles: iconButtonStyles })),
                            React.createElement(DatePicker, { label: "Due Date *", value: this.state.DueDate, onSelectDate: this._DueDateChange, placeholder: "Select a date...", ariaLabel: "Select a date", minDate: new Date(), formatDate: this._onFormatDate }),
                            this.state.dueDateMadatory === "Yes" &&
                                React.createElement("label", { style: { color: 'Red' } }, "This field is mandatory"),
                            React.createElement(TextField, { id: "comments", autoComplete: 'true', label: "Comments", onChange: this._commentChange, value: this.state.comments, multiline: true }),
                            React.createElement(PrimaryButton, { style: { float: "right", marginTop: "7px", marginBottom: "9px" }, className: styles.modalButton, id: "b2", onClick: this.onConfirmReview }, "Confirm")))),
                React.createElement("br", null))));
    };
    return CreateDocument;
}(React.Component));
export default CreateDocument;
//# sourceMappingURL=CreateDocument.js.map