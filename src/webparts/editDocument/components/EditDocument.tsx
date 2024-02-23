import * as React from 'react';
import styles from './EditDocument.module.scss';
import type { IEditDocumentProps, IEditDocumentState } from '../interfaces';
import { Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, getTheme, IChoiceGroupOption, IChoiceGroupStyles, IconButton, IDropdownOption, IIconProps, ITooltipHostStyles, Label, mergeStyleSets, MessageBar, Modal, Pivot, PivotItem, PrimaryButton, ProgressIndicator, Spinner, TextField, TooltipHost } from '@fluentui/react';
import { Web } from '@pnp/sp/webs';
import * as moment from 'moment';
import { MSGraphClientV3, HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import SimpleReactValidator from 'simple-react-validator';
import * as _ from 'lodash';
import replaceString from 'replace-string';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { cdmsEditService } from '../services';

const back: IIconProps = { iconName: 'ChromeBack' };

export default class EditDocument extends React.Component<IEditDocumentProps, IEditDocumentState> {
  private _Service: cdmsEditService;
  private validator: SimpleReactValidator;
  private currentEmail;
  private currentId;
  private getSelectedReviewers: any[] = [];
  private QDMSUrl = Web(window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl);
  private revisionHistoryUrl;
  private revokeUrl;
  private today;
  private directPublish;
  private editDocument;
  private sourceDocumentID;
  private sourceDocumentLibraryId;
  private documentIndexID;
  private documentNameExtension;
  private postUrl;
  private siteUrl;
  private indexUrl;
  private isDocument;
  private myfile;
  private permissionpostUrl;
  public constructor(props: IEditDocumentProps) {
    super(props);
    this.state = {
      hideProject: "none",
      hideDoc: "",
      hidePublish: "none",
      hideExpiry: "",
      projectEditDocumentView: "none",
      createDocumentProject: "none",
      createDocumentView: "",
      qdmsEditDocumentView: "none",
      revokeExpiryView: "none",
      hideCreate: "",
      messageBar: "none",
      hidebutton: "",
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      cancelConfirmMsg: "none",
      confirmDialog: true,
      accessDeniedMessageBar: "none",
      title: "",
      businessUnitID: null,
      departmentId: null,
      categoryId: null,
      dateValid: "none",
      uploadOrTemplateRadioBtn: "",
      subCategoryKey: "",
      legalEntityId: null,
      createDocument: false,
      replaceDocumentCheckbox: false,
      templateDocuments: "",
      directPublishCheck: false,
      approvalDate: "",
      publishOptionKey: "",
      reviewersName: "",
      expiryCheck: "",
      expiryDate: null,
      expiryLeadPeriod: "",
      criticalDocument: false,
      templateDocument: false,
      titleReadonly: true,
      saveDisable: false,
      legalEntityOption: [],
      businessUnitOption: [],
      departmentOption: [],
      approvalDateEdit: new Date(),
      categoryOption: [],
      businessUnitCode: "",
      departmentCode: "",
      subCategoryArray: [],
      subCategoryId: null,
      reviewers: [],
      approver: null,
      approverEmail: "",
      approverName: "",
      owner: "",
      ownerEmail: "",
      ownerName: "",
      templateId: "",
      publishOption: "Native",
      incrementSequenceNumber: "",
      documentid: "",
      documentName: "",
      businessUnit: "",
      category: "",
      subCategory: "",
      department: "",
      newDocumentId: "",
      sourceDocumentId: "",
      templateKey: "",
      dcc: null,
      dccEmail: "",
      dccName: "",
      revisionCoding: "",
      revisionLevel: "",
      transmittalCheck: false,
      externalDocument: false,
      revisionCodingId: null,
      revisionLevelId: null,
      revisionLevelArray: [],
      revisionSettingsArray: [],
      categoryCode: "",
      projectName: "",
      projectNumber: "",
      currentRevision: "",
      previousRevisionItemID: null,
      revisionItemID: "",
      newRevision: "",
      hideloader: true,
      legalEntity: "",
      updateDisable: false,
      hideLoading: true,
      workflowStatus: "",
      leadmsg: "none",
      invalidQueryParam: "",
      isdocx: "none",
      nodocx: "",
      insertdocument: "none",
      loaderDisplay: "",
      checkdirect: "none",
      hideDirect: "",
      validApprover: "none",
      createDocumentCheckBoxDiv: "",
      replaceDocument: "none",
      hideSelectTemplate: "none",
      validDocType: "none",
      checkrename: "none",
      subContractorNumber: "",
      customerNumber: "",
      linkToDoc: "",
      hideCreateLoading: "none",
      norefresh: "none",
      upload: false,
      template: false,
      hideupload: "none",
      sourceId: "Quality",
      hidetemplate: "none",
      hidesource: "none",
      showReviewModal: false,
      DueDate: new Date(),
      sendForReview: false,
      dueDateMadatory: "",
      comments: ""
    };
    this._Service = new cdmsEditService(this.props.context, window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._selectedOwner = this._selectedOwner.bind(this);
    this._selectedReviewers = this._selectedReviewers.bind(this);
    this._selectedApprover = this._selectedApprover.bind(this);
    this._onCreateDocChecked = this._onCreateDocChecked.bind(this);
    this._templatechange = this._templatechange.bind(this);
    this._onDirectPublishChecked = this._onDirectPublishChecked.bind(this);
    this._onApprovalDatePickerChange = this._onApprovalDatePickerChange.bind(this);
    this._publishOptionChange = this._publishOptionChange.bind(this);
    this._onExpiryDateChecked = this._onExpiryDateChecked.bind(this);
    this._onExpDatePickerChange = this._onExpDatePickerChange.bind(this);
    this._expLeadPeriodChange = this._expLeadPeriodChange.bind(this);
    this._onCriticalChecked = this._onCriticalChecked.bind(this);
    this._onTemplateChecked = this._onTemplateChecked.bind(this);
    this._sourcechange = this._sourcechange.bind(this);
    this._onUploadCheck = this._onUploadCheck.bind(this);
    this._onTemplateCheck = this._onTemplateCheck.bind(this);
    this._updateDocumentIndex = this._updateDocumentIndex.bind(this);
    this._revisionCoding = this._revisionCoding.bind(this);
    this._updateSourceDocument = this._updateSourceDocument.bind(this);
    this._updateDocument = this._updateDocument.bind(this);
    this._updatePublishDocument = this._updatePublishDocument.bind(this);
    this._onUpdateClick = this._onUpdateClick.bind(this);
    this._updateWithoutDocument = this._updateWithoutDocument.bind(this);
    this._add = this._add.bind(this);
    // this._checkRename = this._checkRename.bind(this);
    this._onReplaceDocumentChecked = this._onReplaceDocumentChecked.bind(this);
    this._onSendForReview = this._onSendForReview.bind(this);
    this.onConfirmReview = this.onConfirmReview.bind(this);
    this._dialogCloseButton = this._dialogCloseButton.bind(this);
    this._closeModal = this._closeModal.bind(this);
  }

  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: { required: "This field is mandatory" }
    });
  }
  //On Load
  public async componentDidMount() {
    //Huburl
    this.siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    this.indexUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList;


    //Get Current User
    const user = await this._Service.getCurrentUser();
    this.currentEmail = user.Email;
    this.currentId = user.Id;
    //Get Today
    this.today = new Date();
    this.setState({ approvalDate: this.today });
    // //for getting  sourcedoument library ID
    // this._Service.libraryByTitle(this.props.sourceDocumentViewLibrary).then(results => {
    //   console.log(results.Id);
    //   this.sourceDocumentLibraryId = results.Id;
    // });
    // console.log(this.sourceDocumentLibraryId);


    this._queryParamGetting();
  }
  //Search Query
  private async _queryParamGetting() {
    this.setState({ accessDeniedMessageBar: "none", createDocumentView: "none", createDocumentProject: "none", qdmsEditDocumentView: "none", projectEditDocumentView: "none", revokeExpiryView: "none", });
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let documentindexid = params.get('did');
    this.documentIndexID = documentindexid;
    if (documentindexid != "" && documentindexid != null) {
      this._Service.itemsFromIndex(this.props.siteUrl, this.props.documentIndexList, Number(documentindexid)).then(DocumentStatus => {
        this.sourceDocumentID = DocumentStatus.SourceDocumentID;
        if ((DocumentStatus.WorkflowStatus != "Under Review" && DocumentStatus.WorkflowStatus != "Under Approval")) {
          if (DocumentStatus.DocumentStatus == "Active") {
            this.setState({ accessDeniedMessageBar: "none", qdmsEditDocumentView: "none", projectEditDocumentView: "none" });
            //Permission handiling 
            this.setState({
              qdmsEditDocumentView: "", projectEditDocumentView: "none", accessDeniedMessageBar: "none", loaderDisplay: "none"
            });
            this._bindDataEditQdms(this.documentIndexID);
            // this._accessGroups('QDMS_EditDocument');
          }
          else {
            this.setState({
              qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "", loaderDisplay: "none",
              statusMessage: { isShowMessage: true, message: "Document is not active right now", messageType: 1 },
            });
            setTimeout(() => {
              this.setState({ accessDeniedMessageBar: 'none', });
              window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
            }, 10000);
          }
        }
        else {
          this.setState({
            qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "",
            statusMessage: { isShowMessage: true, message: "Document is already gone in a workflow", messageType: 1 },
          });
          setTimeout(() => {
            this.setState({ accessDeniedMessageBar: 'none', });
            window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
          }, 10000);
        }
      });
    }
    else {
      this.setState({
        qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "",
        statusMessage: { isShowMessage: true, message: this.state.invalidQueryParam, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMessageBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
      }, 10000);
    }

    this._userMessageSettings();
  }
  // Bind data in qdms
  public async _bindDataEditQdms(documentindexid) {
    // this._checkRename('QDMS_RenameDocument');
    const indexItems = await this._Service.getItemsByID(this.props.siteUrl, this.props.documentIndexList, Number(documentindexid));

    console.log("dataForEdit", indexItems);
    let tempReviewers: any[] = [];
    let temReviewersID: any[] = [];
    if (documentindexid != "" && documentindexid != null) {

      this._Service.itemsFromIndexExpanded(this.props.siteUrl, this.props.documentIndexList, documentindexid).then(async dataForEdit => {
        console.log("dataForEdit", dataForEdit);
        this.setState({
          title: dataForEdit.Title,
          documentid: dataForEdit.DocumentID,
          documentName: dataForEdit.DocumentName,
          businessUnit: dataForEdit.BusinessUnit,
          department: dataForEdit.DepartmentName,
          category: dataForEdit.Category,
          ownerName: dataForEdit.Owner.Title,
          expiryLeadPeriod: dataForEdit.ExpiryLeadPeriod,
          owner: dataForEdit.Owner.ID,
          ownerEmail: dataForEdit.Owner.EMail,
          legalEntity: dataForEdit.LegalEntity,
          subCategory: dataForEdit.SubCategory,
          businessUnitID: dataForEdit.BusinessUnitID,
          departmentId: dataForEdit.DepartmentID
        });
        if (indexItems.ApproverId != null) {
          this.setState({
            approver: dataForEdit.Approver.ID,
            approverName: dataForEdit.Approver.Title
          });
        }
        if (dataForEdit.SourceDocument != null) {
          this.setState({
            linkToDoc: dataForEdit.SourceDocument.Url,
          });
        }
        for (var k in dataForEdit.Reviewers) {
          temReviewersID.push(dataForEdit.Reviewers[k].ID);
          this.setState({
            reviewers: temReviewersID,
          });
          tempReviewers.push(dataForEdit.Reviewers[k].Title);
        }

        if (indexItems.SubCategoryID != null) {
          this.setState({
            subCategoryId: parseInt(dataForEdit.SubCategoryID)
          });
        }
        if (dataForEdit.ExpiryDate != null) {
          let date = new Date(dataForEdit.ExpiryDate);
          this.setState({ expiryDate: date, expiryCheck: true, hideExpiry: "" });
        }
        if (dataForEdit.CriticalDocument == true) {
          this.setState({ criticalDocument: true });
        }
        if (dataForEdit.CreateDocument == true) {
          this.setState({ createDocument: true, hideCreate: "", createDocumentCheckBoxDiv: "none", replaceDocument: "", hidePublish: "none", hideDoc: "none", hideDirect: "none" });
          this.isDocument = "Yes";
        }
        if (dataForEdit.Template == true) {
          this.setState({ templateDocument: true });
        }
        if (dataForEdit.DirectPublish == true) {
          let date = new Date(dataForEdit.ApprovedDate);
          this.setState({ directPublishCheck: true, hidePublish: "none", publishOptionKey: dataForEdit.PublishFormat, approvalDateEdit: date });
        }
        this.setState({
          reviewersName: tempReviewers,
        });
        // if(dataForEdit.WorkflowStatus == "Draft"||dataForEdit.WorkflowStatus == "Published"){
        //   this.setState({
        //     hidebutton: "none",
        //   });
        // }
      });
    }

  }
  // Check permission to rename
  // public async _checkRename(type) {
  //   this.setState({ checkrename: "" });
  //   const laUrl = await this._Service.getQDMSPermissionWebpart(this.props.siteUrl, this.props.requestList);
  //   console.log("Posturl", laUrl[0].PostUrl);
  //   this.permissionpostUrl = laUrl[0].PostUrl;
  //   let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
  //   const postURL = this.permissionpostUrl;

  //   const requestHeaders: Headers = new Headers();
  //   requestHeaders.append("Content-type", "application/json");
  //   const body: string = JSON.stringify({
  //     'PermissionTitle': type,
  //     'SiteUrl': siteUrl,
  //     'CurrentUserEmail': this.currentEmail

  //   });
  //   const postOptions: IHttpClientOptions = {
  //     headers: requestHeaders,
  //     body: body
  //   };

  //   let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  //   let responseJSON = await response.json();
  //   console.log(responseJSON);
  //   if (response.ok) {
  //     console.log(responseJSON['Status']);
  //     if (responseJSON['Status'] == "Valid") {
  //       this.setState({
  //         titleReadonly: false,
  //         checkrename: "none"
  //       });
  //     }
  //     else {
  //       this.setState({
  //         titleReadonly: true,
  //         checkrename: "none"
  //       });
  //     }
  //   }

  //   else { }
  // }
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.getItemsFromUserMsgSettings(this.props.siteUrl, this.props.userMessageSettings);
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "DirectPublishSuccess") {
        var publishmsg = userMessageSettings[i].Message;
        this.directPublish = replaceString(publishmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "EditDocumentSuccess") {
        var editmsg = userMessageSettings[i].Message;
        this.editDocument = replaceString(editmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "InvalidQueryParams") {
        this.setState({
          invalidQueryParam: userMessageSettings[i].Message,
        });
      }
    }
  }
  //Title Change
  public _titleChange = (ev: React.FormEvent<HTMLInputElement>, title?: string) => {
    this.setState({ title: title || '', saveDisable: false });
  }
  //Owner Change
  public _selectedOwner = (items: any[]) => {
    let ownerEmail;
    let ownerName;
    let getSelectedOwner: any[] = [];
    for (let item in items) {
      ownerEmail = items[item].secondaryText,
        ownerName = items[item].text,
        getSelectedOwner.push(items[item].id);
    }
    this.setState({ owner: getSelectedOwner[0], ownerEmail: ownerEmail, ownerName: ownerName, saveDisable: false });
  }
  //Reviewer Change
  public _selectedReviewers = (items: any[]) => {
    this.getSelectedReviewers = [];
    for (let item in items) {
      this.getSelectedReviewers.push(items[item].id);
    }
    this.setState({ reviewers: this.getSelectedReviewers });
  }
  //Approver Change
  public _selectedApprover = async (items: any[]) => {
    let approverEmail;
    let approverName;
    let getSelectedApprover: any[] = [];
    // if (this.props.project) {
    //   for (let item in items) {
    //     approverEmail = items[item].secondaryText,
    //       approverName = items[item].text,
    //       getSelectedApprover.push(items[item].id);
    //   }
    //   this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
    // }
    {
      this.setState({ validApprover: "", approver: null, approverEmail: "", approverName: "", });

      const departments = await this._Service.getItemsFromDepartments(this.props.siteUrl, this.props.department);
      for (let i = 0; i < departments.length; i++) {
        if (departments[i].ID == this.state.departmentId) {
          const deptapprove = await this._Service.getUserIdByEmail(departments[i].Approver.EMail);
          approverEmail = departments[i].Approver.EMail;
          approverName = departments[i].Approver.Title;
          getSelectedApprover.push(deptapprove.Id);
        }

      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
      setTimeout(() => {
        this.setState({ validApprover: "none" });
      }, 5000);
    }
  }
  //Create Document Change
  public _onCreateDocChecked = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    // this.setState({ checkdirect: "" });
    let publishedDocumentArray: any[] = [];
    let sorted_PublishedDocument: any[];
    if (isChecked) {
      let publishedDocument: any[] = await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary);
      for (let i = 0; i < publishedDocument.length; i++) {
        if (publishedDocument[i].Template == true) {
          let publishedDocumentdata = {
            key: publishedDocument[i].ID,
            text: publishedDocument[i].DocumentName,
          };
          publishedDocumentArray.push(publishedDocumentdata);
        }
      }
      sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);

      this.setState({
        hideDoc: "",
        hideSelectTemplate: "",
        createDocument: true,
        templateDocuments: sorted_PublishedDocument
      });

    }

    else if (!isChecked) {
      this.myfile.value = "";
      this.setState({ hideDirect: "", checkdirect: "none", hideDoc: "", createDocument: false, hidePublish: "none", directPublishCheck: false });
    }

  }
  // On replace document checked
  public _onReplaceDocumentChecked = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    let upload;
    this.setState({ validDocType: "none" });
    if (isChecked) {
      this.setState({
        hideSelectTemplate: "none",
        hideDoc: "none",
        hideupload: "",
        replaceDocumentCheckbox: true,
      });
    }
    else {
      this.setState({
        hideSelectTemplate: "none",
        hideDoc: "none",
        hideupload: "none",
        replaceDocumentCheckbox: false,
      });
      // if (this.props.project) {
      //   upload = "#editproject";
      // }
      // else {
      upload = "#editqdms";
      //}
      // @ts-ignore
      (document.querySelector(upload) as HTMLInputElement).value = null;
    }
  }
  private _onUploadCheck = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ upload: true, template: false, hidesource: "none", hidetemplate: "none" });
    }
    else if (!isChecked) {
      // this.myfile.value = "";
      this.setState({ upload: false, hideupload: "none" });
    }
  }
  private _onTemplateCheck = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    let publishedDocumentArray: any[] = [];
    let sorted_PublishedDocument: any[];
    let qdms = window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl;
    // this.QDMSUrl = window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.QDMSUrl;
    if (isChecked) {
      console.log("site :" + this.siteUrl);
      console.log("qdms :" + qdms);
      //if (!this.props.project)
      // {
        if (this.siteUrl == qdms) {
          this.setState({ hidesource: "none" })
        }

        else {
          this.setState({ hidesource: "" })
        }
        this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
        let publishedDocument: any[] = await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary);
        for (let i = 0; i < publishedDocument.length; i++) {
          if (publishedDocument[i].Template === true && publishedDocument[i].Category === this.state.category) {
            let publishedDocumentdata = {
              key: publishedDocument[i].ID,
              text: publishedDocument[i].DocumentName,
            };
            publishedDocumentArray.push(publishedDocumentdata);
          }
        }
        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
        this.setState({ templateDocuments: sorted_PublishedDocument });
      // }
      // else
      // {

      //   this.setState({ template: true, hidesource: "", upload: false, hideupload: "none", hidetemplate: "" });
      // }

    }
    else if (!isChecked) {
      this.setState({ template: false, hidesource: "none", hidetemplate: "none" });
    }
  }
  // On upload document change
  public _add(e) {
    this.setState({ insertdocument: "none", validDocType: "none" });
    this.myfile = e.target.value;
    //let upload;
    let type;
    let doctype;
    this.isDocument = "Yes";
    // if (this.props.project) {
    //   //  upload = "#editproject";
    // }
    // else {
    //   // upload = "#editqdms";
    // }
    let myfile = e.target.files[0];
    console.log(myfile);
    this.isDocument = "Yes";
    var splitted = myfile.name.split(".");
    type = splitted[splitted.length - 1];
    if (this.state.replaceDocumentCheckbox == true) {
      var docsplitted = this.state.documentName.split(".");
      doctype = docsplitted[docsplitted.length - 1];
      if (doctype != type) {
        this.setState({ validDocType: "" });
        // @ts-ignore
        (document.querySelector("#editqdms") as HTMLInputElement).value = null;
      }
    }
    if (type == "docx") {
      this.setState({ isdocx: "", nodocx: "none" });
    }
    else {
      this.setState({ isdocx: "none", nodocx: "" });
    }
  }
  public async _sourcechange(option: { key: any; text: any }) {
    this.setState({ hidetemplate: "", sourceId: option.key });
    let publishedDocumentArray: any[] = [];
    let sorted_PublishedDocument: any[];
    if (option.key == "Quality") {
      let publishedDocument: any[] = await this._Service.getqdmsLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary);
      for (let i = 0; i < publishedDocument.length; i++) {
        if (publishedDocument[i].Template == true) {
          let publishedDocumentdata = {
            key: publishedDocument[i].ID,
            text: publishedDocument[i].DocumentName,
          };
          publishedDocumentArray.push(publishedDocumentdata);
        }
      }
      sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
      this.setState({ templateDocuments: sorted_PublishedDocument, sourceId: option.key });
    }
    else {
      let publishedDocument: any[] = await this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
      for (let i = 0; i < publishedDocument.length; i++) {
        if (publishedDocument[i].Template == true) {
          let publishedDocumentdata = {
            key: publishedDocument[i].ID,
            text: publishedDocument[i].DocumentName,
          };
          publishedDocumentArray.push(publishedDocumentdata);
        }
      }
      sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);

      this.setState({ templateDocuments: sorted_PublishedDocument, sourceId: option.key });
    }
  }
  //Template change
  public async _templatechange(option: { key: any; text: any }) {
    this.setState({ insertdocument: "none" });
    this.setState({ templateId: option.key, templateKey: option.text });
    let type: any;
    let publishName: any;
    this.isDocument = "Yes";
    if (this.state.sourceId == "Quality") {
      await this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
        console.log(publishdoc);
        for (let i = 0; i < publishdoc.length; i++) {
          if (publishdoc[i].Id == this.state.templateId) {

            publishName = publishdoc[i].LinkFilename;
          }
        }
        var split = publishName.split(".", 2);
        type = split[1];
        if (type == "docx") {
          this.setState({ isdocx: "", nodocx: "none" });
        }
        else {
          this.setState({ isdocx: "none", nodocx: "" });
        }
      });
    }
    else {
      await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
        console.log(publishdoc);
        for (let i = 0; i < publishdoc.length; i++) {
          if (publishdoc[i].Id == this.state.templateId) {
            publishName = publishdoc[i].LinkFilename;
          }
        }
        var split = publishName.split(".", 2);
        type = split[1];
        if (type == "docx") {
          this.setState({ isdocx: "", nodocx: "none" });
        }
        else {
          this.setState({ isdocx: "none", nodocx: "" });
        }
      });
    }
  }
  // On direct publih checked
  public async _checkdirectPublish(type) {
    const laUrl = await this._Service.getQDMSPermissionWebpart(this.props.siteUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.permissionpostUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.permissionpostUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': type,
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.currentEmail

    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    console.log(responseJSON);
    if (response.ok) {
      console.log(responseJSON['Status']);
      if (responseJSON['Status'] == "Valid") {
        this.setState({ checkdirect: "none", hideDirect: "", hidePublish: "" });
      }
      else {
        this.setState({ checkdirect: "none", hideDirect: "", hidePublish: "none" });
      }
    }
    else { }
  }
  //Direct Publish change
  private _onDirectPublishChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      // if (this.props.project) {
      //   this.setState({ checkdirect: "", });
      //   this._checkdirectPublish('Project_DirectPublish');
      // }
      // else
      {
        this.setState({ checkdirect: "", });
        this._checkdirectPublish('QDMS_DirectPublish');
      }
      this.setState({ hidePublish: "none", directPublishCheck: true, approvalDate: new Date() });
    }
    else if (!isChecked) { this.setState({ hidePublish: "none", directPublishCheck: false, approvalDate: new Date(), publishOption: "" }); }
  }
  //Approval Date Change
  public _onApprovalDatePickerChange = (date: Date): void => {
    this.setState({
      approvalDate: date,
      approvalDateEdit: date
    });
  }
  //PublishOption Change
  public _publishOptionChange(option: { key: any; text: any }) {
    this.setState({ publishOption: option.key, saveDisable: false });
  }
  //Expiry Change
  public _onExpiryDateChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ hideExpiry: "", expiryCheck: true,dateValid: ""  }); }
    else if (!isChecked) { this.setState({ hideExpiry: "", expiryCheck: false, expiryDate: null, expiryLeadPeriod: "" }); }
  }
  //Expiry Date Change
  public _onExpDatePickerChange = (date?: Date): void => {
    this.setState({ expiryDate: date, expiryCheck: true });
  }
  //Expiry Lead Period Change
  public _expLeadPeriodChange = (ev: React.FormEvent<HTMLInputElement>, expiryLeadPeriod: string) => {
    let LeadPeriodformat = /^[0-9]*$/;
    if (expiryLeadPeriod.match(LeadPeriodformat)) {
      if (Number(expiryLeadPeriod) < 101) {
        this.setState({ expiryLeadPeriod: expiryLeadPeriod || '', leadmsg: "none" });
      }
      else {
        this.setState({ leadmsg: "" });
      }
    }
    else {
      this.setState({ leadmsg: "" });
    }
  }
  //Critical Document Change
  public _onCriticalChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ criticalDocument: true }); }
    else if (!isChecked) { this.setState({ criticalDocument: false }); }
  }
  // Template Change
  public _onTemplateChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ templateDocument: true }); }
    else if (!isChecked) { this.setState({ templateDocument: false }); }
  }
  // Back button of version history & revision history in edit form
  public _back = () => {
    window.location.replace(this.props.siteUrl + "/SitePages/EditDocument.aspx?did=" + this.documentIndexID);
  }
  // Format date
  private _onFormatDate = (date: Date): string => {
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  }
  // On update click
  public async _onUpdateClick() {
    if (this.state.createDocument == true && this.isDocument == "Yes" || this.state.createDocument == false) {
      if (this.state.expiryCheck == true) {

        //Validation without direct publish
        if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')) {
          this.setState({ updateDisable: true, saveDisable: false });
          await this._updateDocument();
          this.validator.hideMessages();
        }
        //Validation with direct publish
        else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')) {
          this.setState({ updateDisable: true, hideloader: false, saveDisable: false });
          await this._updateDocument();
          this.validator.hideMessages();
        }
        else {
          this.validator.showMessages();
          this.forceUpdate();
        }

      }
      else {

        //Validation without direct publish
        if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({ updateDisable: true, saveDisable: false, });
          await this._updateDocument();
          this.validator.hideMessages();
        }
        //Validation with direct publish
        else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({ updateDisable: true, hideloader: false, saveDisable: false });
          await this._updateDocument();
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
  }
  //trigger sendForRevew
  protected async _triggerSendForReview(sourceDocumentID, documentIndexId) {
    const laUrl = await this._Service.DocumentSendForReview(this.props.siteUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'IndexId': documentIndexId,
      'Title': this.state.title,
      'DueDate': this.state.DueDate,
      'Comments': this.state.comments,
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  // On update document
  public async _updateDocument() {
    this.setState({
      hideCreateLoading: " ",
      norefresh: " "
    });
    this._userMessageSettings();
    let documentNameExtension;
    let sourceDocumentId;
    //let upload;
    let documentIdname;
    //  upload = "#editqdms";

    // With document
    if (this.state.createDocument == true) {
      await this._updateDocumentIndex();
      // Get file from form
      // @ts-ignore: Object is possibly 'null'.
      console.log((document.querySelector("#editqdms") as HTMLInputElement).files.length)
      // @ts-ignore: Object is possibly 'null'.
      if ((document.querySelector("#editqdms") as HTMLInputElement).files.length != 0) {
        // @ts-ignore: Object is possibly 'null'.
        let myfile = (document.querySelector("#editqdms") as HTMLInputElement).files[0];
        console.log(myfile);
        var splitted = myfile.name.split(".");
        if (this.state.replaceDocumentCheckbox == true) {
          if (this.state.titleReadonly == true) {
            documentNameExtension = this.state.documentName;
          }
          else {
            documentNameExtension = this.state.documentid + " " + this.state.title + '.' + splitted[splitted.length - 1];
          }
        }
        else {
          if (this.state.titleReadonly == true) {
            documentNameExtension = this.state.documentName + '.' + splitted[splitted.length - 1];
          }
          else {
            documentNameExtension = this.state.documentid + " " + this.state.title + '.' + splitted[splitted.length - 1];
          }
        }
        documentIdname = this.state.documentid + '.' + splitted[splitted.length - 1];
        this.documentNameExtension = documentNameExtension;
        // alert(this.documentNameExtension);
        if (myfile.size) {
          // add file to source library
          const fileUploaded = await this._Service.uploadDocument(documentIdname, myfile, this.props.sourceDocumentLibrary);
          if (fileUploaded) {
            console.log("File Uploaded");
            const item = await fileUploaded.file.getItem();
            console.log(item);
            sourceDocumentId = item["ID"];

            this.sourceDocumentID = sourceDocumentId;
            this.setState({ sourceDocumentId: sourceDocumentId });
            // update metadata
            await this._updateSourceDocument();
            if (item) {
              let revision;

              revision = "0";

              this._updatePublishDocument();
              if (this.state.replaceDocumentCheckbox == true) {
                let itemUpdate = {
                  SourceDocument: {
                    Description: this.documentNameExtension,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  DocumentName: this.documentNameExtension,
                  WorkflowStatus: "Draft",
                }
                this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemUpdate);
                this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                setTimeout(() => {
                  window.location.replace(this.indexUrl);
                }, 5000);
              }
              else {
                let logItems = {
                  Title: this.state.documentid,
                  Status: "Document Created",
                  LogDate: this.today,
                  Revision: revision,
                  DocumentIndexId: this.documentIndexID,
                }
                await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems);

                // update document index
                if (this.state.directPublishCheck == false) {
                  let indexItems = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    }
                  }
                  await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, indexItems);


                }
                else {

                  let indexItems = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    ApprovedDate: this.state.approvalDate,
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    },
                  }
                  await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, indexItems);

                }
                await this._triggerPermission(sourceDocumentId);
                this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                if (this.state.directPublishCheck == true) {

                  this.setState({ hideLoading: false, hideCreateLoading: "none" });
                  await this._publish();

                }
                else {
                  this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.indexUrl);
                  }, 5000);
                }
              }
            }
          }
        }
      }
      else if (this.state.templateId != "") {
        let publishName;
        let extension;
        let newDocumentName;

        // Get template
        if (this.state.sourceId === "Quality") {
          let publishdoc = await this.QDMSUrl.getList(this.props.QDMSUrl + "/" + this.props.publisheddocumentLibrary).items.select("LinkFilename,ID")();
          console.log(publishdoc);
          for (let i = 0; i < publishdoc.length; i++) {
            if (publishdoc[i].Id == this.state.templateId) {
              publishName = publishdoc[i].LinkFilename;
            }
          }
          var split = publishName.split(".", 2);
          extension = split[1];
          if (publishdoc) {
            // Add template document to source document
            newDocumentName = this.state.documentName + "." + extension;
            this.documentNameExtension = newDocumentName;
            documentIdname = this.state.documentid + '.' + extension;
            let siteUrl = this.props.siteUrl + "/" + this.props.publisheddocumentLibrary + "/" + this.state.category + + "/" + publishName;
            this._Service.getBuffer(siteUrl).then(templateData => {
              return this._Service.uploadDocument(documentIdname, templateData, this.props.sourceDocumentLibrary);
            }).then(fileUploaded => {
              console.log("File Uploaded");
              fileUploaded.file.getItem().then(async item => {
                console.log(item);
                sourceDocumentId = item["ID"];

                this.sourceDocumentID = sourceDocumentId;
                this.setState({ sourceDocumentId: sourceDocumentId });
                await this._updateSourceDocument();
              }).then(async updateDocumentIndex => {
                let revision;
                // if (this.props.project) {
                //   revision = "-";
                // }
                // else {
                revision = "0";
                //}
                this._updatePublishDocument();
                let logItems = {
                  Title: this.state.documentid,
                  Status: "Document Created",
                  LogDate: this.today,
                  Revision: revision,
                  DocumentIndexId: this.documentIndexID,
                }
                await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems);
                // update document index
                if (this.state.directPublishCheck == false) {
                  let itemUpdateItems = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    DocumentName: this.documentNameExtension,

                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    },
                  }
                  this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemUpdateItems);
                }
                else {
                  let itemUpdateItems = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    ApprovedDate: this.state.approvalDate,
                    DocumentName: this.documentNameExtension,

                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    }
                  }

                  this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemUpdateItems);

                }
                await this._triggerPermission(sourceDocumentId);
                this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                if (this.state.directPublishCheck == true) {
                  this.setState({ hideLoading: false, hideCreateLoading: "none" });
                  await this._publish();
                }
                else {
                  this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.indexUrl);
                  }, 5000);
                }
              });
            });
          }
        }


      }
      else {
        this._updateWithoutDocument();
      }
    }
    else {
      await this._updateDocumentIndex();
      this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
      setTimeout(() => {
        window.location.replace(this.indexUrl);
      }, 5000);
    }
  }
  // Update Document Index
  public _updateDocumentIndex() {
    // Without Expiry date
    if ((this.state.expiryCheck == false) || (this.state.expiryCheck == "")) {
      let itemUpdate = {
        Title: this.state.title,
        DocumentName: this.state.documentid + " " + this.state.title,
        SubCategoryID: this.state.subCategoryId,
        SubCategory: this.state.subCategory,
        OwnerId: this.state.owner,
        ApproverId: this.state.approver,
        CreateDocument: this.state.createDocument,
        Template: this.state.templateDocument,
        CriticalDocument: this.state.criticalDocument,
        PublishFormat: this.state.publishOption,
        DirectPublish: this.state.directPublishCheck,
        ApprovedDate: this.state.approvalDateEdit,
        ReviewersId: this.state.reviewers,
        WorkflowStatus: this.state.sendForReview === true ? "Under Review" : "Draft",
      }
      this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemUpdate).then(afteradd => {
        this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID + "";
        this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.documentIndexID + "&mode=expiry";
      });
    }
  }
  // Update Source Document
  public _updateSourceDocument() {
    // Without Expiry date
    if (this.state.expiryCheck == false) {
      let updateItems = {
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
        DocumentIndexId: this.documentIndexID,
        PublishFormat: this.state.publishOption,
        Template: this.state.templateDocument,
        OwnerId: this.state.owner,
        DepartmentName: this.state.department,
        CriticalDocument: this.state.criticalDocument,
        RevisionHistory: {
          Description: "Revision History",
          Url: this.revisionHistoryUrl
        }
      }
      this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, updateItems);
    }
    // With Expiry Date
    else {
      let updateItems = {
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
        WorkflowStatus: this.state.sendForReview === true ? "Under Review" : "Draft",
        DocumentStatus: "Active",
        DocumentIndexId: this.documentIndexID,
        PublishFormat: this.state.publishOption,
        Template: this.state.templateDocument,
        OwnerId: this.state.owner,
        DepartmentName: this.state.department,
        CriticalDocument: this.state.criticalDocument,
        RevisionHistory: {
          Description: "Revision History",
          Url: this.revisionHistoryUrl
        }
      }

      this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, updateItems);

    }
  }
  // Update Publish Document
  public async _updatePublishDocument() {
    const publishDocumentID: any[] = await this._Service.itemIDFromPublish(this.props.siteUrl, this.props.publisheddocumentLibrary, this.documentIndexID);
    console.log("publishDocumentID", publishDocumentID);
    if (publishDocumentID.length > 0) {
      // Without Expiry date
      if (this.state.expiryCheck == false) {

        for (var s in publishDocumentID) {
          let libraryUpdate = {
            Title: this.state.title,
            DocumentName: this.documentNameExtension,
            SubCategory: this.state.subCategory,
            OwnerId: this.state.owner,
            ApproverId: this.state.approver,
            Template: this.state.templateDocument,
            PublishFormat: this.state.publishOption,
            ReviewersId: this.state.reviewers
          }
          this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.publisheddocumentLibrary, publishDocumentID[s].ID, libraryUpdate);

        }

      }
      // With Expiry date
      else {

        for (var j in publishDocumentID) {
          let libraryUpdate = {
            Title: this.state.title,
            DocumentName: this.documentNameExtension,
            SubCategory: this.state.subCategory,
            OwnerId: this.state.owner,
            ExpiryLeadPeriod: this.state.expiryLeadPeriod,
            ExpiryDate: this.state.expiryDate,
            ApproverId: this.state.approver,
            Template: this.state.templateDocument,
            PublishFormat: this.state.publishOption,
            ReviewersId: this.state.reviewers,
          }
          this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.publisheddocumentLibrary, publishDocumentID[j].ID, libraryUpdate);
        }

      }
    }
  }
  // Update without document
  public async _updateWithoutDocument() {
    let sourceUrl;
    let extensionSplit;
    const sourceLink: any = await this._Service.getSourceLink(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
    console.log(sourceLink);
    sourceUrl = sourceLink.SourceDocument.Url;
    var split = sourceLink.SourceDocument.Description.split(".", 2);
    extensionSplit = split[1];
    if (sourceLink) {

      if (this.state.directPublishCheck == false) {
        this.setState({
          approvalDate: null,
        });
      }

      let libraryUpdate = {
        Title: this.state.title,
        DocumentName: this.state.documentid + this.state.title + "." + extensionSplit,
        OwnerId: this.state.owner,
        ExpiryLeadPeriod: this.state.expiryLeadPeriod,
        ExpiryDate: this.state.expiryDate,
        ApproverId: this.state.approver,
        Template: this.state.templateDocument,
        PublishFormat: this.state.publishOption,
        ReviewersId: this.state.reviewers
      }
      const results = await this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, libraryUpdate);
      if (results) {
        const publishDocumentID: any[] = await this._Service.itemIDFromPublish(this.props.siteUrl, this.props.publisheddocumentLibrary, this.documentIndexID);
        console.log("publishDocumentID", publishDocumentID);
        for (var k in publishDocumentID) {
          let libraryUpdate = {
            Title: this.state.title,
            DocumentName: this.state.documentid + this.state.title + "." + extensionSplit,
            OwnerId: this.state.owner,
            ExpiryLeadPeriod: this.state.expiryLeadPeriod,
            ExpiryDate: this.state.expiryDate,
            ApproverId: this.state.approver,
            Template: this.state.templateDocument,
            PublishFormat: this.state.publishOption,
            ReviewersId: this.state.reviewers
          }
          this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.publisheddocumentLibrary, publishDocumentID[k].ID, libraryUpdate);

        }
        let indexUpdate = {
          DocumentName: this.state.documentid + " " + this.state.title + "." + extensionSplit,
          SourceDocument: {
            Description: this.state.documentid + " " + this.state.title + "." + extensionSplit,
            Url: sourceUrl
          }
        }
        this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, indexUpdate);

      }
      // }

      this.setState({
        statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 },
        messageBar: "",
        hideCreateLoading: "none", norefresh: "none"
      });
      setTimeout(() => {
        window.location.replace(this.indexUrl);
      }, 5000);


    }
  }
  // Document permission
  protected async _triggerPermission(sourceDocumentID) {
    const laUrl = await this._Service.DocumentPermission(this.props.siteUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  //Document Published
  protected async _publish() {

    await this._revisionCoding();
    const laUrl = await this._Service.DocumentPublish(this.props.siteUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'Status': 'Published',
      'PublishFormat': this.state.publishOption,
      'SourceDocumentID': this.state.sourceDocumentId,
      'SiteURL': siteUrl,
      'PublishedDate': this.today,
      'DocumentName': this.state.documentName,
      'Revision': this.state.newRevision,
      'SourceDocumentLibrary': this.props.sourceDocumentViewLibrary,
      'WorkflowStatus': "Published",
      'RevisionUrl': this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.state.newDocumentId,
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };

    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    console.log(responseJSON);
    if (response.ok) {
      this._publishUpdate(responseJSON.PublishDocID);
    }
    else { }
  }
  // QDMS revision coding
  public _revisionCoding = async () => {
    let revision = parseInt("0");
    let rev = revision + 1;
    this.setState({ newRevision: rev.toString() });

  }
  // Published Document Metadata update
  public async _publishUpdate(publishid) {
    await this._Service.itemFromLibraryByID(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId);
    let itemToUpdate = {
      PublishFormat: this.state.publishOption,
      WorkflowStatus: "Published",
      Revision: this.state.newRevision,
      ApprovedDate: new Date()
    }
    await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemToUpdate);

    if (this.state.owner != this.currentId) {
      this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
    }
    let itemToLog = {
      Title: this.state.documentid,
      Status: "Published",
      LogDate: this.today,
      Revision: this.state.newRevision,
      DocumentIndexId: this.documentIndexID,
    }
    await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, itemToLog);

    this.setState({ hideLoading: true, norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.directPublish, messageType: 4 } });
    setTimeout(() => {
      window.location.replace(this.siteUrl);
    }, 5000);
  }
  //Send Mail
  public _sendMail = async (emailuser, type, name) => {
    let formatday = moment(this.today).format('DD/MMM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    console.log(this.state.criticalDocument);
    const notificationPreference: any[] = await this._Service.itemFromPrefernce(this.props.siteUrl, this.props.notificationPreference, emailuser);
    console.log(notificationPreference[0].Preference);
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference == "Send all emails") {
        mailSend = "Yes";
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.criticalDocument == true) {
        mailSend = "Yes";
      }
      else {
        mailSend = "No";
      }
    }
    else if (this.state.criticalDocument == true) {
      mailSend = "Yes";
    }
    if (mailSend == "Yes") {
      const emailNotification: any[] = await this._Service.getItems(this.props.siteUrl, this.props.emailNotification);
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title == type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }
      }
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      let replaceRequester = replaceString(Body, '[Sir/Madam]', name);
      let replaceDate = replaceString(replaceRequester, '[PublishedDate]', day);
      let replaceApprover = replaceString(replaceDate, '[Approver]', this.state.approverName);
      let replaceBody = replaceString(replaceApprover, '[DocumentName]', this.state.documentName);
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject,
          "body": {
            "contentType": "Text",
            "content": replaceBody
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
        .then((client: MSGraphClientV3): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody);
        });
    }
  }
  //Cancel Document
  private _onCancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });
  }
  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  //Cancel confirm
  private _confirmYesCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
    window.location.replace(this.siteUrl);
  }
  //Not Cancel
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to cancel?',
  };
  private modalProps = {
    isBlocking: true,
  };
  private onUploadOrTemplateRadioBtnChange = async (ev: React.FormEvent<HTMLElement | HTMLInputElement>, option: IChoiceGroupOption) => {

    this.setState({
      uploadOrTemplateRadioBtn: option.key,
      createDocument: true
    });
    if (option.key === "Upload") {
      this.setState({ upload: true, hideupload: "", template: false, hidesource: "none", hidetemplate: "none" });
    }
    if (option.key === "Template") {
      let publishedDocumentArray: any[] = [];
      let sorted_PublishedDocument: any[];
      this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
      let publishedDocument: any[] = await this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
      for (let i = 0; i < publishedDocument.length; i++) {
        if (publishedDocument[i].Template === true) {
          let publishedDocumentdata = {
            key: publishedDocument[i].ID,
            text: publishedDocument[i].DocumentName,
          };
          publishedDocumentArray.push(publishedDocumentdata);
        }
      }
      sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
      this.setState({ templateDocuments: sorted_PublishedDocument });
    }
  }
  public _subCategoryChange(option: { key: any; text: any }) {
    this.setState({ subCategoryId: option.key, subCategory: option.text });
  }


  private _closeModal() {
    this.setState({ showReviewModal: false });
  }
  public _DueDateChange = (date: Date): void => {
    this.setState({ DueDate: date, dueDateMadatory: "" });
  }
  public _commentChange = (ev: React.FormEvent<HTMLInputElement>, comments?: string) => {
    this.setState({ comments: comments || '', });
  }
  public onConfirmReview = async () => {
    if (this.state.DueDate !== null) {
      await this.setState({
        sendForReview: true,
        showReviewModal: false,
        dueDateMadatory: "",
        saveDisable: true, hideCreateLoading: " ",
        norefresh: " "
      }); await this._onUpdateClick();
    }
    else {
      this.setState({ dueDateMadatory: "Yes" });
    }

  }

  private async _onSendForReview() {

    if (this.state.createDocument === true && this.isDocument === "Yes") {
      if (this.state.expiryCheck === true) {
        //Validation without direct publish
        if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')) {
          this.setState({ updateDisable: true });
          if (this.isDocument === "Yes") {
            this.setState({
              showReviewModal: true,
            });
          } else {
            this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
            setTimeout(() => {
              this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
            }, 5000);
          }

          this.validator.hideMessages();
        }
        //Validation with direct publish
        else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')) {
          this.setState({ updateDisable: true, hideloader: false });
          if (this.isDocument === "Yes") {
            this.setState({
              showReviewModal: true,
            });
          } else {
            this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
            setTimeout(() => {
              this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
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
        if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({ updateDisable: true });
          if (this.isDocument === "Yes") {
            this.setState({
              showReviewModal: true,
            });
          } else {
            this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
            setTimeout(() => {
              this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
            }, 5000);
          }
          this.validator.hideMessages();
        }
        //Validation with direct publish
        else if ((this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({ updateDisable: true, hideloader: false });
          if (this.isDocument === "Yes") {
            this.setState({
              showReviewModal: true,
            });
          } else {
            this.setState({ statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
            setTimeout(() => {
              this.setState({ statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
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
      this.setState({ messageBar: "", statusMessage: { isShowMessage: true, message: "Please select documnet", messageType: 1 } });
      setTimeout(() => {
        this.setState({ messageBar: "none", statusMessage: { isShowMessage: false, message: "Please select documnet", messageType: 4 } });
      }, 5000);
    }

  }

  public render(): React.ReactElement<IEditDocumentProps> {

    const publishOptions: IDropdownOption[] = [
      { key: 'PDF', text: 'PDF' },
      { key: 'Native', text: 'Native' },
    ];
    const publishOption: IDropdownOption[] = [
      { key: 'Native', text: 'Native' },
    ];
    const Source: IDropdownOption[] = [
      { key: 'Quality', text: 'Quality' },
      { key: 'Current Site', text: 'Current Site' }
    ];
    const uploadOrTemplateRadioBtnOptions:
      IChoiceGroupOption[] = [
        { key: 'Upload', text: 'Upload existing file' },
        { key: 'Template', text: 'Create document using existing template', styles: { field: { marginLeft: "35px" } } },
      ];
    const calloutProps = { gapSpace: 0 };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const choiceGroupStyles: Partial<IChoiceGroupStyles> = { root: { display: 'flex' }, flexContainer: { display: "flex", justifyContent: 'space-between' } };
    const cancelIcon: IIconProps = { iconName: 'Cancel' };
    const theme = getTheme();
    const contentStyles = mergeStyleSets({
      container: {
        width: "40%",
        marginLeft: "8%",
        borderRadius: "12px"
      }
    });
    const iconButtonStyles = {
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

    return (
      <section className={`${styles.editDocument}`}>
        <div>
          <div style={{ display: this.state.loaderDisplay }}>
            <ProgressIndicator label="Loading......" />
          </div>
          {/* Edit Document QDMS */}
          <div style={{ display: this.state.qdmsEditDocumentView }} >
            <div className={styles.border}>
              <div className={styles.alignCenter}>{this.props.webpartHeader}</div>
              <Pivot aria-label="Links of Tab Style Pivot Example" >
                <PivotItem headerText="Document Info" >
                  <div style={{ display: "flex" }}>
                    <Label>Document ID : </Label> <div className={styles.divLabel}>{this.state.documentid}</div>
                  </div>
                  <div>
                    <TextField required id="t1"
                      label="Title"
                      onChange={this._titleChange}
                      value={this.state.title} readOnly={this.state.titleReadonly} />
                    <div style={{ color: "#dc3545", fontWeight: "bold", display: this.state.checkrename }}>Checking your permission to rename.Please wait...</div>
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Name", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
                  </div>
                  <div className={styles.divrow}>
                    <div className={styles.wdthrgt}>
                      <TextField
                        label="Department"
                        value={this.state.department} readOnly />
                    </div>
                    <div className={styles.wdthlft}>
                      <TextField
                        label="Category"
                        value={this.state.category} readOnly />
                    </div>
                    <div className={styles.wdthlft}>
                      <TextField
                        label="Sub Category"
                        value={this.state.subCategory} readOnly />
                    </div>
                  </div>
                  <div style={{ display: this.state.hideCreate }}>
                    <div className={styles.divrow}>
                      <div style={{ display: this.state.replaceDocument }}>
                        <div style={{ display: "flex" }}><Label >Document :</Label>
                          <div className={styles.divLabel}> <a href={this.state.linkToDoc} target="_blank">
                            {this.state.documentName}</a>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className={styles.documentMainDiv} style={{ display: this.state.createDocumentCheckBoxDiv }}>
                      <div className={styles.radioDiv} style={{ display: this.state.hideDoc }}>
                        <ChoiceGroup selectedKey={this.state.uploadOrTemplateRadioBtn}
                          onChange={this.onUploadOrTemplateRadioBtnChange}
                          options={uploadOrTemplateRadioBtnOptions} styles={choiceGroupStyles}
                        />
                      </div>
                      <div className={styles.uploadDiv} style={{ display: this.state.hideupload }}>
                        <div ><input type="file" name="myFile" id="editqdms" onChange={this._add} /></div>
                        <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select Document </div>
                      </div>
                      <div className={styles.templateDiv} style={{ display: this.state.hidetemplate }}>
                        <div className={styles.divColumn2} style={{ display: "flex" }}>
                          <div className={styles.divColumn2}>
                            <Dropdown id="t7"
                              label="Source"
                              placeholder="Select an option"
                              selectedKey={this.state.sourceId}
                              options={Source}
                              onChanged={this._sourcechange} />
                          </div>
                          <div className={styles.divColumn2} >
                            <Dropdown id="t7"
                              label="Select a Template"
                              placeholder="Select an option"
                              selectedKey={this.state.templateId}
                              options={this.state.templateDocuments}
                              onChanged={this._templatechange} style={{ width: "150%", maxWidth: "150%" }} />
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className={styles.divrow}>
                      <div className={styles.wdthfrst}>
                        {/* <div style={{ display: this.state.createDocumentCheckBoxDiv }}>
                          <TooltipHost
                            content="Check if the template or attachment is added"
                            //id={tooltipId}
                            calloutProps={calloutProps}
                            styles={hostStyles}>
                            <Checkbox label="Create Document ? " boxSide="start"
                              onChange={this._onCreateDocChecked}
                              checked={this.state.createDocument} />
                          </TooltipHost>
                        </div> */}
                        <div style={{ display: this.state.replaceDocument }}>
                          <TooltipHost
                            content="Check if the template or attachment is added"
                            //id={tooltipId}
                            calloutProps={calloutProps}
                            styles={hostStyles}>
                            <Checkbox label="Replace Document ? " boxSide="start"
                              onChange={this._onReplaceDocumentChecked}
                              checked={this.state.replaceDocumentCheckbox} />
                          </TooltipHost>
                        </div>
                      </div>
                      {/* <div className={styles.wdthmid} style={{ display: this.state.hideDoc }}>
                        <Checkbox label="Upload existing file " boxSide="start"
                          onChange={this._onUploadCheck}
                          checked={this.state.upload} />
                      </div>
                      <div className={styles.wdthlst} style={{ display: this.state.hideDoc }}>
                        <Checkbox label="Create document using existing template" boxSide="start"
                          onChange={this._onTemplateCheck}
                          checked={this.state.template} />
                      </div> */}
                    </div>
                    {this.state.replaceDocumentCheckbox === true &&
                      <div className={styles.divrow} style={{ marginTop: "10px" }}>
                        <div className={styles.wdthfrst}><Label>Upload Document:</Label></div>
                        <div className={styles.wdthmid}>
                          <input type="file" name="myFile" id="editqdms" onChange={this._add} />
                          <div style={{ display: this.state.validDocType, color: "#dc3545" }}>Please select valid Document </div>
                          <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select valid Document or Please uncheck Create Document</div>
                        </div>


                      </div>
                    }
                    <div className={styles.divrow}>
                      <div className={styles.wdthfrst} style={{ display: this.state.hideDirect }}>
                        <TooltipHost
                          content="The document to published library without sending it for review/approval"
                          //id={tooltipId}
                          calloutProps={calloutProps}
                          styles={hostStyles}>
                          <Checkbox label="Direct Publish?" boxSide="start" onChange={this._onDirectPublishChecked} checked={this.state.directPublishCheck} />
                        </TooltipHost></div>
                      <div className={styles.wdthmid} style={{ display: this.state.checkdirect }}><Spinner label={'Please Wait...'} /></div>
                      <div className={styles.wdthmid} style={{ display: this.state.hidePublish }}>
                        <DatePicker label="Published Date"
                          style={{ width: '200px' }}
                          value={this.state.approvalDateEdit}
                          onSelectDate={this._onApprovalDatePickerChange}
                          placeholder="Select a date..."
                          ariaLabel="Select a date" minDate={new Date()} maxDate={new Date()}
                          formatDate={this._onFormatDate} /></div>
                      <div className={styles.wdthlst} style={{ display: this.state.hidePublish }}>
                        <div style={{ display: this.state.isdocx }}>
                          <Dropdown id="t2" required={true}
                            label="Publish Option"
                            selectedKey={this.state.publishOption}
                            placeholder="Select an option"
                            options={publishOptions}
                            onChanged={this._publishOptionChange} /></div>
                        <div style={{ display: this.state.nodocx }}>
                          <Dropdown id="t2" required={true}
                            label="Publish Option"
                            selectedKey={this.state.publishOption}
                            placeholder="Select an option"
                            options={publishOption}
                            onChanged={this._publishOptionChange} /></div>
                        <div style={{ color: "#dc3545" }}>
                          {this.validator.message("publish", this.state.publishOption, "required")}{""}</div>
                      </div>
                    </div>
                  </div>
                  <div className={styles.divrow}>
            <div style={{ width: "77%" }}>
              <PeoplePicker
                context={this.props.context as any}
                titleText="Owner"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users    
                showtooltip={true}
                required={true}
                disabled={false}
                ensureUser={true}
                onChange={this._selectedOwner}
                defaultSelectedUsers={[this.props.context.pageContext.user.email]}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />

            </div>
            <div style={{ width: "75%", marginLeft: "10px" }}>
              <PeoplePicker
                context={this.props.context as any}
                titleText="Reviewer(s)"
                personSelectionLimit={10}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                required={false}
                disabled={false}
                ensureUser={true}
                showHiddenInUI={false}
                onChange={(items) => this._selectedReviewers(items)}
                defaultSelectedUsers={this.state.reviewersName}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                peoplePickerCntrlclassName={"testClass"}
              />

            </div>
            <div className={styles.divApprover}>
              <PeoplePicker
                context={this.props.context as any}
                titleText="Approver"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users    
                showtooltip={true}
                required={true}
                disabled={false}
                ensureUser={true}
                onChange={this._selectedApprover}
                showHiddenInUI={false}
                defaultSelectedUsers={[this.state.approverName]}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />

            </div>
          </div>
                  <div className={styles.divrow}>
            <div className={styles.divDate} style={{ display: this.state.hideExpiry }}>
              <DatePicker label="Expiry Date"
                value={this.state.expiryDate}
                onSelectDate={this._onExpDatePickerChange}
                placeholder="Select a date..."
                ariaLabel="Select a date"
                minDate={new Date()}
                formatDate={this._onFormatDate} />
              <div style={{ display: this.state.dateValid }}>
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("expiryDate", this.state.expiryDate, "required")}{""}
                </div></div>
            </div>
            <div className={styles.wdthmid} style={{ display: this.state.hideExpiry, width: "14.5%", }}>
              <TextField id="Expiry Reminder" name="Expiry Reminder (Days)"
                label="Expiry Reminder(Days)" onChange={this._expLeadPeriodChange}
                value={this.state.expiryLeadPeriod}>
              </TextField>
              <div style={{ display: this.state.dateValid }}>
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("ExpiryLeadPeriod", this.state.expiryLeadPeriod, "required")}{""}
                </div></div>
              <div style={{ color: "#dc3545", display: this.state.leadmsg }}>
                Enter only numbers less than 100
              </div>
            </div>
            <div style={{ marginTop: "35px", marginLeft: "11px" }}> <TooltipHost
              content="Do you want to make this as a template?"
              //id={tooltipId}
              calloutProps={calloutProps}
              styles={hostStyles}>
              <Checkbox label="Save as template " boxSide="start" onChange={this._onTemplateChecked} checked={this.state.templateDocument} />
            </TooltipHost></div>

            <div style={{ display: this.state.hideDirect, marginTop: "36px", marginLeft: "15px" }}>
              <TooltipHost
                content="Without review or approval, the document will be published."
                //id={tooltipId}
                calloutProps={calloutProps}
                styles={hostStyles}>
                <Checkbox label="Direct Publish?" boxSide="start" onChange={this._onDirectPublishChecked} checked={this.state.directPublishCheck} />
              </TooltipHost></div>
            <div style={{ marginLeft: "31px", display: this.state.hidePublish }}>
              <Dropdown id="t2" required={true}
                label="Publish Option"
                selectedKey={this.state.publishOption}
                placeholder="Select an option"
                options={this.state.isdocx === "" ? publishOptions : publishOption}
                onChanged={this._publishOptionChange} />
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("publish", this.state.publishOption, "required")}{""}</div>
            </div>

          </div>

                  <div style={{ display: this.state.messageBar }}>
                    {/* Show Message bar for Notification*/}
                    {this.state.statusMessage.isShowMessage ?
                      <MessageBar
                        messageBarType={this.state.statusMessage.messageType}
                        isMultiline={false}
                        dismissButtonAriaLabel="Close"
                      >{this.state.statusMessage.message}</MessageBar>
                      : ''}
                  </div>
                  <div className={styles.mt}>
                    <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
                  </div>
                  <div className={styles.mt}>
                    <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Updating...'} /></div>
                  </div>
                  <div className={styles.mt}>
                    <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                      <Label>***PLEASE DON'T REFRESH***</Label>
                    </div>
                  </div>
                  <div className={styles.divrow}>
            <div style={{ fontStyle: "italic", fontSize: "12px", position: "absolute" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
            <div className={styles.rgtalign} >
              <PrimaryButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._onSendForReview}>Update & Send for review and submit</PrimaryButton >
              <PrimaryButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._onUpdateClick}>Update</PrimaryButton >
              <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
            </div>
          </div>
                  

                  {/* {/ {/ Cancel Dialog Box /} /} */}

                  <div>
                    <Dialog
                      hidden={this.state.confirmDialog}
                      dialogContentProps={this.dialogContentProps}
                      onDismiss={this._dialogCloseButton}
                      styles={this.dialogStyles}
                      modalProps={this.modalProps}>
                      <DialogFooter>
                        <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                        <DefaultButton onClick={this._confirmNoCancel} text="No" />
                      </DialogFooter>
                    </Dialog>
                  </div>
                  <div style={{ padding: "18px" }} >
                    <Modal
                      isOpen={this.state.showReviewModal}
                      isModeless={true}
                      containerClassName={contentStyles.container}>
                      <div style={{ padding: "18px" }}>
                        <div className={styles.modalHeading} style={{ display: "flex" }}>
                          <span style={{ textAlign: "center", display: "flex", justifyContent: "center", flexGrow: "1" }}><b>Send For Review</b></span>
                          <IconButton
                            iconProps={cancelIcon}
                            ariaLabel="Close popup modal"
                            onClick={this._closeModal}
                            styles={iconButtonStyles}
                          />
                        </div>
                        <DatePicker label="Due Date *"
                          value={this.state.DueDate}
                          onSelectDate={this._DueDateChange}
                          placeholder="Select a date..."
                          ariaLabel="Select a date"
                          minDate={new Date()}
                          formatDate={this._onFormatDate} />
                        {this.state.dueDateMadatory === "Yes" &&
                          <label style={{ color: 'Red' }}>This field is mandatory</label>}
                        <TextField id="comments" autoComplete='true' label="Comments" onChange={this._commentChange} value={this.state.comments} multiline />
                        <PrimaryButton style={{ float: "right", marginTop: "7px", marginBottom: "9px" }} className={styles.modalButton} id="b2" onClick={this.onConfirmReview} >Confirm</PrimaryButton >
                      </div>

                    </Modal>
                  </div>
                  <br />

                  {/* editDocument div close */}

                </PivotItem>

                <PivotItem headerText="Version History">
                  <div>
                    <IconButton iconProps={back} title="Back" onClick={this._back} />
                    {/* <Iframe
                      id="iframeModal"
                      url={this.props.siteUrl + "/_layouts/15/Versions.aspx?list={" + this.sourceDocumentLibraryId + "}&ID=" + this.sourceDocumentID + "&IsDlg=0"}
                      width={"100%"}
                      frameBorder={0}
                      height={"500rem"} /> */}
                  </div>
                </PivotItem>
                <PivotItem headerText="Revision History">
                  <div>
                    <IconButton iconProps={back} title="Back" onClick={this._back} />
                    {/* <Iframe
                      id="iframeModal"
                      url={this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID}
                      width={"100%"}
                      frameBorder={0}
                      height={"500rem"} /> */}
                  </div>
                </PivotItem>
              </Pivot>
            </div>
          </div>

          <div style={{ display: this.state.accessDeniedMessageBar }}>
            {/* Show Message bar for Notification*/}
            {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''}
          </div>
        </div>
      </section>
    );
  }
}
