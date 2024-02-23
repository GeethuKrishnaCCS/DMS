import * as React from 'react';
import styles from './CreateDocument.module.scss';
import { ICreateDocumentProps, ICreateDocumentState } from '../interfaces';
import * as moment from 'moment';
import { MSGraphClientV3, HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import SimpleReactValidator from 'simple-react-validator';
import * as _ from 'lodash';
import replaceString from 'replace-string';
import { DMSService } from '../services';
import { Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, getTheme, IChoiceGroupOption, IChoiceGroupStyles, IconButton, IDropdownOption, IIconProps, ITooltipHostStyles, Label, mergeStyleSets, MessageBar, Modal, PrimaryButton, Spinner, TextField, TooltipHost } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export default class CreateDocument extends React.Component<ICreateDocumentProps, ICreateDocumentState> {

  private _Service: DMSService;
  private validator: SimpleReactValidator;
  private siteUrl;
  private currentEmail;
  private currentId;
  private currentUser;
  private today;
  private createDocument;
  private directPublish;
  private getSelectedReviewers: any[] = [];
  private myfile;
  private isDocument;
  private permissionpostUrl;
  private documentNameExtension;
  private revokeUrl;
  private Timeout = 5000;
  private documentIndexID;
  private revisionHistoryUrl;
  private postUrl;
  public constructor(props: ICreateDocumentProps) {
    super(props);
    this.state = {
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
    }
    this._Service = new DMSService(this.props.context, window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl);
    this._bindData = this._bindData.bind(this);
    this._departmentChange = this._departmentChange.bind(this);
    this._categoryChange = this._categoryChange.bind(this);
    this._subCategoryChange = this._subCategoryChange.bind(this);
    this._selectedOwner = this._selectedOwner.bind(this);
    this._selectedReviewers = this._selectedReviewers.bind(this);
    this._selectedApprover = this._selectedApprover.bind(this);
    this._onCreateDocChecked = this._onCreateDocChecked.bind(this);
    this._sourcechange = this._sourcechange.bind(this);
    this._templatechange = this._templatechange.bind(this);
    this._onDirectPublishChecked = this._onDirectPublishChecked.bind(this);
    this._onApprovalDatePickerChange = this._onApprovalDatePickerChange.bind(this);
    this._publishOptionChange = this._publishOptionChange.bind(this);
    this._onExpiryDateChecked = this._onExpiryDateChecked.bind(this);
    this._onExpDatePickerChange = this._onExpDatePickerChange.bind(this);
    this._expLeadPeriodChange = this._expLeadPeriodChange.bind(this);
    this._onCriticalChecked = this._onCriticalChecked.bind(this);
    this._onTemplateChecked = this._onTemplateChecked.bind(this);
    this._onCreateDocument = this._onCreateDocument.bind(this);
    this._documentidgeneration = this._documentidgeneration.bind(this);
    this._incrementSequenceNumber = this._incrementSequenceNumber.bind(this);
    this._documentCreation = this._documentCreation.bind(this);
    this._addSourceDocument = this._addSourceDocument.bind(this);
    this._createDocumentIndex = this._createDocumentIndex.bind(this);
    this._revisionCoding = this._revisionCoding.bind(this);
    this._legalEntityChange = this._legalEntityChange.bind(this);
    this._add = this._add.bind(this);
    this._checkdirectPublish = this._checkdirectPublish.bind(this);
    this._onUploadCheck = this._onUploadCheck.bind(this);
    this._onTemplateCheck = this._onTemplateCheck.bind(this);
    this._onSendForReview = this._onSendForReview.bind(this);
    this.onConfirmReview = this.onConfirmReview.bind(this);
    this._dialogCloseButton = this._dialogCloseButton.bind(this);
    this._closeModal = this._closeModal.bind(this);
  }


  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: { required: "This field is mandatory" }
    });
  }
  // On load
  public async componentDidMount() {
    //Huburl
    this.siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // this.hubSite = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList;
    //Get Current User
    const user = await this._Service.getCurrentUser();
    this.currentEmail = user.Email;
    this.currentId = user.Id;
    this.currentUser = user.Title;
    //Get Today
    this.today = new Date();
    this.setState({ approvalDate: this.today });
    this.setState({ loaderDisplay: "none" });
    this._bindData();
    this._checkdirectPublish('QDMS_DirectPublish');

  }
  //Bind dropdown in create
  public async _bindData() {
    let businessUnitArray: any[] = [];
    let sorted_BusinessUnit: any[];
    let departmentArray: any[] = [];
    let sorted_Department: any[];
    let categoryArray: any[] = [];
    let sorted_Category: any[];
    let legalEntityArray: any[] = [];
    let sorted_LegalEntity: any[];
    //Get Business Unit
    const businessUnit: any[] = await this._Service.getItems(this.props.siteUrl, this.props.businessUnit);
    for (let i = 0; i < businessUnit.length; i++) {
      let businessUnitdata = {
        key: businessUnit[i].ID,
        text: businessUnit[i].BusinessUnitName,
      };
      businessUnitArray.push(businessUnitdata);
    }
    sorted_BusinessUnit = _.orderBy(businessUnitArray, 'text', ['asc']);
    //Get Department
    const department: any[] = await this._Service.getItems(this.props.siteUrl, this.props.department);
    if (this.props.siteUrl === "/sites/Quality" || "/sites/PropertyManagement") {
      for (let i = 0; i < department.length; i++) {
        let departmentdata: any = {
          key: department[i].ID,
          text: department[i].Department,
        };
        departmentArray.push(departmentdata);
      }
    }
    else {
      for (let i = 0; i < department.length; i++) {
        if (this.props.siteUrl === "/sites/" + department[i].Title) {
          this.setState({
            departmentId: department[i].ID,
          });
          let departmentdata = {
            key: department[i].ID,
            text: department[i].Department,
          };
          departmentArray.push(departmentdata);
          this._departmentChange(departmentdata);
        }
      }

    }

    sorted_Department = _.orderBy(departmentArray, 'text', ['asc']);
    //Get Category
    const category: any[] = await this._Service.getItems(this.props.siteUrl, this.props.category);
    let categorydata;
    for (let i = 0; i < category.length; i++) {
      if (category[i].QDMS == true) {
        categorydata = {
          key: category[i].ID,
          text: category[i].Category,
        };
        categoryArray.push(categorydata);
      }
    }
    sorted_Category = _.orderBy(categoryArray, 'text', ['asc']);
    //Get Legal Entity
    const legalEntity: any = await this._Service.getItems(this.props.siteUrl, this.props.legalEntity);
    for (let i = 0; i < legalEntity.length; i++) {
      let legalEntityItemdata = {
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
  }
  //Messages
  private async _userMessageSettings() {
    // const userMessageSettings: any[] = await this._Service.getItemsFromUserMsgSettings(this.props.siteUrl, this.props.userMessageSettings);
    const userMessageSettings: any[] = await this._Service.getSelectFilter(this.props.siteUrl, this.props.userMessageSettings, "Title,Message", "PageName eq 'DocumentIndex'");
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "CreateDocumentSuccess") {
        var successmsg = userMessageSettings[i].Message;
        this.createDocument = replaceString(successmsg, '[DocumentName]', this.state.documentName);
      }
      if (userMessageSettings[i].Title == "DirectPublishSuccess") {
        var publishmsg = userMessageSettings[i].Message;
        this.directPublish = replaceString(publishmsg, '[DocumentName]', this.state.documentName);
      }

    }
  }
  //Title Change
  public _titleChange = (ev: React.FormEvent<HTMLInputElement>, title?: string) => {
    this.setState({ title: title || '', saveDisable: false });
  }
  //Department Change
  public async _departmentChange(option: { key: any; text: any }) {
    let getApprover: any[] = [];
    let approverEmail;
    let approverName;
    const department = await this._Service.getItemsByID(this.props.siteUrl, this.props.department, option.key);
    let departmentCode = department.Title;
    this.setState({ departmentId: option.key, departmentCode: departmentCode, department: option.text, saveDisable: false });
    if (this.state.businessUnitCode == "") {
      // const departments = await this._Service.getItemsFromDepartments(this.props.siteUrl, this.props.department);
      const departments = await this._Service.getSelectExpand(this.props.siteUrl, this.props.department, "ID,Title,Approver/Title,Approver/ID,Approver/EMail", "Approver");
      for (let i = 0; i < departments.length; i++) {
        if (departments[i].ID == option.key) {
          const deptapprove = await this._Service.getUserIdByEmail(departments[i].Approver.EMail);
          approverEmail = departments[i].Approver.EMail;
          approverName = departments[i].Approver.Title;
          getApprover.push(deptapprove.Id);
        }
      }
      this.setState({ approver: getApprover[0], approverEmail: approverEmail, approverName: approverName });
    }
  }

  //Category Change
  public async _categoryChange(option: { key: any; text: any }) {
    let subcategoryArray: any[] = [];
    let sorted_subcategory: any[];
    let category = await this._Service.getItemsByID(this.props.siteUrl, this.props.category, option.key);
    let categoryCode = category.Title;
    await this._Service.getItems(this.props.siteUrl, this.props.subCategory).then(subcategory => {
      for (let i = 0; i < subcategory.length; i++) {
        if (subcategory[i].CategoryId == option.key) {
          let subcategorydata = {
            key: subcategory[i].ID,
            text: subcategory[i].SubCategory,
          };
          subcategoryArray.push(subcategorydata);
        }
      }
      sorted_subcategory = _.orderBy(subcategoryArray, 'text', ['asc']);
      this.setState({
        categoryId: option.key,
        subCategoryArray: sorted_subcategory,
        category: option.text,
        categoryCode: categoryCode,
        saveDisable: false
      });
    });
    let publishedDocumentArray: any[] = [];
    let sorted_PublishedDocument: any[];
    // let publishedDocument: any[] = await this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
    let publishedDocument: any[] = await this._Service.getItemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
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
  }
  //SubCategory Change
  public _subCategoryChange(option: { key: any; text: any }) {
    this.setState({ subCategoryId: option.key, subCategory: option.text });
  }
  // Legal Entity Change
  public _legalEntityChange(option: { key: any; text: any }) {
    this.setState({ legalEntityId: option.key, legalEntity: option.text });
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

    this.setState({ validApprover: "", approver: null, approverEmail: "", approverName: "", });
    if (this.state.businessUnitCode != "") {
    }
    else {
      // const departments = await this._Service.getItemsFromDepartments(this.props.siteUrl, this.props.department);
      const departments = await this._Service.getSelectExpand(this.props.siteUrl, this.props.department, "ID,Title,Approver/Title,Approver/ID,Approver/EMail", "Approver");
      for (let i = 0; i < departments.length; i++) {
        if (departments[i].ID == this.state.departmentId) {
          const deptapprove = await this._Service.getUserIdByEmail(departments[i].Approver.EMail);
          approverEmail = departments[i].Approver.EMail;
          approverName = departments[i].Approver.Title;
          getSelectedApprover.push(deptapprove.Id);
        }
      }
    }
    this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
    setTimeout(() => {
      this.setState({ validApprover: "none" });
    }, 5000);



  }
  //Create Document Change
  public _onCreateDocChecked = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
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

  }
  private onUploadOrTemplateRadioBtnChange = async (ev: React.FormEvent<HTMLElement | HTMLInputElement>, option: IChoiceGroupOption) => {

    this.setState({
      uploadOrTemplateRadioBtn: option.key,
      createDocument: true
    });
    if (option.key == "Upload") {
      this.setState({ upload: true, hideupload: "", template: false, hidesource: "none", hidetemplate: "none" });
    }
    if (option.key == "Template") {
      let publishedDocumentArray: any[] = [];
      let sorted_PublishedDocument: any[];
      this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
      // let publishedDocument: any[] = await this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
      let publishedDocument: any[] = await this._Service.getItemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
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
    }
  }
  private _onUploadCheck = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ upload: true, hideupload: "", template: false, hidesource: "none", hidetemplate: "none" });
    }
    else if (!isChecked) {
      this.setState({ upload: false, hideupload: "none" });
    }
  }
  private _onTemplateCheck = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    let publishedDocumentArray: any[] = [];
    let sorted_PublishedDocument: any[];
    if (isChecked) {
      this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
      // let publishedDocument: any[] = await this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
      let publishedDocument: any[] = await this._Service.getItemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
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
    else if (!isChecked) {
      this.setState({ template: false, hidesource: "none", hidetemplate: "none" });
    }
  }
  // On upload
  public _add(e) {
    this.setState({ insertdocument: "none" });
    this.myfile = e.target.value;
    let type;
    let myfile;
    this.isDocument = "Yes";
    // @ts-ignore: Object is possibly 'null'.
    myfile = (document.querySelector("#addqdms") as HTMLInputElement).files[0];
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
  }
  public async _sourcechange(option: { key: any; text: any }) {
    this.setState({ hidetemplate: "", sourceId: option.key });
    let publishedDocumentArray: any[] = [];
    let sorted_PublishedDocument: any[];
    if (option.key === "Quality") {
      // let publishedDocument: any[] = await this._Service.getqdmsLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary);
      let publishedDocument: any[] = await this._Service.getItemFromLibrary(this.props.QDMSUrl, this.props.publisheddocumentLibrary);
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
      this.setState({ templateDocuments: sorted_PublishedDocument, sourceId: option.key });
    }
    else {
      // let publishedDocument: any[] = await this._Service.itemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
      let publishedDocument: any[] = await this._Service.getItemFromLibrary(this.props.siteUrl, this.props.publisheddocumentLibrary);
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
    if (this.state.sourceId === "Quality") {
      // await this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
      await this._Service.getSelectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary, "LinkFilename,ID,FileLeafRef,DocumentName").then((publishdoc: any) => {
        console.log(publishdoc);
        for (let i = 0; i < publishdoc.length; i++) {
          if (publishdoc[i].Id === this.state.templateId) {

            publishName = publishdoc[i].LinkFilename;
          }
        }
        var split = publishName.split(".", 2);
        type = split[1];
        if (type === "docx") {
          this.setState({ isdocx: "", nodocx: "none" });
        }
        else {
          this.setState({ isdocx: "none", nodocx: "" });
        }
      });
    }
    else {
      // await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
      await this._Service.getSelectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary, "LinkFilename,ID,Template,DocumentName").then((publishdoc: any) => {
        console.log(publishdoc);
        for (let i = 0; i < publishdoc.length; i++) {
          if (publishdoc[i].Id === this.state.templateId) {
            publishName = publishdoc[i].LinkFilename;
          }
        }
        var split = publishName.split(".", 2);
        type = split[1];
        if (type === "docx") {
          this.setState({ isdocx: "", nodocx: "none" });
        }
        else {
          this.setState({ isdocx: "none", nodocx: "" });
        }
      });
    }
  }
  //Direct Publish change
  private _onDirectPublishChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      // this.setState({ checkdirect: "", });
      // this._checkdirectPublish('QDMS_DirectPublish');
      this.setState({ hidePublish: "", directPublishCheck: true, approvalDate: new Date() });
    }
    else if (!isChecked) {
      this.setState({ hidePublish: "none", directPublishCheck: false, approvalDate: new Date(), publishOption: "" });
    }
  }
  // Direct publish change
  public async _checkdirectPublish(type) {
    // const laUrl = await this._Service.getQDMSPermissionWebpart(this.props.siteUrl, this.props.requestList);
    const laUrl = await this._Service.getItemsFilter(this.props.siteUrl, this.props.requestList, "Title eq 'QDMS_PermissionWebpart'");
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
      if (responseJSON['Status'] === "Valid") {
        if (this.props.directPublish === true) {
          this.setState({ checkdirect: "none", hideDirect: "", hidePublish: "none" });
        }
      }
      else {
        this.setState({ checkdirect: "none", hideDirect: "none", hidePublish: "none" });
      }
    }
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
    if (isChecked) { this.setState({ hideExpiry: "", expiryCheck: true, dateValid: "" }); }
    else if (!isChecked) { this.setState({ hideExpiry: "", expiryCheck: false, expiryDate: null, expiryLeadPeriod: "" }); }
  }
  //Expiry Date Change
  public _onExpDatePickerChange = (date?: Date): void => {
    this.setState({ expiryDate: date });
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
  //On create button click
  public async _onCreateDocument() {
    if (this.state.createDocument === true && this.isDocument === "Yes" || this.state.createDocument === false) {
      if (this.state.expiryCheck === true) {
        //Validation without direct publish
        if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({
            saveDisable: true, hideCreateLoading: " ",
            norefresh: " "
          });
          await this._documentidgeneration();
          this.validator.hideMessages();
        }
        //Validation with direct publish
        else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({
            saveDisable: true, hideloader: false, hideCreateLoading: " ",
            norefresh: " "
          });
          await this._documentidgeneration();
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
          this.setState({
            saveDisable: true, hideCreateLoading: " ",
            norefresh: " "
          });
          await this._documentidgeneration();
          this.validator.hideMessages();
        }
        //Validation with direct publish
        else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
          this.setState({
            saveDisable: true, hideloader: false, hideCreateLoading: " ",
            norefresh: " "
          });
          await this._documentidgeneration();
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


  //Documentid generation
  public async _documentidgeneration() {
    let separator;
    let sequenceNumber;
    let idcode;
    let counter;
    var incrementstring;
    let increment;
    let documentid;
    let isValue = "false";
    let settingsid;
    let documentname;
    // Get document id settings
    const documentIdSettings = await this._Service.getItems(this.props.siteUrl, this.props.documentIdSettings);
    console.log("documentIdSettings", documentIdSettings);
    separator = documentIdSettings[0].Separator;
    sequenceNumber = documentIdSettings[0].SequenceDigit;
    idcode = this.state.departmentCode + separator + this.state.categoryCode;
    if (documentIdSettings) {
      // Get sequence of id
      const documentIdSequenceSettings = await this._Service.getItems(this.props.siteUrl, this.props.documentIdSequenceSettings);
      console.log("documentIdSequenceSettings", documentIdSequenceSettings);
      for (var k in documentIdSequenceSettings) {
        if (documentIdSequenceSettings[k].Title === idcode) {
          counter = documentIdSequenceSettings[k].Sequence;
          settingsid = documentIdSequenceSettings[k].ID;
          isValue = "true";
        }
      }
      if (documentIdSequenceSettings) {
        // No sequence
        if (isValue === "false") {
          increment = 1;
          incrementstring = increment.toString();
          let idsettings = {
            Title: idcode,
            Sequence: incrementstring
          }
          const addidseq = await this._Service.createNewItem(this.props.siteUrl, this.props.documentIdSequenceSettings, idsettings);
          if (addidseq) {
            await this._incrementSequenceNumber(incrementstring, sequenceNumber);

            if (this.state.departmentCode != "") {
              documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
            }
            else {
              documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
            }
            documentname = documentid + " " + this.state.title;

            this.setState({ documentid: documentid, documentName: documentname });
            await this._documentCreation();
          }
        }
        // Has sequence
        else {
          increment = parseInt(counter) + 1;
          incrementstring = increment.toString();
          let idItems = {
            Title: idcode,
            Sequence: incrementstring
          }
          // const afterCounter = await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIdSequenceSettings, settingsid, idItems);
          const afterCounter = await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIdSequenceSettings, settingsid, idItems);
          if (afterCounter) {
            await this._incrementSequenceNumber(incrementstring, sequenceNumber);
            if (this.state.departmentCode != "") {
              documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
            }
            else {
              documentid = this.state.departmentCode + separator + this.state.categoryCode + separator + this.state.incrementSequenceNumber;
            }
            documentname = documentid + " " + this.state.title;
            this.setState({ documentid: documentid, documentName: documentname });
            await this._documentCreation();
          }
        }
      }
    }
  }
  // Append sequence to the count
  public _incrementSequenceNumber(incrementvalue, sequenceNumber) {
    var incrementSequenceNumber = incrementvalue;
    while (incrementSequenceNumber.length < sequenceNumber)
      incrementSequenceNumber = "0" + incrementSequenceNumber;
    console.log(incrementSequenceNumber);
    this.setState({
      incrementSequenceNumber: incrementSequenceNumber,
    });
  }// Create item with id
  public async _documentCreation() {
    await this._userMessageSettings();
    let documentNameExtension;
    let sourceDocumentId;
    let upload;
    let docinsertname;
    upload = "#addqdms";
    // With document
    if (this.state.createDocument === true) {
      // Create document index item
      await this._createDocumentIndex();
      // Get file from form
      // @ts-ignore: Object is possibly 'null'.
      if ((document.querySelector(upload) as HTMLInputElement).files[0] != null) {
        // @ts-ignore: Object is possibly 'null'.
        let myfile = (document.querySelector(upload) as HTMLInputElement).files[0];
        console.log(myfile);
        var splitted = myfile.name.split(".");
        documentNameExtension = this.state.documentName + '.' + splitted[splitted.length - 1];
        this.documentNameExtension = documentNameExtension;
        docinsertname = this.state.documentid + '.' + splitted[splitted.length - 1];
        if (myfile.size) {
          // add file to source library
          const fileUploaded = await this._Service.uploadDocument(docinsertname, myfile, this.props.sourceDocumentLibrary);
          if (fileUploaded) {
            const filePath = window.location.protocol + "//" + window.location.host + fileUploaded.data.ServerRelativeUrl;
            const item = await fileUploaded.file.getItem();
            console.log(item);
            sourceDocumentId = item["ID"];
            this.setState({ sourceDocumentId: sourceDocumentId });
            // update metadata
            await this._addSourceDocument();
            if (item) {
              let revision;
              revision = "0";
              let logItems = {
                Title: this.state.documentid,
                Status: "Document Created",
                LogDate: this.today,
                Revision: revision,
                DocumentIndexId: parseInt(this.state.newDocumentId),
              }
              await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems);
              // update document index
              if (this.state.directPublishCheck === false) {
                let indexItems = {
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
                }
                // await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexItems);
                await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexItems);
              }
              else {
                let indexItems = {
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
                }
                // await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexItems);
                await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexItems);
              }
              await this._triggerPermission(sourceDocumentId);
              if (this.state.directPublishCheck === true) {
                this.setState({ hideLoading: false, hideCreateLoading: "none" });
                await this._publish();
              }
              else {
                if (this.state.sendForReview === true) {
                  this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                  this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.siteUrl);
                  }, 5000);
                }
                else {
                  this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.siteUrl);
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
          // this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary)
          this._Service.getSelectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary, "LinkFilename,ID,FileLeafRef,DocumentName")
            .then(publishdoc => {
              console.log(publishdoc);
              for (let i = 0; i < publishdoc.length; i++) {
                if (publishdoc[i].Id === this.state.templateId) {
                  publishName = publishdoc[i].DocumentName;
                }
              }
              var split = publishName.split(".", 2);
              extension = split[1];
            }).then(cpysrc => {
              // Add template document to source document
              newDocumentName = this.state.documentName + "." + extension;
              this.documentNameExtension = newDocumentName;
              docinsertname = this.state.documentid + '.' + extension;
              let filePath: string;
              this._Service.getPathOfSelectedTemplate(publishName, "SourceDocuments").then((items) => {
                if (items.length > 0) {
                  // Get the first item (assuming the file names are unique)
                  const fileItem = items[0];

                  // Access the server-relative URL of the file
                  filePath = fileItem.FileDirRef + '/' + publishName;
                  console.log(filePath)
                }
              }).then(afterPath => {
                this._Service.getBuffer(filePath).then(templateData => {
                  return this._Service.uploadDocument(docinsertname, templateData, this.props.sourceDocumentLibrary);
                }).then(fileUploaded => {
                  const filePath = window.location.protocol + "//" + window.location.host + fileUploaded.data.ServerRelativeUrl;
                  console.log("File Uploaded");
                  fileUploaded.file.getItem().then(async item => {
                    console.log(item);
                    sourceDocumentId = item["ID"];
                    this.setState({ sourceDocumentId: sourceDocumentId });
                    await this._addSourceDocument();
                  }).then(async updateDocumentIndex => {
                    let revision;
                    revision = "0";
                    let logItems = {
                      Title: this.state.documentid,
                      Status: "Document Created",
                      LogDate: this.today,
                      Revision: revision,
                      DocumentIndexId: parseInt(this.state.newDocumentId),
                    }
                    await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems);
                    if (this.state.directPublishCheck === false) {
                      let indexUpdateItems = {
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
                      }
                      // this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                      this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);

                    }
                    else {
                      let indexUpdateItems = {
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
                      }
                      // this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                      this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                    }
                    await this._triggerPermission(sourceDocumentId);
                    if (this.state.directPublishCheck === true) {
                      this.setState({ hideLoading: false, hideCreateLoading: "none" });
                      await this._publish();
                    }
                    else {
                      if (this.state.sendForReview === true) {
                        this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                        this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                        setTimeout(() => {
                          window.location.replace(this.siteUrl);
                        }, 5000);
                      }
                      else {
                        this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                        setTimeout(() => {
                          window.location.replace(this.siteUrl);
                        }, 5000);
                      }
                    }
                  });
                });
              })
              // let siteUrl = this.props.QDMSUrl + "/" + this.props.publisheddocumentLibrary + "/" + publishName;

            });
        }
        else {
          // this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary)
          this._Service.getSelectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary, "LinkFilename,ID,Template,DocumentName")
            .then(publishdoc => {
              console.log(publishdoc);
              for (let i = 0; i < publishdoc.length; i++) {
                if (publishdoc[i].Id === this.state.templateId) {
                  publishName = publishdoc[i].LinkFilename;
                }
              }
              var split = publishName.split(".", 2);
              extension = split[1];
            }).then(cpysrc => {
              // Add template document to source document
              newDocumentName = this.state.documentName + "." + extension;
              this.documentNameExtension = newDocumentName;
              docinsertname = this.state.documentid + '.' + extension;
              let siteUrl = this.props.siteUrl + "/" + this.props.publisheddocumentLibrary + "/" + this.state.category + "/" + publishName;
              this._Service.getBuffer(siteUrl).then(templateData => {
                return this._Service.uploadDocument(docinsertname, templateData, this.props.sourceDocumentLibrary);
              }).then(fileUploaded => {
                const filePath = window.location.protocol + "//" + window.location.host + fileUploaded.data.ServerRelativeUrl;
                console.log("File Uploaded");
                fileUploaded.file.getItem().then(async item => {
                  console.log(item);
                  sourceDocumentId = item["ID"];
                  this.setState({ sourceDocumentId: sourceDocumentId });
                  await this._addSourceDocument();
                }).then(async updateDocumentIndex => {
                  let revision;
                  revision = "0";
                  let logItems = {
                    Title: this.state.documentid,
                    Status: "Document Created",
                    LogDate: this.today,
                    Revision: revision,
                    DocumentIndexId: parseInt(this.state.newDocumentId),
                  }
                  await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logItems);
                  if (this.state.directPublishCheck === false) {
                    let indexUpdateItems = {
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
                    }
                    // this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                    this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);

                  }
                  else {
                    let indexUpdateItems = {
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
                    }
                    // this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                    this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.state.newDocumentId, indexUpdateItems);
                  }
                  await this._triggerPermission(sourceDocumentId);
                  if (this.state.directPublishCheck === true) {
                    this.setState({ hideLoading: false, hideCreateLoading: "none" });
                    await this._publish();
                  }
                  else {
                    if (this.state.sendForReview === true) {
                      this._triggerSendForReview(sourceDocumentId, this.state.newDocumentId);
                      this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                      setTimeout(() => {
                        window.location.replace(this.siteUrl);
                      }, 5000);
                    }
                    else {
                      this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                      setTimeout(() => {
                        window.location.replace(this.siteUrl);
                      }, 5000);
                    }
                  }
                });
              });
            });
        }

      }
      else { }
    }
    // without document
    else {
      await this._createDocumentIndex();

      this.setState({ statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 }, norefresh: "none", hideCreateLoading: "none" });

      setTimeout(() => {
        window.location.replace(this.siteUrl);
      }, this.Timeout);

    }
  }
  // Create Document Index
  public _createDocumentIndex() {
    let documentIndexId;
    // Without Expiry date
    if (this.state.expiryCheck === false) {
      let indexItems = {
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
      }
      this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, indexItems).then(async newdocid => {
        console.log(newdocid);
        this.documentIndexID = newdocid.data.ID;
        documentIndexId = newdocid.data.ID;
        this.setState({ newDocumentId: documentIndexId });
        this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
        this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
      });

    }
    // With Expiry date
    else {
      let indexItems = {
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
      }
      this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, indexItems).then(async newdocid => {
        console.log(newdocid);
        this.documentIndexID = newdocid.data.ID;
        documentIndexId = newdocid.data.ID;
        this.setState({ newDocumentId: documentIndexId });
        this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
        this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
      });
    }
  }
  // Add Source Document metadata
  public async _addSourceDocument() {
    // Without Expiry Date
    if (this.state.expiryCheck === false) {
      let sourceUpdate = {
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
      }
      // await this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId, sourceUpdate);
      await this._Service.getByIdUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId, sourceUpdate);


    }
    // With Expiry Date
    else {
      let sourceUpdate = {
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
      }
      // await this._Service.itemFromLibraryUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId, sourceUpdate);
      await this._Service.getByIdUpdate(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId, sourceUpdate);


    }
  }
  // Set permission for document
  protected async _triggerPermission(sourceDocumentID) {
    // const laUrl = await this._Service.DocumentPermission(this.props.siteUrl, this.props.requestList);
    const laUrl = await this._Service.getItemsFilter(this.props.siteUrl, this.props.requestList, "Title eq 'QDMS_DocumentPermission-Create Document'");
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
  //trigger sendForRevew
  protected async _triggerSendForReview(sourceDocumentID, documentIndexId) {
    // const laUrl = await this._Service.DocumentSendForReview(this.props.siteUrl, this.props.requestList);
    const laUrl = await this._Service.getItemsFilter(this.props.siteUrl, this.props.requestList, "Title eq 'Send For Review New DMS'");
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
  //Document Published
  protected async _publish() {
    await this._revisionCoding();
    // const laUrl = await this._Service.DocumentPublish(this.props.siteUrl, this.props.requestList);
    const laUrl = await this._Service.getItemsFilter(this.props.siteUrl, this.props.requestList, "Title eq 'QDMS_DocumentPublish'");
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
      'SourceDocumentLibrary': this.props.sourceDocumentLibrary,
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
      this._publishUpdate();
    }
    else { }
  }
  // qdms revision
  public _revisionCoding = async () => {
    let revision = parseInt("0");
    let rev = revision + 1;
    this.setState({ newRevision: rev.toString() });

  }
  // Published Document Metadata update
  public async _publishUpdate() {

    await this._Service.itemFromLibraryByID(this.props.siteUrl, this.props.sourceDocumentLibrary, this.state.sourceDocumentId);
    let itemToUpdate = {
      PublishFormat: this.state.publishOption,
      WorkflowStatus: "Published",
      Revision: this.state.newRevision,
      ApprovedDate: new Date()
    }
    await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemToUpdate);
    // await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, itemToUpdate);

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

    this.setState({ hideLoading: true, norefresh: "none", hideCreateLoading: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.directPublish, messageType: 4 } });
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
    let link;

    console.log(this.state.criticalDocument);
    // const notificationPreference: any[] = await this._Service.itemFromPrefernce(this.props.siteUrl, this.props.notificationPreference, emailuser);
    const notificationPreference: any[] = await this._Service.getSelectFilter(this.props.siteUrl, this.props.notificationPreference, "Preference", "EmailUser/EMail eq '" + emailuser + "'");
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
    if (mailSend === "Yes") {
      const emailNotification: any[] = await this._Service.getItems(this.props.siteUrl, this.props.emailNotification);
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title === type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }
      }
      let linkValue = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2";
      // link = `<a href=${window.location.protocol + "//" + window.location.hostname+this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"}>`+this.state.documentName+`</a>`;
      link = `<a href=${linkValue}>` + this.state.documentName + `</a>`;
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      let replaceRequester = replaceString(Body, '[Sir/Madam]', name);
      let replaceDate = replaceString(replaceRequester, '[PublishedDate]', day);
      let replaceApprover = replaceString(replaceDate, '[Approver]', this.state.approverName);
      let replaceBody = replaceString(replaceApprover, '[DocumentName]', this.state.documentName);
      let replacelink = replaceString(replaceBody, '[DocumentLink]', link);
      let FinalBody = replacelink;
      //Create Body for Email  
      let emailPostBody: any = {
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
  //For dialog box of cancel
  private _dialogCloseButton = () => {
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
  // On format date
  private _onFormatDate = (date: Date): string => {

    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
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
  private async _onSendForReview() {
    if (this.state.createDocument === true && this.isDocument === "Yes" || this.state.createDocument === false) {
      if (this.state.expiryCheck === true) {
        //Validation without direct publish
        if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
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
        else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('SubCategory') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {

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
        if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {

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
        else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {

          if (this.isDocument === "Yes") {
            this.setState({
              showReviewModal: true,
            });
          } else {
            this.setState({ statusMessage: { isShowMessage: true, message: "Please select document", messageType: 1 } });
            setTimeout(() => {
              this.setState({ statusMessage: { isShowMessage: false, message: "Please select document", messageType: 4 } });
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


  }
  public onConfirmReview = async () => {
    if (this.state.DueDate !== null) {
      await this.setState({
        sendForReview: true,
        showReviewModal: false,
        dueDateMadatory: "",
        saveDisable: true, hideCreateLoading: " ",
        norefresh: " "
      });
      this._documentidgeneration();
    }
    else {
      this.setState({ dueDateMadatory: "Yes" });
    }

  }

  public render(): React.ReactElement<ICreateDocumentProps> {

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

    const calloutProps = { gapSpace: 0 };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const uploadOrTemplateRadioBtnOptions:
      IChoiceGroupOption[] = [
        { key: 'Upload', text: 'Upload existing file' },
        { key: 'Template', text: 'Create document using existing template', styles: { field: { marginLeft: "35px" } } },
      ];
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
      <section className={`${styles.createDocument}`}>

        <div className={styles.border}>
          <div className={styles.alignCenter}>{this.props.webpartHeader}</div>
          <div>
            <TextField required id="t1"
              label="Title"
              onChange={this._titleChange}
              value={this.state.title} ></TextField>
            <div style={{ color: "#dc3545" }}>
              {this.validator.message("Title", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
          </div>

          <div className={styles.divrow}>
            <div className={styles.divColumn1}>
              <Dropdown id="t3" label="Category"
                selectedKey={this.state.departmentId}
                placeholder="Select an option"
                defaultSelectedKey={this.state.departmentId}
                required
                //disabled={this.state.departmentId !== ""}
                options={this.state.departmentOption}
                onChanged={this._departmentChange} />
              <div style={{ color: "#dc3545", textAlign: "center" }}>
                {this.validator.message("BU/Dep", this.state.businessUnitID || this.state.departmentId, "required")}{""}
              </div>
            </div>
            <div className={styles.divColumn2}>
              <Dropdown id="t2" required={true} label="Doc Category"
                placeholder="Select an option"
                selectedKey={this.state.categoryId}
                options={this.state.categoryOption}
                onChanged={this._categoryChange} />
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("category", this.state.categoryId, "required")}{" "}</div>
            </div>
            <div className={styles.divColumn2}>
              <Dropdown id="t2" required={true} label="Doc Type"
                placeholder="Select an option"
                selectedKey={this.state.subCategoryId}
                options={this.state.subCategoryArray}
                onChanged={this._subCategoryChange} />
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("subCategory", this.state.subCategoryId, "required")}{" "}</div>
            </div>
          </div>

          <div className={styles.documentMainDiv}>
            <div className={styles.radioDiv} style={{ display: this.state.hideDoc }}>
              <ChoiceGroup selectedKey={this.state.uploadOrTemplateRadioBtn}
                onChange={this.onUploadOrTemplateRadioBtnChange}
                options={uploadOrTemplateRadioBtnOptions} styles={choiceGroupStyles}
              />
            </div>
            <div className={styles.uploadDiv} style={{ display: this.state.hideupload }}>
              <div ><input type="file" name="myFile" id="addqdms" onChange={this._add}></input></div>
              <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select  document </div>
            </div>
            <div className={styles.templateDiv} style={{ display: this.state.hidetemplate }}>
              <div className={styles.divColumn2} style={{ display: "flex" }}>
                {this.props.siteUrl !== "/sites/Quality" &&
                  <div className={styles.divColumn2}>
                    <Dropdown id="t7"
                      label="Source"
                      placeholder="Select an option"
                      selectedKey={this.state.sourceId}
                      options={Source}
                      onChanged={this._sourcechange} />
                  </div>
                }
                <div className={styles.divColumn2} style={{ maxWidth: (this.props.siteUrl === "/sites/Quality") ? "26.8rem" : "163.8rem" }}>
                  <Dropdown id="t7"
                    label="Select a Template"
                    placeholder="Select an option"
                    selectedKey={this.state.templateId}
                    options={this.state.templateDocuments}
                    onChanged={this._templatechange} style={{ width: "150%", }} />
                </div>
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
            <div style={{ width: "77%" }}>

              <div style={{ color: "#dc3545" }}>
                {this.validator.message("Owner", this.state.owner, "required")}{" "}</div>
            </div>
            <div style={{ width: "75%", marginLeft: "10px" }}>

              <div style={{ color: "#dc3545" }}>
              </div>
            </div>
            <div className={styles.divApprover}>

              <div style={{ display: this.state.validApprover, color: "#dc3545" }}>Not able to change approver</div>
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
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
          <div className={styles.divrow}>
            <div className={styles.wdthmid} style={{ display: this.state.checkdirect }}>
              <Spinner label={'Please Wait...'} /></div>
          </div>
          <div> {this.state.statusMessage.isShowMessage ?
            <MessageBar
              messageBarType={this.state.statusMessage.messageType}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >{this.state.statusMessage.message}</MessageBar>
            : ''} </div>
          <div className={styles.mt}>
            <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
          </div>
          <div className={styles.mt}>
            <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Creating...'} /></div>
          </div>
          <div className={styles.mt}>
            <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
              <Label>***PLEASE DON'T REFRESH***</Label>
            </div>
          </div>
          <div className={styles.divrow}>
            <div style={{ fontStyle: "italic", fontSize: "12px", position: "absolute" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
            <div className={styles.rgtalign} >
              <PrimaryButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._onSendForReview}>Send for review and submit</PrimaryButton >
              <PrimaryButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._onCreateDocument}>Submit</PrimaryButton >
              <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
            </div>
          </div>
          {/* {/ Cancel Dialog Box /} */}
          <div style={{ display: this.state.cancelConfirmMsg }}>
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
                <TextField id="comments" autoComplete='true' label="Comments" onChange={this._commentChange} value={this.state.comments} multiline ></TextField>
                <PrimaryButton style={{ float: "right", marginTop: "7px", marginBottom: "9px" }} className={styles.modalButton} id="b2" onClick={this.onConfirmReview} >Confirm</PrimaryButton >
              </div>

            </Modal>
          </div>
          <br />
        </div>
      </section>
    );
  }
}
