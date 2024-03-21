import * as React from 'react';
import styles from './SendRequest.module.scss';
import type { ISendRequestProps, ISendRequestState } from '../interfaces';
import { DatePicker, DefaultButton, Dialog, DialogFooter, DialogType,Label, Link,MessageBar, PrimaryButton, ProgressIndicator, Spinner, TextField } from '@fluentui/react';
import SimpleReactValidator from 'simple-react-validator';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as moment from 'moment';
import { IHttpClientOptions, HttpClient } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import * as _ from 'lodash';
import { DMSService } from '../services';

export default class SendRequest extends React.Component<ISendRequestProps, ISendRequestState> {

  private _Service: DMSService;
  private validator: SimpleReactValidator;
  private siteUrl;
  private documentIndexID;
  private invalidUser;
  private currentEmail;
  private currentId;
  private today;
  //private time;
  private workflowStatus;
  private sourceDocumentID;
  private newheaderid;
  private newDetailItemID;
  private dccReview;
  private underApproval;
  private underReview;
  private redirectUrl = this.props.siteUrl + this.props.redirectUrl;
  private invalidSendRequestLink;
  private getSelectedReviewers: any[] = [];
  //private valid;
  private noDocument;
  private taskDelegate = "No";
  private taskDelegateDccReview;
  private taskDelegateUnderApproval;
  private taskDelegateUnderReview;
  //private departmentExists;
  private postUrl;
  private postUrlForUnderReview;
  private permissionpostUrl;
  private postUrlForAdaptive;
  private TaskID;
  public constructor(props: ISendRequestProps) {
    super(props);
    this.state = {

      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      documentID: "",
      linkToDoc: "",
      documentName: "",
      revision: "",
      ownerName: "",
      currentUser: "",
      hideProject: true,
      revisionLevel: [],
      revisionLevelvalue: "",
      dcc: "",
      reviewer: "",
      dueDate: "",
      approver: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      saveDisable: "",
      requestSend: 'none',
      statusKey: "",
      access: "none",
      accessDeniedMsgBar: "none",
      reviewers: [],
      ownerId: "",
      delegatedToId: "",
      delegateToIdInSubSite: "",
      delegateForIdInSubSite: "",
      reviewerEmail: "",
      reviewerName: "",
      delegatedFromId: "",
      detailIdForReviewer: "",
      approverEmail: "",
      approverName: "",
      hubSiteUserId: 0,
      detailIdForApprover: "",
      criticalDocument: "",
      dccReviewerName: "",
      dccReviewerEmail: "",
      dccReviewer: "",
      revisionLevelArray: [],
      revisionCoding: "",
      currentUserReviewer: [],
      projectName: "",
      projectNumber: "",
      acceptanceCodeId: "",
      transmittalRevision: "",
      reviewersName: [],
      hideLoading: true,
      sameRevision: false,
      loaderDisplay: "",
      businessUnitID: null,
      departmentId: null,
      validApprover: "none",
      hideCreateLoading: "none",
      

    };
    this._Service = new DMSService(this.props.context);
    this.componentDidMount = this.componentDidMount.bind(this);
    this._userMessageSettings = this._userMessageSettings.bind(this);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._checkWorkflowStatus = this._checkWorkflowStatus.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
    this._bindSendRequestForm = this._bindSendRequestForm.bind(this);
    this._dccReviewerChange = this._dccReviewerChange.bind(this);
    this._reviewerChange = this._reviewerChange.bind(this);
    this._approverChange = this._approverChange.bind(this);
    this._submitSendRequest = this._submitSendRequest.bind(this);
    this._underApprove = this._underApprove.bind(this);
    this._underReview = this._underReview.bind(this);
    this._onSameRevisionChecked = this._onSameRevisionChecked.bind(this);
    this._adaptiveCard = this._adaptiveCard.bind(this);
    this._LaUrlGettingAdaptive = this._LaUrlGettingAdaptive.bind(this);
  }

  public UNSAFE_componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "This field is mandatory"
      }
    });
  }
  //Page Load
  public async componentDidMount() {

    this.redirectUrl = this.props.redirectUrl;
    this.setState({ access: "none", accessDeniedMsgBar: "none" });
    // Get User Messages
    await this._userMessageSettings();
    //Get Current User
    const user = await this._Service.getCurrentUser();
    this.currentEmail = user.Email;
    this.currentId = user.Id;
    //Get Parameter from URL
    this._queryParamGetting();

    // if (this.props.project) {
    //   this.setState({ hideProject: false });
    // }

    const currentUserReviewer: any[] = [];
    currentUserReviewer.push(this.currentId);
    //Get Today
    this.today = new Date();
    this.setState({
      currentUser: user.Title,
      currentUserReviewer: currentUserReviewer
    });
    this.siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl
  }
  
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.getSelectFilter(this.props.siteUrl, this.props.userMessageSettings, "Title,Message", "PageName eq 'SendRequest'");
    // const userMessageSettings: any[] = await this._Service.getItemsFromUserMsgSettings(this.props.siteUrl, this.props.userMessageSettings);

    for (const i in userMessageSettings) {
      if (userMessageSettings[i].Title === "InvalidSendRequestUser") {
        this.invalidUser = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title === "InvalidSendRequestLink") {
        this.invalidSendRequestLink = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title === "NoDocument") {
        this.noDocument = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title === "WorkflowStatusError") {
        this.workflowStatus = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title === "DccReview") {
        const DccReview = userMessageSettings[i].Message;
        this.dccReview = replaceString(DccReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title === "UnderApproval") {
        const UnderApproval = userMessageSettings[i].Message;
        this.underApproval = replaceString(UnderApproval, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title === "UnderReview") {
        const UnderReview = userMessageSettings[i].Message;
        this.underReview = replaceString(UnderReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title === "TaskDelegateDccReview") {
        const TaskDelegateDccReview = userMessageSettings[i].Message;
        this.taskDelegateDccReview = replaceString(TaskDelegateDccReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title === "TaskDelegateUnderApproval") {
        const TaskDelegateUnderApproval = userMessageSettings[i].Message;
        this.taskDelegateUnderApproval = replaceString(TaskDelegateUnderApproval, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title === "TaskDelegateUnderReview") {
        const TaskDelegateUnderReview = userMessageSettings[i].Message;
        this.taskDelegateUnderReview = replaceString(TaskDelegateUnderReview, '[DocumentName]', this.state.documentName);

      }
    }

  }
  //Get Parameter from URL
  private async _queryParamGetting() {
    //Query getting...
    const params = new URLSearchParams(window.location.search);
    const documentindexid = params.get('did');

    if (documentindexid !== "" && documentindexid !== null) {
      this.documentIndexID = parseInt(documentindexid);
      //Get Access
      this.setState({ access: "none", accessDeniedMsgBar: "none" });
      // if (this.props.project) {
        // await this._checkWorkflowStatus();
        // this._checkPermission('Project_SendRequest');
      // }
      // else {
        // await this._accessGroups();
        
        await this._checkWorkflowStatus();

      // }
    }
    else {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidSendRequestLink, messageType: 1 },
      });
      setTimeout(() => {
        window.location.replace(this.redirectUrl);
      }, 10000);
    }
  }
  //Workflow Status Checking
  private async _checkWorkflowStatus() {
    const documentIndexItem: any = await this._Service.getByIdSelect(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, "WorkflowStatus,SourceDocument,DocumentStatus");
    // const documentIndexItem: any = await this._Service.getWorkflowStatus(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
    //getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).select("WorkflowStatus,SourceDocument,DocumentStatus").get();
    if (documentIndexItem.WorkflowStatus === "Under Review" || documentIndexItem.WorkflowStatus === "Under Approval") {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.workflowStatus, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
    else if (documentIndexItem.DocumentStatus !== "Active") {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: "Document is not currently Active", messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
    else if (documentIndexItem.SourceDocument === null) {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        access: "none",
        statusMessage: { isShowMessage: true, message: this.noDocument, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
    else {
      this.setState({ access: "", accessDeniedMsgBar: "none", loaderDisplay: "none" });
      await this._bindSendRequestForm();
    }

  }
  //Bind Send Request Form
  public async _bindSendRequestForm() {
    this._Service.getItemsByID(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID).then(async indexItems => {
      //this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).get().then(async indexItems => {

      // const documentID;
      // const documentName;
      // const ownerName;
      // const ownerId;
      // const revision;
      // const linkToDocument;
      // const criticalDocument;
      // const approverName;
      // const approverId;
      // const approverEmail;
      const  temReviewersID: any[] = [];
      const tempReviewers: any[] = [];
      // const businessUnitID;
      // const departmentId;
      //Get Document Index
      const documentIndexItem: any = await this._Service.getByIdSelectExpand(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, "DocumentID,DocumentName,DepartmentID,Owner/ID,Owner/Title,Owner/EMail,Approver/ID,Approver/Title,Approver/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID,Reviewers/ID,Reviewers/Title,Reviewers/EMail", "Owner,Approver,Reviewers");
      // const documentIndexItem: any = await this._Service.getByIdSelectExpand(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, "DocumentID,DocumentName,DepartmentID,BusinessUnitID,Owner/ID,Owner/Title,Owner/EMail,Approver/ID,Approver/Title,Approver/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID,Reviewers/ID,Reviewers/Title,Reviewers/EMail", "Owner,Approver,Reviewers");
      // const documentIndexItem: any = await this._Service.getDocumentIndexItem(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
      //const documentIndexItem: any = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).select("DocumentID,DocumentName,DepartmentID,BusinessUnitID,Owner/ID,Owner/Title,Owner/EMail,Approver/ID,Approver/Title,Approver/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID,Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Owner,Approver,Reviewers").get();

      const documentID = documentIndexItem.DocumentID;
      const documentName = documentIndexItem.DocumentName;
      const ownerName = documentIndexItem.Owner.Title;
      const ownerId = documentIndexItem.Owner.ID;
      const revision = documentIndexItem.Revision;
      const linkToDocument = documentIndexItem.SourceDocument.Url;
      // this.SourceDocumentID = DocumentIndexItem.SourceDocumentID;
      const criticalDocument = documentIndexItem.CriticalDocument;
      // const approverName = documentIndexItem.Approver.Title;
      // const approverId = documentIndexItem.Approver.ID;
      // const approverEmail = documentIndexItem.Approver.EMail;
      // const businessUnitID = documentIndexItem.BusinessUnitID;
      const departmentId = documentIndexItem.DepartmentID;
      for (const k in documentIndexItem.Reviewers) {
        temReviewersID.push(documentIndexItem.Reviewers[k].ID);
        this.setState({
          reviewers: temReviewersID,
        });
        tempReviewers.push(documentIndexItem.Reviewers[k].Title);
      }
      if (indexItems.ApproverId !== null) {
        this.setState({
          approver: documentIndexItem.Approver.ID,
          approverName: documentIndexItem.Approver.Title,
          approverEmail: documentIndexItem.Approver.EMail
        });
      }
      this.setState({
        documentID: documentID,
        documentName: documentName,
        ownerName: ownerName,
        ownerId: ownerId,
        revision: revision,
        linkToDoc: linkToDocument,
        criticalDocument: criticalDocument,
        // approver: approverId,
        // approverName: approverName,
        reviewersName: tempReviewers,
        // businessUnitID: businessUnitID,
        departmentId: departmentId
      });
      // const sourceDocumentItem: any = await this._Service.getSourceDocumentItem(this.props.siteUrl, this.props.sourceDocumentLibrary, this.documentIndexID);
      const sourceDocumentItem: any = await this._Service.getLibraryFilter(this.props.siteUrl, this.props.sourceDocumentLibrary, 'DocumentIndexId eq ' + this.documentIndexID);
      //const sourceDocumentItem: any = await this._Service.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.filter('DocumentIndexId eq ' + this.documentIndexID).get();

      this.sourceDocumentID = sourceDocumentItem[0].ID;
      await this._userMessageSettings();
    });
  }
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID);
  }
  //Critical Document Change
  public _onSameRevisionChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ sameRevision: true }); }
    else if (!isChecked) { this.setState({ sameRevision: false }); }
  }
  // on dccreviewer change
  public _dccReviewerChange = (items: any[]) => {
    this.setState({ saveDisable: "" });
    let dccreviewerEmail;
    let dccreviewerName;

    const getSelecteddccreviewer: any[] = [];
    for (let item in items) {
      dccreviewerEmail = items[item].secondaryText;
        dccreviewerName = items[item].text;
        getSelecteddccreviewer.push(items[item].id);
    }
    this.setState({
      dccReviewer: getSelecteddccreviewer[0],
      dccReviewerEmail: dccreviewerEmail,
      dccReviewerName: dccreviewerName
    });
  }
  // on reviewer change
  public _reviewerChange = (items: any[]) => {
    this.setState({ saveDisable: "" });
    this.getSelectedReviewers = [];
    for (const item in items) {
      this.getSelectedReviewers.push(items[item].id);
    }
    this.setState({ reviewers: this.getSelectedReviewers });
  }
  // on approver change
  public _approverChange = async (items: any[]) => {
    this.setState({ saveDisable: "" });
    let approverEmail;
    let approverName;
    const getSelectedApprover: any[] = [];
    // if (this.props.project) {
    //   for (let item in items) {
    //     approverEmail = items[item].secondaryText,
    //       approverName = items[item].text,
    //       getSelectedApprover.push(items[item].id);
    //   }
    //   this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName });
    // }
    // else {
      this.setState({ validApprover: "", approver: null, approverEmail: "", approverName: "", });
      if (this.state.businessUnitID !== null) {
        const businessUnit = await this._Service.getSelectExpand(this.props.siteUrl, this.props.businessUnitList, "ID,Title,Approver/Title,Approver/ID,Approver/EMail", "Approver");
        // const businessUnit = await this._Service.getBusinessUnit(this.props.siteUrl, this.props.businessUnitList);
        //const businessUnit = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.businessUnitList).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver").get();
        for (let i = 0; i < businessUnit.length; i++) {
          if (businessUnit[i].ID === this.state.businessUnitID) {
            const approve = await this._Service.getUserByEmail(businessUnit[i].Approver.EMail);
            //const approve = await this._Service.siteUsers.getByEmail(businessUnit[i].Approver.EMail).get();
            approverEmail = businessUnit[i].Approver.EMail;
            approverName = businessUnit[i].Approver.Title;
            getSelectedApprover.push(approve.Id);
          }
        }
      }
      else {
        const departments = await this._Service.getSelectExpand(this.props.siteUrl, this.props.departmentList, "ID,Title,Approver/Title,Approver/ID,Approver/EMail", "Approver");
        // const departments = await this._Service.getDepartments(this.props.siteUrl, this.props.departmentList);
        //const departments = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.departmentList).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver").get();
        for (let i = 0; i < departments.length; i++) {
          if (departments[i].ID === this.state.departmentId) {
            const deptapprove = await this._Service.getUserByEmail(departments[i].Approver.EMail);
            //const deptapprove = await this._Service.siteUsers.getByEmail(departments[i].Approver.EMail).get();
            approverEmail = departments[i].Approver.EMail;
            approverName = departments[i].Approver.Title;
            getSelectedApprover.push(deptapprove.Id);
          }
        }
      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName });
      setTimeout(() => {
        this.setState({ validApprover: "none" });
      }, 5000);
    // }
  }
  // on expirydate change
  private _onExpDatePickerChange = (date?: Date): void => {
    this.setState({ saveDisable: "" });
    this.setState({ dueDate: date });
  }
  // on format date
  private _onFormatDate = (date: Date): string => {
    const selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  };
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, saveDisable: "" });
  }
  // on submit send request
  private _submitSendRequest = async () => {
    //this.setState({ saveDisable: true, hideLoading: false });
    let sorted_previousHeaderItems: any[] = [];
    let previousHeaderItem = 0;
    const previousHeaderItems = await this._Service.getSelectFilter(this.props.siteUrl, this.props.workflowHeaderList, "ID", "DocumentIndex eq '" + this.documentIndexID + "' and(WorkflowStatus eq 'Returned with comments')");
    // const previousHeaderItems = await this._Service.getPreviousHeaderItems(this.props.siteUrl, this.props.workflowHeaderList, this.documentIndexID);
    //const previousHeaderItems = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.select("ID").filter("DocumentIndex eq '" + this.documentIndexID + "' and(WorkflowStatus eq 'Returned with comments')").get();
    if (previousHeaderItems.length !== 0) {
      sorted_previousHeaderItems = _.orderBy(previousHeaderItems, 'ID', ['desc']);
      previousHeaderItem = sorted_previousHeaderItems[0].ID;
    }
    if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate")) {
      if (this.state.reviewers.length === 0) {
        this.setState({ saveDisable: "none", hideLoading: false });
        this._underApprove(previousHeaderItem);
        
      }
      else {
        this.setState({ saveDisable: "none", hideLoading: false });
        this._underReview(previousHeaderItem);
      }
      this.validator.hideMessages();
      this.setState({ requestSend: "" });
      setTimeout(() => this.setState({ requestSend: 'none', saveDisable: "none" }), 3000);
    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }

  }
  //  qdms request to review
  public async _underReview(previousHeaderItem) {
    await this._LAUrlGettingForUnderReview();
    // this._LaUrlGettingAdaptive();
    const itemtobeadded = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: this.state.reviewers,
      ApproverId: this.state.approver,
      OwnerId: this.state.ownerId,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      
      PreviousReviewHeader: previousHeaderItem.toString()
    }
    // const header = await this._Service.addToWorkflowHeaderList(this.props.siteUrl, this.props.workflowHeaderList, itemtobeadded);
    const header = await this._Service.addItem(this.props.siteUrl, this.props.workflowHeaderList, itemtobeadded);
    /* const header = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: this.state.reviewers ,
      ApproverId: this.state.approver,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }); */
    if (header) {
      this.newheaderid = header.data.ID;
      const revisionitem = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }
      // const log = await this._Service.addToDocumentRevision(this.props.siteUrl, this.props.documentRevisionLogList, revisionitem);
      await this._Service.addItem(this.props.siteUrl, this.props.documentRevisionLogList, revisionitem);
      /* const log = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.add({
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }); */
      //for reviewers if exist
      for (let k = 0; k < this.state.reviewers.length; k++) {
        const user = await this._Service.getSiteUserById(this.state.reviewers[k]);
        //const user = await this._Service.siteUsers.getById(this.state.reviewers[k]).get();
        if (user) {
          // const hubsieUser = await this._Service.getUserByEmail(user.Email);
          //const hubsieUser = await this._Service.siteUsers.getByEmail(user.Email).get();
          // if (hubsieUser) {
            //Task delegation 
            // const taskDelegation: any[] = await this._Service.getSelectExpandFilter(this.props.siteUrl, this.props.taskDelegationSettings, "DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate", "DelegatedFor,DelegatedTo", "DelegatedFor/ID eq '" + hubsieUser.Id + "' and(Status eq 'Active')")
            // const taskDelegation: any[] = await this._Service.getTaskDelegation(this.props.siteUrl, this.props.taskDelegationSettings, hubsieUser.Id)
            //const taskDelegation: any[] = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + hubsieUser.Id + "' and(Status eq 'Active')").get();

            // if (taskDelegation.length > 0) {
            //   let duedate = moment(this.state.dueDate).toDate();
            //   let toDate = moment(taskDelegation[0].ToDate).toDate();
            //   let fromDate = moment(taskDelegation[0].FromDate).toDate();
            //   duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
            //   toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
            //   fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
            //   if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
            //     this.taskDelegate = "Yes";
            //     this.setState({
            //       approverEmail: taskDelegation[0].DelegatedTo.EMail,
            //       approverName: taskDelegation[0].DelegatedTo.Title,
            //       delegatedToId: taskDelegation[0].DelegatedTo.ID,
            //       delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            //     });
            //     //Get Delegated To ID
            //     const DelegatedTo = await this._Service.getUserByEmail(taskDelegation[0].DelegatedTo.EMail);
            //     //const DelegatedTo = await this._Service.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get();
            //     if (DelegatedTo) {
            //       this.setState({
            //         delegateToIdInSubSite: DelegatedTo.Id,
            //       });
            //       //Get Delegated For ID
            //       const DelegatedFor = await this._Service.getUserByEmail(taskDelegation[0].DelegatedFor.EMail);
            //       //const DelegatedFor = await this._Service.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get();
            //       if (DelegatedFor) {
            //         this.setState({
            //           delegateForIdInSubSite: DelegatedFor.Id,
            //         });
            //         //detail list adding an item for reviewers
            //         const index = this.state.reviewers.indexOf(DelegatedFor.Id);

            //         this.state.reviewers[index] = DelegatedTo.Id;

            //         const detailitem = {
            //           HeaderIDId: Number(this.newheaderid),
            //           Workflow: "Review",
            //           Title: this.state.documentName,
            //           ResponsibleId: (this.state.delegatedToId !== "" ? DelegatedTo.Id : user.Id),
            //           DueDate: this.state.dueDate,
            //           DelegatedFromId: (this.state.delegatedToId !== "" ? DelegatedFor.Id : parseInt("")),
            //           ResponseStatus: "Under Review",
            //           SourceDocument: {
            //             Description: this.state.documentName,
            //             Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            //           },
            //           OwnerId: this.state.ownerId,
            //         }
            //         const detail = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, detailitem);
            //         // const detail = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, detailitem);
            //         /* const detail = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add({
            //           HeaderIDId: Number(this.newheaderid),
            //           Workflow: "Review",
            //           Title: this.state.documentName,
            //           ResponsibleId: (this.state.delegatedToId !== "" ? DelegatedTo.Id : user.Id),
            //           DueDate: this.state.dueDate,
            //           DelegatedFromId: (this.state.delegatedToId !== "" ? DelegatedFor.Id : parseInt("")),
            //           ResponseStatus: "Under Review",
            //           SourceDocument: {
            //             Description: this.state.documentName,
            //             Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            //           },
            //           OwnerId: this.state.ownerId,
            //         }); */
            //         if (detail) {
            //           this.setState({ detailIdForApprover: detail.data.ID });
            //           this.newDetailItemID = detail.data.ID;
            //           const updateitem = {
            //             Link: {
            //               Description: this.state.documentName + "-Review",
            //               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
            //             },
            //           }
            //           // await this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, updateitem)
            //           await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, updateitem)
            //           /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detail.data.ID).update({
            //             Link: {
            //               Description: this.state.documentName + "-Review",
            //               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
            //             },
            //           }); *///Update link

            //           //MY tasks list updation with delegated from
            //           const taskitem = {
            //             Title: "Review '" + this.state.documentName + "'",
            //             Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
            //             DueDate: this.state.dueDate,
            //             StartDate: this.today,
            //             AssignedToId: (this.state.delegatedToId !== "" ? this.state.delegatedToId : hubsieUser.Id),
            //             Priority: (this.state.criticalDocument === true ? "Critical" : ""),
            //             DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
            //             // Source: (this.props.project ? "Project" : "QDMS"),
            //             Source: "QDMS",
            //             DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegatedFromId : parseInt("")),
            //             Workflow: "Review",
            //             Link: {
            //               Description: this.state.documentName + "-Review",
            //               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
            //             },
            //           }
            //           // const task = await this._Service.addToWorkflowTasksList(this.props.siteUrl, this.props.workflowTasksList, taskitem);
            //           const task = await this._Service.addItem(this.props.siteUrl, this.props.workflowTasksList, taskitem);
            //           /* const task = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTasksList).items.add({
            //             Title: "Review '" + this.state.documentName + "'",
            //             Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
            //             DueDate: this.state.dueDate,
            //             StartDate: this.today,
            //             AssignedToId: (this.state.delegatedToId !== "" ? this.state.delegatedToId : hubsieUser.Id),
            //             Priority: (this.state.criticalDocument === true ? "Critical" : ""),
            //             DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
            //             Source: (this.props.project ? "Project" : "QDMS"),
            //             DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegatedFromId : parseInt("")),
            //             Workflow: "Review",
            //             Link: {
            //               Description: this.state.documentName + "-Review",
            //               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
            //             },
            //           }); */
            //           if (task) {
            //             this.TaskID = task.data.ID;
            //             const dataitem = {
            //               TaskID: task.data.ID,
            //             }
            //             // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, dataitem)
            //             await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, dataitem)
            //             /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detail.data.ID).update
            //               ({
            //                 TaskID: task.data.ID,
            //               }); */
            //             await this._sendmail(DelegatedTo.Email, "DocReview", DelegatedTo.Title);
            //             await this._adaptiveCard("Review", DelegatedTo.Email, DelegatedTo.Title, "General");
            //           }//taskID
            //         }//r
            //       }//Delegated For
            //     }//Delegated To
            //   }
            //   else {
            //     //detail list adding an item for reviewers
            //     const dataitem = {
            //       HeaderIDId: Number(this.newheaderid),
            //       Workflow: "Review",
            //       Title: this.state.documentName,
            //       ResponsibleId: user.Id,
            //       DueDate: this.state.dueDate,
            //       ResponseStatus: "Under Review",
            //       SourceDocument: {
            //         Description: this.state.documentName,
            //         Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            //       },
            //       OwnerId: this.state.ownerId,
            //     }
            //     // const details = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, dataitem);
            //     const details = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, dataitem);
            //     /* const details = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add({
            //       HeaderIDId: Number(this.newheaderid),
            //       Workflow: "Review",
            //       Title: this.state.documentName,
            //       ResponsibleId: user.Id,
            //       DueDate: this.state.dueDate,
            //       ResponseStatus: "Under Review",
            //       SourceDocument: {
            //         Description: this.state.documentName,
            //         Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            //       },
            //       OwnerId: this.state.ownerId,
            //     }); */
            //     if (details) {
            //       this.setState({ detailIdForApprover: details.data.ID });
            //       this.newDetailItemID = details.data.ID;
            //       const dataitem = {
            //         Link: {
            //           Description: this.state.documentName + "-Review",
            //           Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
            //         },
            //       }
            //       // const updatedetail = await this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, dataitem)
            //       const updatedetail = await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, dataitem)
            //       /* const updatedetail = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(details.data.ID).update({
            //         Link: {
            //           Description: this.state.documentName + "-Review",
            //           Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
            //         },
            //       }); */
            //       //MY tasks list updation with delegated from
            //       const taskitem = {
            //         Title: "Review '" + this.state.documentName + "'",
            //         Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
            //         DueDate: this.state.dueDate,
            //         StartDate: this.today,
            //         AssignedToId: hubsieUser.Id,
            //         Priority: (this.state.criticalDocument === true ? "Critical" : ""),
            //         // Source: (this.props.project ? "Project" : "QDMS"),
            //         Source: "QDMS",
            //         Workflow: "Review",
            //         Link: {
            //           Description: this.state.documentName + "-Review",
            //           Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
            //         },
            //       }
            //       // const task = await this._Service.addToWorkflowTasksList(this.props.siteUrl, this.props.workflowTasksList, taskitem)
            //       const task = await this._Service.addItem(this.props.siteUrl, this.props.workflowTasksList, taskitem)
            //       /* const task = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTasksList).items.add({
            //         Title: "Review '" + this.state.documentName + "'",
            //         Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
            //         DueDate: this.state.dueDate,
            //         StartDate: this.today,
            //         AssignedToId: hubsieUser.Id,
            //         Priority: (this.state.criticalDocument === true ? "Critical" : ""),
            //         Source: (this.props.project ? "Project" : "QDMS"),
            //         Workflow: "Review",
            //         Link: {
            //           Description: this.state.documentName + "-Review",
            //           Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
            //         },
            //       }); */
            //       if (task) {
            //         const dataitem = {
            //           TaskID: task.data.ID,
            //         }
            //         const updatetask = await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, dataitem)
            //         // const updatetask = await this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, dataitem)
            //         /* const updatetask = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(details.data.ID).update
            //           ({
            //             TaskID: task.data.ID,
            //           }); */
            //         if (updatetask) {
            //           // await this._sendmail(user.Email, "DocReview", user.Title);
            //           // await this._adaptiveCard("Review", user.Email, user.Title, "General");
            //         }
            //       }//taskId
            //     }//r
            //   }//else

            // }//IF
            //If no task delegation
            // else {
              // alert("no task delegation")
              //detail list adding an item for reviewers
              const detailitem = {
                HeaderIDId: Number(this.newheaderid),
                Workflow: "Review",
                Title: this.state.documentName,
                ResponsibleId: user.Id,
                DueDate: this.state.dueDate,
                ResponseStatus: "Under Review",
                SourceDocument: {
                  Description: this.state.documentName,
                  Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                },
                OwnerId: this.state.ownerId,
              }
              // const details = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, detailitem)
              const details = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, detailitem)
              /* const details = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add({
                HeaderIDId: Number(this.newheaderid),
                Workflow: "Review",
                Title: this.state.documentName,
                ResponsibleId: user.Id,
                DueDate: this.state.dueDate,
                ResponseStatus: "Under Review",
                SourceDocument: {
                  Description: this.state.documentName,
                  Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                },
                OwnerId: this.state.ownerId,
              }); */
              if (details) {
                this.setState({ detailIdForApprover: details.data.ID });
                this.newDetailItemID = details.data.ID;
                const detailsdata = {
                  Link: {
                    Description: this.state.documentName + "-Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                  },
                }
                await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, detailsdata)
                // const updatedetail = await this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, detailsdata)
                /* const updatedetail = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(details.data.ID).update({
                  Link: {
                    Description: this.state.documentName + "-Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                  },
                }); */
                //MY tasks list updation with delegated from
                const taskitem = {
                  Title: "Review '" + this.state.documentName + "'",
                  Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: user.Id,
                  Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                  // Source: (this.props.project ? "Project" : "QDMS"),
                  Source: "QDMS",
                  Workflow: "Review",
                  Link: {
                    Description: this.state.documentName + "-Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                  },
                }
                // const task = await this._Service.addToWorkflowTasksList(this.props.siteUrl, this.props.workflowTasksList, taskitem);
                const task = await this._Service.addItem(this.props.siteUrl, this.props.workflowTasksList, taskitem);
                /* const task = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTasksList).items.add({
                  Title: "Review '" + this.state.documentName + "'",
                  Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: hubsieUser.Id,
                  Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                  Source: (this.props.project ? "Project" : "QDMS"),
                  Workflow: "Review",
                  Link: {
                    Description: this.state.documentName + "-Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                  },
                }); */
                if (task) {
                  this.TaskID = task.data.ID;
                  const detailitem = {
                    TaskID: task.data.ID,
                  }
                  const updatetask = await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, detailitem);
                  // const updatetask = await this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, details.data.ID, detailitem);
                  /* const updatetask = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(details.data.ID).update
                    ({
                      TaskID: task.data.ID,
                    }); */
                  if (updatetask) {
                    await this._sendmail(user.Email, "DocReview", user.Title);
                    // await this._adaptiveCard("Review", user.Email, user.Title, "General");
                  }
                }//taskId
              }//r
              // this.setState({ hideCreateLoading: "none", statusMessage: { isShowMessage: true, message: this.underApproval, messageType: 4 } });
              // setTimeout(() => {
              //   window.location.replace(this.siteUrl);
              // }, 3000);


            // }//else
          // }//hubsiteuser
        }//user
      }
      const indexitem = {
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ReviewersId: this.state.reviewers,
      ApproverId: this.state.approver,
      OwnerId: this.state.ownerId,
        WorkflowDueDate: this.state.dueDate
      }
      // await this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, indexitem);
      await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, indexitem);
      /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers ,
        WorkflowDueDate: this.state.dueDate
      }); */
      const sourceitem = {
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ReviewersId: this.state.reviewers,
        ApproverId: this.state.approver,
        OwnerId: this.state.ownerId,
      }
      // await this._Service.updateItemById(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourceitem);
      await this._Service.getByIdUpdateSourceLibrary(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourceitem);
      /* await this._Service.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
      }); */
      const headeritem = {
        ReviewersId: this.state.reviewers
      }
      // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderList, parseInt(this.newheaderid), headeritem);
      await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowHeaderList, parseInt(this.newheaderid), headeritem);
      /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(parseInt(this.newheaderid)).update({
        ReviewersId: this.state.reviewers
      }); */
      const flowtrigger = await this._triggerDocumentUnderReview(this.sourceDocumentID, "Review");
      const logitem = {
        Title: this.state.documentID,
        Status: "Under Review",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Review",
        DueDate: this.state.dueDate,
      }
      const logupdate = await this._Service.addItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem);
      // await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem);
      /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.add({
        Title: this.state.documentID,
        Status: "Under Review",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Review",
        DueDate: this.state.dueDate,
      }); */
      const final = [await flowtrigger,await logupdate];
      if(final){
      if (this.taskDelegate === "Yes") {
        this.setState({
          hideLoading: true,
          statusMessage: { isShowMessage: true, message: this.taskDelegateUnderReview, messageType: 4 },
        });
      }
      else {
        //alert("hi");
        this.setState({
          hideLoading: true,
          saveDisable: "none",
          statusMessage: { isShowMessage: true, message: this.underReview, messageType: 4 },
        });
      }
      // setTimeout(() => {
      //   window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      // }, 10000);

      this.setState({ hideCreateLoading: "none", statusMessage: { isShowMessage: true, message: this.underReview, messageType: 4 } });
                    setTimeout(() => {
                      window.location.replace(this.siteUrl);
                    }, 3000);
                  }
    }
  }
  // la for under review permission
  private _LAUrlGettingForUnderReview = async () => {
    const laUrl = await this._Service.getFilter(this.props.siteUrl, this.props.requestList, "Title eq 'QDMS_DocumentPermission_UnderReview'");
    // const laUrl = await this._Service.getUnderReview(this.props.siteUrl, this.props.requestList);
    //const laUrl = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.requestList).items.filter("Title eq 'QDMS_DocumentPermission_UnderReview'").get();

    this.postUrlForUnderReview = laUrl[0].PostUrl;
  }

  // set permission for reviewer
  protected async _triggerDocumentUnderReview(sourceDocumentID, type) {
    const siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrlForUnderReview;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': "Under Review",
      'Workflow': type
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //  qdms request to approve
  public async _underApprove(previousHeaderItem) {
    this._LAUrlGetting();
    // this._LaUrlGettingAdaptive();
    const headeritem = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      ReviewedDate: this.today,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: this.state.currentUserReviewer,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }
    // const header = await this._Service.addToWorkflowHeaderList(this.props.siteUrl, this.props.workflowHeaderList, headeritem);
    const header = await this._Service.addItem(this.props.siteUrl, this.props.workflowHeaderList, headeritem);
    /* const header = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      ReviewedDate: this.today,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: this.state.currentUserReviewer,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }); */
    if (header) {
      this.newheaderid = header.data.ID;

      const logitem = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,

      }
      await this._Service.addItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem);
      // const log = await this._Service.addToDocumentRevision(this.props.siteUrl, this.props.documentRevisionLogList, logitem);
      /* const log = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.add({
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,

      }); */
      const detailitem = {
        HeaderIDId: Number(this.newheaderid),
        Workflow: "Review",
        Title: this.state.documentName,
        ResponsibleId: this.currentId,
        DueDate: this.state.dueDate,
        ResponseStatus: "Reviewed",
        ResponseDate: this.today,
        SourceDocument: {
          Description: this.state.documentName,
          Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
        },
        OwnerId: this.state.ownerId,
      }
      // const detailadd = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, detailitem);
      const detailadd = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, detailitem);
      /* const detailadd = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add({
        HeaderIDId: Number(this.newheaderid),
        Workflow: "Review",
        Title: this.state.documentName,
        ResponsibleId: this.currentId,
        DueDate: this.state.dueDate,
        ResponseStatus: "Reviewed",
        ResponseDate: this.today,
        SourceDocument: {
          Description: this.state.documentName,
          Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
        },
        OwnerId: this.state.ownerId,
      }); */
      if (detailadd) {
        this.setState({ detailIdForApprover: detailadd.data.ID });
        this.newDetailItemID = detailadd.data.ID;
        const detailitem = {
          Link: {
             Description: this.state.documentName + "-Review",
            Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailadd.data.ID + ""
          },
        }
        await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detailadd.data.ID, detailitem);
        // await this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, detailadd.data.ID, detailitem);
        /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detailadd.data.ID).update({
          Link: {
            Description: this.state.documentName + "-Review",
            Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailadd.data.ID + ""
          },
        }); */
      }

      // this.setState({ hideCreateLoading: "none", statusMessage: { isShowMessage: true, message: this.underApproval, messageType: 4 } });
      //               setTimeout(() => {
      //                 window.location.replace(this.siteUrl);
      //               }, 3000);


      //Task delegation getting user id from hubsite
      const user = await this._Service.getUserByEmail(this.state.approverEmail);
      //const user = await this._Service.siteUsers.getByEmail(this.state.approverEmail).get();
      if (user) {
        // this.setState({
        //   hubSiteUserId: user.Id,
        // });

        //Task delegation 
        // const taskDelegation: any[] = await this._Service.getSelectExpandFilter(this.props.siteUrl, this.props.taskDelegationSettings, "DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate", "DelegatedFor,DelegatedTo", "DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')");
        // const taskDelegation: any[] = await this._Service.getTaskDelegation(this.props.siteUrl, this.props.taskDelegationSettings, user.Id);
        /* const taskDelegation: any[] = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationSettings).items
        .select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate")
        .expand("DelegatedFor,DelegatedTo")
        .filter("DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')").get(); */

      //   if (taskDelegation.length > 0) {
      //     let duedate = moment(this.state.dueDate).toDate();
      //     let toDate = moment(taskDelegation[0].ToDate).toDate();
      //     let fromDate = moment(taskDelegation[0].FromDate).toDate();
      //     duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
      //     toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
      //     fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
      //     if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
      //       this.taskDelegate = "Yes";
      //       this.setState({
      //         approverEmail: taskDelegation[0].DelegatedTo.EMail,
      //         approverName: taskDelegation[0].DelegatedTo.Title,

      //         delegatedToId: taskDelegation[0].DelegatedTo.ID,
      //         delegatedFromId: taskDelegation[0].DelegatedFor.ID,
      //       });
      //       //detail list adding an item for approval
      //       const DelegatedTo = await this._Service.getUserByEmail(taskDelegation[0].DelegatedTo.EMail);
      //       //const DelegatedTo = await this._Service.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get();
      //       if (DelegatedTo) {
      //         this.setState({
      //           delegateToIdInSubSite: DelegatedTo.Id,
      //         });
      //         const DelegatedFor = await this._Service.getUserByEmail(taskDelegation[0].DelegatedFor.EMail);
      //         //const DelegatedFor = await this._Service.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get();
      //         if (DelegatedFor) {
      //           this.setState({
      //             delegateForIdInSubSite: DelegatedFor.Id,
      //           });
      //           const itemdetail = {
      //             HeaderIDId: Number(this.newheaderid),
      //             Workflow: "Approval",
      //             Title: this.state.documentName,
      //             ResponsibleId: (this.state.delegatedToId !== "" ? this.state.delegateToIdInSubSite : this.state.approver),
      //             DueDate: this.state.dueDate,
      //             DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegateForIdInSubSite : parseInt("")),
      //             ResponseStatus: "Under Approval",
      //             SourceDocument: {
      //               Description: this.state.documentName,
      //               Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
      //             },
      //             OwnerId: this.state.ownerId,
      //           }
      //           const detailsAdd = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, itemdetail);
      //           // const detailsAdd = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, itemdetail);
      //           /* const detailsAdd = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add
      //             ({
      //               HeaderIDId: Number(this.newheaderid),
      //               Workflow: "Approval",
      //               Title: this.state.documentName,
      //               ResponsibleId: (this.state.delegatedToId !== "" ? this.state.delegateToIdInSubSite : this.state.approver),
      //               DueDate: this.state.dueDate,
      //               DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegateForIdInSubSite : parseInt("")),
      //               ResponseStatus: "Under Approval",
      //               SourceDocument: {
      //                 Description: this.state.documentName,
      //                 Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
      //               },
      //               OwnerId: this.state.ownerId,
      //             }); */
      //           if (detailsAdd) {
      //             this.setState({ detailIdForApprover: detailsAdd.data.ID });
      //             this.newDetailItemID = detailsAdd.data.ID;
      //             const itemtoupdate = {
      //               Link: {
      //                 Description: this.state.documentName + "-Approve",
      //                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailsAdd.data.ID + ""
      //               },
      //             }
      //             await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detailsAdd.data.ID, itemtoupdate)
      //             // await this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, detailsAdd.data.ID, itemtoupdate)
      //             /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detailsAdd.data.ID).update({
      //               Link: {
      //                 Description: this.state.documentName + "-Approve",
      //                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailsAdd.data.ID + ""
      //               },
      //             }); */
      //             const initem = {
      //               ApproverId: this.state.delegateToIdInSubSite,
      //             }
      //             // await this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, initem);
      //             await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, initem);
      //             /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
      //               ApproverId: this.state.delegateToIdInSubSite,
      //             }); */
      //             const sourceitem = {
      //               ApproverId: this.state.delegateToIdInSubSite,
      //             }
      //             await this._Service.getByIdUpdateSourceLibrary(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourceitem);
      //             // await this._Service.updateItemById(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourceitem);
      //             /* await this._Service.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
      //               ApproverId: this.state.delegateToIdInSubSite,
      //             }); */
      //             const headeritem = {
      //               ApproverId: this.state.delegateToIdInSubSite,
      //             }
      //             await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowHeaderList, this.newheaderid, headeritem);
      //             // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderList, this.newheaderid, headeritem);
      //             /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.newheaderid).update({
      //               ApproverId: this.state.delegateToIdInSubSite,
      //             }); */
      //             //MY tasks list updation
      //             const taskitem = {
      //               Title: "Approve '" + this.state.documentName + "'",
      //               Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
      //               DueDate: this.state.dueDate,
      //               StartDate: this.today,
      //               AssignedToId: (this.state.delegatedToId),
      //               Workflow: "Approval",
      //               // Priority:(this.state.criticalDocument === true ? "Critical" :""),
      //               DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
      //               // Source: (this.props.project ? "Project" : "QDMS"),
      //               Source: "QDMS",
      //               DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegatedFromId : 0),
      //               Link: {
      //                 Description: this.state.documentName + "-Approve",
      //                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailsAdd.data.ID + ""
      //               },

      //             }
      //             // const task = await this._Service.addToWorkflowTasksList(this.props.siteUrl, this.props.workflowTasksList, taskitem);
      //             const task = await this._Service.addItem(this.props.siteUrl, this.props.workflowTasksList, taskitem);
      //             /* const task = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTasksList).items.add
      //               ({
      //                 Title: "Approve '" + this.state.documentName + "'",
      //                 Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
      //                 DueDate: this.state.dueDate,
      //                 StartDate: this.today,
      //                 AssignedToId: (this.state.delegatedToId),
      //                 Workflow: "Approval",
      //                 // Priority:(this.state.criticalDocument === true ? "Critical" :""),
      //                 DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
      //                 Source: (this.props.project ? "Project" : "QDMS"),
      //                 DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegatedFromId : 0),
      //                 Link: {
      //                   Description: this.state.documentName + "-Approve",
      //                   Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailsAdd.data.ID + ""
      //                 },

      //               }); */
      //             if (task) {
      //               const taskitem = {
      //                 TaskID: task.data.ID,
      //               }
      //               this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detailsAdd.data.ID, taskitem);
      //               // this._Service.updateWorkflowDetailsList(this.props.siteUrl, this.props.workflowDetailsList, detailsAdd.data.ID, taskitem);
      //               /* this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detailsAdd.data.ID).update
      //                 ({
      //                   TaskID: task.data.ID,
      //                 }); */
      //               //notification preference checking                                 
      //               await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
      //               await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");

      //             }//taskID
      //           }//r

      //         }//DelegatedFor
      //       }//DelegatedTo
      //     }//duedate checking
      //     else {
      //       const detailitem = {
      //         HeaderIDId: Number(this.newheaderid),
      //         Workflow: "Approval",
      //         Title: this.state.documentName,
      //         ResponsibleId: this.state.approver,
      //         DueDate: this.state.dueDate,
      //         ResponseStatus: "Under Approval",
      //         SourceDocument: {
      //           Description: this.state.documentName,
      //           Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
      //         },
      //         OwnerId: this.state.ownerId,
      //       }
      //       // const detail = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, detailitem);
      //       const detail = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, detailitem);
      //       /* const detail = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add
      //         ({
      //           HeaderIDId: Number(this.newheaderid),
      //           Workflow: "Approval",
      //           Title: this.state.documentName,
      //           ResponsibleId: this.state.approver,
      //           DueDate: this.state.dueDate,
      //           ResponseStatus: "Under Approval",
      //           SourceDocument: {
      //             Description: this.state.documentName,
      //             Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
      //           },
      //           OwnerId: this.state.ownerId,
      //         }); */
      //       if (detail) {
      //         this.setState({ detailIdForApprover: detail.data.ID });
      //         this.newDetailItemID = detail.data.ID;
      //         const detailitem = {
      //           Link: {
      //             Description: this.state.documentName + "-Approve",
      //             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
      //           },
      //         }
      //         await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detailitem)
      //         // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detailitem)
      //         /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detail.data.ID).update({
      //           Link: {
      //             Description: this.state.documentName + "-Approve",
      //             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
      //           },
      //         }); */
      //         const inditem = {
      //           ApproverId: this.state.approver,
      //         }
      //         await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, inditem)
      //         // await this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, inditem)
      //         /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
      //           ApproverId: this.state.approver,
      //         }); */
      //         const sourceitem = {
      // ApproverId: this.state.approver
      //         }
      //         await this._Service.getByIdUpdateSourceLibrary(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourceitem)
      //         // await this._Service.updateItemById(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourceitem)
      //         /* await this._Service.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
      //           ApproverId: this.state.approver,

      //         }); */
      //         const headeritem = {
      //           ApproverId: this.state.approver,
      //         }
      //         await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowHeaderList, this.newheaderid, headeritem)
      //         // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderList, this.newheaderid, headeritem)
      //         /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.newheaderid).update({
      //           ApproverId: this.state.approver,
      //         }); */
      //         //MY tasks list updation
      //         const taskitem =
      //         {
      //           Title: "Approve '" + this.state.documentName + "'",
      //           Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
      //           DueDate: this.state.dueDate,
      //           StartDate: this.today,
      //           AssignedToId: user.Id,
      //           Workflow: "Approval",
      //           Priority: (this.state.criticalDocument === true ? "Critical" : ""),
      //           // Source: (this.props.project ? "Project" : "QDMS"),
      //           Source: "QDMS",
      //           Link: {
      //             Description: this.state.documentName + "-Approve",
      //             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
      //           },
      //         }
      //         // const task = await this._Service.addToWorkflowTasksList(this.props.siteUrl, this.props.workflowTasksList, taskitem)
      //         const task = await this._Service.addItem(this.props.siteUrl, this.props.workflowTasksList, taskitem)
      //         /* const task = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTasksList).items.add
      //           ({
      //             Title: "Approve '" + this.state.documentName + "'",
      //             Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
      //             DueDate: this.state.dueDate,
      //             StartDate: this.today,
      //             AssignedToId: user.Id,
      //             Workflow: "Approval",
      //             Priority: (this.state.criticalDocument === true ? "Critical" : ""),
      //             Source: (this.props.project ? "Project" : "QDMS"),
      //             Link: {
      //               Description: this.state.documentName + "-Approve",
      //               Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
      //             },

      //           }); */
      //         if (task) {
      //           const detailitem = {
      //             TaskID: task.data.ID,
      //           }
      //           await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detailitem)
      //           // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detailitem)
      //           /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detail.data.ID).update
      //             ({
      //               TaskID: task.data.ID,
      //             }); */
      //           //notification preference checking                                 
      //           await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
      //           await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
      //         }//taskID
      //       }//r
      //     }
      //   }
      //   else {
          const detitem = {
            HeaderIDId: Number(this.newheaderid),
            Workflow: "Approval",
            Title: this.state.documentName,
            ResponsibleId: this.state.approver,
            DueDate: this.state.dueDate,
            ResponseStatus: "Under Approval",
            SourceDocument: {
              Description: this.state.documentName,
              Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            },
            OwnerId: this.state.ownerId,
          }
          // const detail = await this._Service.addToWorkflowDetail(this.props.siteUrl, this.props.workflowDetailsList, detitem)
          const detail = await this._Service.addItem(this.props.siteUrl, this.props.workflowDetailsList, detitem)
          /* const detail = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.add
            ({
              HeaderIDId: Number(this.newheaderid),
              Workflow: "Approval",
              Title: this.state.documentName,
              ResponsibleId: this.state.approver,
              DueDate: this.state.dueDate,
              ResponseStatus: "Under Approval",
              SourceDocument: {
                Description: this.state.documentName,
                Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
              },
              OwnerId: this.state.ownerId,
            }); */
          if (detail) {
            this.setState({ detailIdForApprover: detail.data.ID });
            this.newDetailItemID = detail.data.ID;
            const detaildata = {
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              },
            }
            await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detaildata)
            // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detaildata)
            /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detail.data.ID).update({
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              },
            }); */
            const inddata = {
              ApproverId: this.state.approver,
            }
            await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, inddata)
            // await this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, inddata)
            /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
              ApproverId: this.state.approver,
            }); */
            const sourcedata = {
              ApproverId: this.state.approver,
              WorkflowStatus: "Under Review",
        Workflow: "Review",
        OwnerId: this.state.ownerId,
            }
            await this._Service.getByIdUpdateSourceLibrary(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourcedata)
            // await this._Service.updateItemById(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourcedata)
            /* await this._Service.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
              ApproverId: this.state.approver,

            }); */
            const headdata = {
              ApproverId: this.state.approver,
            }
            await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowHeaderList, this.newheaderid, headdata)
            // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderList, this.newheaderid, headdata)
            /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.newheaderid).update({
              ApproverId: this.state.approver,
            }); */
            //MY tasks list updation
            const taskdata = {
              Title: "Approve '" + this.state.documentName + "'",
              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
              DueDate: this.state.dueDate,
              StartDate: this.today,
              AssignedToId: user.Id,
              Workflow: "Approval",
              Priority: (this.state.criticalDocument === true ? "Critical" : ""),
              // Source: (this.props.project ? "Project" : "QDMS"),
              Source: "QDMS",
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              },

            }
            // const task = await this._Service.addToWorkflowTasksList(this.props.siteUrl, this.props.workflowTasksList, taskdata)
            const task = await this._Service.addItem(this.props.siteUrl, this.props.workflowTasksList, taskdata)
            /* const task = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTasksList).items.add
              ({
                Title: "Approve '" + this.state.documentName + "'",
                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                DueDate: this.state.dueDate,
                StartDate: this.today,
                AssignedToId: user.Id,
                Workflow: "Approval",
                Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                Source: (this.props.project ? "Project" : "QDMS"),
                Link: {
                  Description: this.state.documentName + "-Approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                },

              }); */
            if (task) {
              this.TaskID = task.data.ID
              const detaildata = {
                TaskID: task.data.ID,
              }
              await this._Service.getByIdUpdate(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detaildata)
              // await this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsList, detail.data.ID, detaildata)
              /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(detail.data.ID).update
                ({
                  TaskID: task.data.ID,
                }); */
              //notification preference checking                                 
              await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
              // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");

            }//taskID
          }//r
        // }//else no delegation
      }
      const inddata = {
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: this.state.currentUserReviewer,
        WorkflowDueDate: this.state.dueDate
      }
      // await this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, inddata)
      await this._Service.getByIdUpdate(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID, inddata)
      /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: this.state.currentUserReviewer,
        WorkflowDueDate: this.state.dueDate
      }); */
      const sourcedata = {
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: this.state.currentUserReviewer,
      }
      await this._Service.getByIdUpdateSourceLibrary(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourcedata)
      // await this._Service.updateItemById(this.props.siteUrl, this.props.sourceDocumentLibrary, this.sourceDocumentID, sourcedata)
      /* await this._Service.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: this.state.currentUserReviewer,
      }); */
      const datarevision = {
        Title: this.state.documentID,
        Status: "Under Approval",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Approval",
        DueDate: this.state.dueDate,
      }
      await this._Service.addItem(this.props.siteUrl, this.props.documentRevisionLogList, datarevision)
      // this.setState({ hideCreateLoading: "none", statusMessage: { isShowMessage: true, message: this.underApproval, messageType: 4 } });
      //               setTimeout(() => {
      //                 window.location.replace(this.siteUrl);
      //               }, 3000);
      // await this._Service.addToDocumentRevision(this.props.siteUrl, this.props.documentRevisionLogList, datarevision)
      /* await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.add({
        Title: this.state.documentID,
        Status: "Under Approval",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Approval",
        DueDate: this.state.dueDate,
      }); */
     const permissiontrigger = await this._triggerPermission(this.sourceDocumentID);
     const completed = [await permissiontrigger]
     if(completed){
      this.setState({
        comments: "",
        statusKey: "",
        approverEmail: "",
        approverName: "",
        approver: "",
      });
      if (this.taskDelegate === "Yes") {
        this.setState({
          hideLoading: true,
          statusMessage: { isShowMessage: true, message: this.taskDelegateUnderApproval, messageType: 4 },
        });
        setTimeout(() => {
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
        }, 10000);
      }
      else {
        this.setState({
          hideLoading: true,
          saveDisable: "none",
          statusMessage: { isShowMessage: true, message: this.underApproval, messageType: 4 },
        });
        setTimeout(() => {
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
        }, 10000);
      }
    }   //msg
    }//newheaderid
  }
  // La for under approval permission
  private async _LAUrlGetting() {
    // const laUrl = await this._Service.getUnderApproval(this.props.siteUrl, this.props.requestList);
    const laUrl = await this._Service.getFilter(this.props.siteUrl, this.props.requestList, "Title eq 'QDMS_DocumentPermission_UnderApproval'");
    //const laUrl = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.requestList).items.filter("Title eq 'QDMS_DocumentPermission_UnderApproval'").get();
    this.postUrl = laUrl[0].PostUrl;
  }
  // set permission for approver
  private async _triggerPermission(sourceDocumentID) {
    const siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': "Under Approval"
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  private _LaUrlGettingAdaptive = async () => {
    const laUrl: any[] = await this._Service.getItems(this.props.siteUrl, this.props.requestList);
    //const laUrl: any[] = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.requestList).items.get();
    for (let i = 0; i < laUrl.length; i++) {
      if (laUrl[i].Title === "Adaptive _Card") {
        this.postUrlForAdaptive = laUrl[i].PostUrl;
      }
    }
  }
  public _adaptiveCard = async (Workflow, Email, Name, Type) => {
    const siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrlForAdaptive;
    const splitted = this.state.documentName.split(".");
    const ext = splitted[splitted.length - 1];
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'Workflow': Workflow,
      'DocumentIndexID': String(this.documentIndexID),
      'SourceDocumentID': String(this.sourceDocumentID),
      'HeaderID': String(this.newheaderid),
      'DetailID': String(this.newDetailItemID),
      'TaskID': String(this.TaskID),
      'ext': ext,
      'Email': Email,
      'Name': Name,
      'Type': Type
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //Send Mail
  public _sendmail = async (emailuser, type, name) => {
    let mailSend = "No";
    let Subject;
    let Body;
    let link;
    // const notificationPreference: any[] = await this._Service.getSelectFilter(this.props.siteUrl, this.props.notificationPreference, "Preference", "EmailUser/EMail eq '" + emailuser + "'");
    // const notificationPreference: any[] = await this._Service.getMailPreference(this.props.siteUrl, this.props.notificationPreference, emailuser);
    //const notificationPreference: any[] = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.notificationPreference).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference").get();

    // if (notificationPreference.length > 0) {
    //   if (notificationPreference[0].Preference === "Send all emails") {
    //     mailSend = "Yes";
    //   }
    //   else if (notificationPreference[0].Preference === "Send mail for critical document" && this.state.criticalDocument === true) {
    //     mailSend = "Yes";
    //   }
    //   else {
    //     mailSend = "No";
    //   }
    // }
    // else if (this.state.criticalDocument === true) {
    //   mailSend = "Yes";
    // }
    if (mailSend === "Yes") {
      const emailNotification: any[] = await this._Service.getFilter(this.props.siteUrl, this.props.emailNotification, "Title eq '" + type + "'")
      // const emailNotification: any[] = await this._Service.getEmailNotification(this.props.siteUrl, this.props.emailNotification, type)
      //const emailNotification: any[] = await this._Service.getList(this.props.siteUrl + "/Lists/" + this.props.emailNotification).items.filter("Title eq '" + type + "'").get();
      Subject = emailNotification[0].Subject;
      Body = emailNotification[0].Body;
      if (type === "DocApproval") {
        link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + this.newDetailItemID}>Link</a>`;
      }
      else if (type === "DocDCCReview") {
        link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + this.newDetailItemID + "&wf=dcc"} >Link</a>`;
      }
      else {
        link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + this.newDetailItemID}>Link</a>`;
      }
      //Replacing the email body with current values
      const dueDateformail = moment(this.state.dueDate).format("DD/MM/YYYY");
      const replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      const replacedSubjectWithDueDate = replaceString(replacedSubject, '[DueDate]', dueDateformail);
      const replaceRequester = replaceString(Body, '[Sir/Madam],', name);
      const replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
      const replacelink = replaceString(replaceBody, '[Link]', link);
       const FinalBody = replacelink;
      //Create Body for Email  
      const emailPostBody: any = {
        "message": {
          "subject": replacedSubjectWithDueDate,
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
      await this._Service.sendMail(emailPostBody);
      /* this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody);
        }); */
    }
  }
  // on cancel
  private _onCancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });
  }
  //Cancel confirm
  private _confirmYesCancel = () => {
    this.setState({
      statusKey: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
    window.location.replace(this.props.siteUrl);
  }
  //Not Cancel
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    //this.validator.hideMessages();
    // window.location.replace(this.RedirectUrl);
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
    //subText: '<b>Do you want to cancel? </b> ',
  };
  private modalProps = {
    isBlocking: true,
  };
  public render(): React.ReactElement<ISendRequestProps> {

    return (
      <section className={`${styles.sendRequest}`}>
        <div>
          <div style={{ display: this.state.loaderDisplay }}>
            <ProgressIndicator label="Loading......" />
          </div>
          <div style={{ display: this.state.access }}>
            <div className={styles.border}>
              <div className={styles.alignCenter}> {this.props.webpartHeader}</div>
              <br/>
              <div className={styles.header}>
                <div className={styles.divMetadataCol1}>
                  <h3 >Document Details</h3>
                  <Link onClick={this._openRevisionHistory} target="_blank" underline style={{ marginLeft: "70%" }}>Revision History</Link>
                </div>
              </div>
              <div className={styles.divMetadata}>
                <div className={styles.divMetadataCol1}>
                  <Label >Document ID : </Label><div className={styles.divLabel}>{this.state.documentID}</div>
                </div>
                <div className={styles.divMetadataCol3}>
                  <Label >Revision :</Label><div className={styles.divLabel}> {this.state.revision}</div>
                </div>
              </div>
              <div className={styles.divRow}>
                <Label >Document :</Label><div className={styles.divLabel}>  <a href={this.state.linkToDoc} target="_blank">{this.state.documentName}</a></div>
              </div>
              <div className={styles.header}>
                <h3 className="ExampleCard-title title-222">Workflow Details</h3>
              </div>
              <div className={styles.divMetadata}>
                <div className={styles.divMetadataCol1}>
                  <Label >Owner : </Label><div className={styles.divLabel}> {this.state.ownerName}</div>
                </div>
                <div className={styles.divMetadataCol2}>
                  <Label >Requester :</Label> <div className={styles.divLabel}>{this.state.currentUser}</div>
                </div>
              </div>

              <div className={styles.divrow}>
            <div style={{ width: "30%" }}>
                  <PeoplePicker
                    context={this.props.context as any}
                    titleText="Reviewer(s)"
                    personSelectionLimit={5}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    disabled={false}
                    ensureUser={true}
                    //selectedItems={(items) => this._reviewerChange(items)}
                    defaultSelectedUsers={this.state.reviewersName}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  /></div>
                <div  style={{ width: "30%", marginLeft: "1em",marginRight:"1em" }}>
                  <PeoplePicker
                    context={this.props.context as any}
                    titleText="Approver *"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    disabled={false}
                    ensureUser={true}
                    //selectedItems={(items) => this._approverChange(items)}
                    defaultSelectedUsers={[this.state.approverName]}
                    showHiddenInUI={false}
                    //isRequired={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  />
                  <div style={{ display: this.state.validApprover, color: "#dc3545" }}>Not able to change approver</div>
                  <div style={{ color: "#dc3545" }}>{this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
                </div>
                <div style={{ width: "30%" }}>
                  <DatePicker label="Due Date *:" id="DueDate"
                    onSelectDate={this._onExpDatePickerChange}
                    placeholder="Select a date..."
                    value={this.state.dueDate}
                    minDate={new Date()}
                    formatDate={this._onFormatDate}
                  /><div style={{ color: "#dc3545" }}>{this.validator.message("DueDate", this.state.dueDate, "required")}{" "}</div>
                </div>
              </div>
              <div className={styles.mt}>
                < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline autoAdjustHeight/></div>
              <div> {this.state.statusMessage.isShowMessage ?
                <MessageBar
                  messageBarType={this.state.statusMessage.messageType}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >{this.state.statusMessage.message}</MessageBar>
                : ''} </div>
              <div className={styles.mt}>
                <div hidden={this.state.hideLoading}><Spinner label={'Document is Sending...'} /></div>
              </div>
              <div className={styles.divRow}>
                <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                <div className={styles.rgtalign} >
                  <PrimaryButton id="b2" className={styles.btn} onClick={this._submitSendRequest} style={{ display: this.state.saveDisable }}>Submit</PrimaryButton >
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
            </div>
          </div>
          <div style={{ display: this.state.accessDeniedMsgBar }}>
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
