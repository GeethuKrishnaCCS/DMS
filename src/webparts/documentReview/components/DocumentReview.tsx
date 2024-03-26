import * as React from 'react';
import styles from './DocumentReview.module.scss';
import type { IDocumentReviewProps, IDocumentReviewState } from '../interfaces';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient, HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import { Accordion, AccordionItem } from 'react-light-accordion';
import 'react-light-accordion/demo/css/index.css';
import { ProgressIndicator, Label, Link, Dropdown, TextField, DialogFooter, MessageBar, PrimaryButton, Dialog, DefaultButton, MessageBarType, IDropdownOption, DialogType } from '@fluentui/react';
import SimpleReactValidator from 'simple-react-validator';
import * as moment from 'moment';
import { DMSService } from '../services';
import { add } from 'lodash';
import * as strings from 'DocumentReviewWebPartStrings';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export default class DocumentReview extends React.Component<IDocumentReviewProps, IDocumentReviewState> {

  private validator: SimpleReactValidator;
  private _Service: DMSService;
  constructor(props: IDocumentReviewProps) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      currentUser: 0,
      status: "",
      statusKey: "",
      comments: "",
      reviewerItems: [],
      access: "",
      accessDeniedMsgBar: "none",
      documentIndexItems: [],
      documentID: "",
      linkToDoc: "",
      documentName: "",
      revision: "",
      owner: "",
      requestor: "",
      requestorComment: "",
      dueDate: "",
      DueDate: new Date(),
      requestorDate: "",
      workflowStatus: "",
      hideReviewersTable: "none",
      detailListID: 0,
      cancelConfirmMsg: "none",
      confirmDialog: true,
      approverEmail: "",
      requestorEmail: "",
      documentControllerEmail: "",
      notificationPreference: "",
      headerListItem: [],
      approverName: "",
      approverId: "",
      ownerEmail: "",
      ownerID: "",
      reviewPending: "No",
      criticalDocument: false,
      currentUserEmail: "",
      userMessageSettings: [],
      invalidMessage: "",
      pageLoadItems: [],
      buttonHidden: "",
      detailIdForApprover: "",
      hubSiteUserId: "",
      delegatedToId: "",
      delegatedFromId: "",
      divForDCC: "none",
      divForReview: "none",
      ifDccComment: "none",
      dcc: "",
      dccComment: "",
      dccCompletionDate: "",
      revisionLogID: "",
      delegateToIdInSubSite: "",
      delegateForIdInSubSite: "",
      noAccess: "",
      invalidQueryParam: "",
      projectName: "",
      projectNumber: "",
      hideproject: true,
      reviewers: [],
      dccReviewItems: [],
      currentReviewComment: "",
      currentReviewItems: [],
      loaderDisplay: "none",
      documentControllerName: ""
    };
    this._Service = new DMSService(this.props.context);
    this._drpdwnStatus = this._drpdwnStatus.bind(this);
    this._onPageLoadDataBind = this._onPageLoadDataBind.bind(this);
    this._currentUser = this._currentUser.bind(this);
    this._loadPreviousReturnWithComments = this._loadPreviousReturnWithComments.bind(this);
    this._docReviewSaveAsDraft = this._docReviewSaveAsDraft.bind(this);
    this._docReviewSubmit = this._docReviewSubmit.bind(this);
    this._cancel = this._cancel.bind(this);
    this._confirmNoCancel = this._confirmNoCancel.bind(this);
    this._confirmYesCancel = this._confirmYesCancel.bind(this);
    this._sendAnEmailUsingMSGraph = this._sendAnEmailUsingMSGraph.bind(this);
    this._checkingReviewStatus = this._checkingReviewStatus.bind(this);
    this._returnWithComments = this._returnWithComments.bind(this);
    this._userMessageSettings = this._userMessageSettings.bind(this);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._documentIndexListBind = this._documentIndexListBind.bind(this);
    //this._docDCCReviewSubmit = this._docDCCReviewSubmit.bind(this);
    this._revisionLogChecking = this._revisionLogChecking.bind(this);
    this._accessGroups = this._accessGroups.bind(this);
    this._projectInformation = this._projectInformation.bind(this);
    this._checkingCurrent = this._checkingCurrent.bind(this);
    this.GetGroupMembers = this.GetGroupMembers.bind(this);
    this._gettingGroupID = this._gettingGroupID.bind(this);
    this._LAUrlGetting = this._LAUrlGetting.bind(this);
    this._LAUrlGettingForUnderReview = this._LAUrlGettingForUnderReview.bind(this);
    this.triggerDocumentReview = this.triggerDocumentReview.bind(this);
    this._LAUrlGettingForPermission = this._LAUrlGettingForPermission.bind(this);
    this.triggerProjectPermissionFlow = this.triggerProjectPermissionFlow.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
  }
  private headerId;
  private documentIndexId;
  private status;
  // private reqWeb = Web(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.hubsite);
  private documentReviewedSuccess;
  private documentSavedAsDraft;
  private detailID;
  private sourceDocumentID;
  private taskID;
  private newDetailItemID;
  private revisionLogID;
  // private RevisionHistoryUrl;
  private RedirectUrl;
  private valid = "ok";
  private noAccess;
  private currentDate = new Date();
  private workFlow;
  private departmentExist;
  private postUrl;
  private postUrlForUnderReview;
  private postUrlForPermission;
  private dueDateWithoutConversion;
  private postUrlForAdaptive;
  public componentWillMount = async () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });
  }
  //Page Load
  public async componentDidMount() {
    console.log(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.hubsite);
    // this._userMessageSettings();
     this._currentUser();
    this._queryParamGetting();
    //Get Approver
    const headerItem: any = await this._Service.getByIdSelect(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, "DocumentIndexID")
    //const headerItem: any = await this._Service.getDocumentIndexID(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId)
    //const headerItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).select("DocumentIndexID").get();
    this.documentIndexId = headerItem.DocumentIndexID;
    //Permission handiling 
    // await this._accessGroups();
    // await this._LAUrlGettingForPermission();
     this._LAUrlGetting();
    // this._LAUrlGettingForUnderReview();
    // console.log('this.state.currentReviewItems: ', this.state.currentReviewItems);
    // this._checkingReviewStatus();
  }
  //user message settings..
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.getItemSelectFilter(this.props.siteUrl, this.props.userMessageSettings, "Title,Message", "PageName eq 'Review'")
    //const userMessageSettings: any[] = await this._Service.getUserMessageForReview(this.props.siteUrl, this.props.userMessageSettings);
    //const userMessageSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.userMessageSettings).items.select("Title,Message").filter("PageName eq 'Review'").get();
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title === "ReviewSubmitSuccess") {
        this.documentReviewedSuccess = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title === "ReviewDraftSuccess") {
        this.documentSavedAsDraft = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title === "NoAccess") {
        this.setState({
          noAccess: userMessageSettings[i].Message,
        });
        this.noAccess = userMessageSettings[i].Message;

      }
      else if (userMessageSettings[i].Title === "InvalidQueryParams") {
        this.setState({
          invalidQueryParam: userMessageSettings[i].Message,
        });
      }
    }
  }
  //Current User
  private async _currentUser() {
    this._Service.getCurrentUser()
      //sp.web.currentUser.get()
      .then(currentUser => {
        this.setState({
          currentUser: currentUser.Id,
          currentUserEmail: currentUser.Email,
        });
        console.log(this.state.currentUser);
      });

  }
  private _queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let id = params.get('hid');
    let detailid = params.get('dtlid');
    this.workFlow = params.get('wf');
    //console.log(id);
    //console.log(this.detailID);
    // if (this.props.project) {
    //   if (id !== "" && id !== null && detailid !== "" && detailid !== null && this.workFlow === "dcc" && this.workFlow !== null) {
    //     this.headerId = parseInt(id);
    //     this.valid = "ok";
    //     this.detailID = parseInt(detailid);
    //     this.setState({
    //       divForDCC: "",
    //       divForReview: "none",
    //       ifDccComment: "none",
    //     });
    //   }
    //   else if (id !== "" && id !== null && detailid !== "" && detailid !== null) {
    //     this.headerId = parseInt(id);
    //     this.valid = "ok";
    //     this.detailID = parseInt(detailid);
    //     this.setState({
    //       divForDCC: "none",
    //       ifDccComment: "none",
    //       divForReview: "",
    //     });
    //   }
    //   else if (id === "" || id === null || detailid === "" || detailid === null || this.workFlow !== "dcc" || this.workFlow === null) {
    //     this.setState({ accessDeniedMsgBar: "", loaderDisplay: "none", invalidMessage: this.state.invalidQueryParam });
    //     setTimeout(() => {
    //       this.setState({ accessDeniedMsgBar: 'none', });
    //       window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);

    //     }, 10000);
    //   }


    // }
    //else 
    {

      // if (id !== "" && id !== null && detailid !== "" && detailid !== null && this.workFlow !== "dcc") {
        if (id !== "" && id !== null && detailid !== "" && detailid !== null ) {
      
        this.headerId = parseInt(id);
        this.valid = "ok";
        this.detailID = parseInt(detailid);
        this.setState({
          divForReview: "",
          ifDccComment: "none",
        });
        this._onPageLoadDataBind();
        this._checkingReviewStatus();
      }
      else {
        this.setState({ accessDeniedMsgBar: "", loaderDisplay: "none", invalidMessage: this.state.invalidQueryParam });
        setTimeout(() => {
          this.setState({ accessDeniedMsgBar: 'none', });
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        }, 10000);
      }
    }

  }
  //Get Access Groups
  private async _accessGroups() {
    let AccessGroup: any[] = [];
    let ok = "No";
    // if (this.props.project) 
    // {
    //   //  AccessGroup = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.accessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'Project_SendReviewWF'").get();
     this._LAUrlGettingForPermission();
    //   this.setState({
    //     // access: "",
    //     accessDeniedMsgBar: "none",
    //     loaderDisplay: "none",
    //   });
    //   // this._queryParamGetting();
    //   this._checkingReviewStatus();
    //   this._onPageLoadDataBind();
    // }
    // else 
    {
      AccessGroup = await this._Service.getItemSelectFilter(this.props.siteUrl, this.props.accessGroups, "AccessGroups,AccessFields", "Title eq 'QDMS_SendReviewWF'")
      //AccessGroup = await this._Service.getQDMS_SendReviewWF(this.props.siteUrl, this.props.accessGroups)
      //AccessGroup = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendReviewWF'").get();
      let AccessGroupItems: any[] = AccessGroup[0].AccessGroups.split(',');
      console.log("AccessGroupItems", AccessGroupItems);
      const DocumentIndexItem: any = await this._Service.getByIdSelect(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, "DepartmentID")
      //const DocumentIndexItem: any = await this._Service.getBusinessDepartment(this.props.siteUrl, this.props.documentIndex, this.documentIndexId)
      //const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).select("DepartmentID,BusinessUnitID").get();
      console.log("DocumentIndexItem", DocumentIndexItem);
      //cheching if department selected
      if (DocumentIndexItem.DepartmentID !== null) {
        this.departmentExist === "Exists";
        let deptid = parseInt(DocumentIndexItem.DepartmentID);
        const departmentItem: any = await this._Service.getItemById(this.props.siteUrl, this.props.departmentList, deptid)
        //const departmentItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.departmentList).items.getById(deptid).get();
        //let AG = DepartmentItem[0].AccessGroups;
        console.log("departmentItem", departmentItem);
        let accessGroupvar = departmentItem.AccessGroups;
        const accessGroupItem: any = await this._Service.getItems(this.props.siteUrl, this.props.accessGroupDetailsList)
        //const accessGroupItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroupDetailsList).items.get();
        let accessGroupID;
        console.log(accessGroupItem.length);
        for (let a = 0; a < accessGroupItem.length; a++) {
          if (accessGroupItem[a].Title === accessGroupvar) {
            accessGroupID = accessGroupItem[a].GroupID;
            this.GetGroupMembers(this.props.context, accessGroupID);
          }
        }
      }
      //if no department  
      else {
        //alert("with bussinessUnit");
        // if (DocumentIndexItem.BusinessUnitID !== null) {
        //   this.departmentExist === "Exists";
        //   let bussinessUnitID = parseInt(DocumentIndexItem.BusinessUnitID);
        //   const bussinessUnitItem: any = await this._Service.getItemById(this.props.siteUrl, this.props.bussinessUnitList, bussinessUnitID)
        //   //const bussinessUnitItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.bussinessUnitList).items.getById(bussinessUnitID).get();
        //   console.log("departmentItem", bussinessUnitItem);
        //   let accessGroupvar = bussinessUnitItem.AccessGroups;
        //   // alert(accessGroupvar);
        //   const accessGroupItem: any = await this._Service.getItems(this.props.siteUrl, this.props.accessGroupDetailsList)
        //   //const accessGroupItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroupDetailsList).items.get();
        //   let accessGroupID;
        //   console.log(accessGroupItem.length);
        //   for (let a = 0; a < accessGroupItem.length; a++) {
        //     if (accessGroupItem[a].Title === accessGroupvar) {
        //       accessGroupID = accessGroupItem[a].GroupID;
        //       this.GetGroupMembers(this.props.context, accessGroupID);
        //     }
        //   }
        // }
      }
    }


  }


  private async _gettingGroupID(AccessGroupItems) {
    let AG;
    for (let a = 0; a < AccessGroupItems.length; a++) {
      AG = AccessGroupItems[a];
      const accessGroupID: any = await this._Service.getItemFilter(this.props.siteUrl, this.props.accessGroupDetailsList, "Title eq '" + AG + "'")
      //const accessGroupID: any = await this._Service.getItemTitleFilter(this.props.siteUrl, this.props.accessGroupDetailsList, AG)
      //const accessGroupID: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroupDetailsList).items.filter("Title eq '" + AG + "'").get();
      let AccessGroupID;
      if (accessGroupID.length > 0) {
        console.log(accessGroupID);
        AccessGroupID = accessGroupID[0].GroupID;
        console.log("AccessGroupID", AccessGroupID);
        this.GetGroupMembers(this.props.context, AccessGroupID);
      }
    }
  }
  private _LAUrlGettingForPermission = async () => {
    const laUrl = await this._Service.getItemFilter(this.props.siteUrl, this.props.requestListName, "Title eq 'QDMS_PermissionWebpart'")
    //const laUrl = await this._Service.getItemTitleFilter(this.props.siteUrl, this.props.requestListName, "QDMS_PermissionWebpart")
    //const laUrl = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestListName).items.filter("Title eq 'QDMS_PermissionWebpart'").get();
    console.log("PosturlForPermission", laUrl[0].PostUrl);
    this.postUrlForPermission = laUrl[0].PostUrl;
    this.triggerProjectPermissionFlow(laUrl[0].PostUrl);
    // this._queryParamGetting();
  }
  protected async triggerProjectPermissionFlow(PostUrl) {
    //alert("triggerProjectPermissionFlow")
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
    const postURL = PostUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': 'Project_SendReviewWF',
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.state.currentUserEmail
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      console.log(responseJSON['Status']);
      if (responseJSON['Status'] === "Valid") {
        this.setState({
          // access: "",
          accessDeniedMsgBar: "none",
          loaderDisplay: "none",
        });
        // this._queryParamGetting();
        this._checkingReviewStatus();
        this._onPageLoadDataBind();
      }
      else {
        this.setState({
          accessDeniedMsgBar: "",
          loaderDisplay: "none",
          statusMessage: { isShowMessage: true, message: this.noAccess, messageType: 1 },
        });
        setTimeout(() => {
          window.location.replace(window.location.protocol + "//" + window.location.hostname + "/" + this.props.siteUrl);
          // this.RedirectUrl;
        }, 20000);
      }

    }
    else { }

  }

  //checking current user  is a reviewer
  private _checkingReviewStatus() {
    this._Service.getItemSelectExpandFilter(
      this.props.siteUrl,
      this.props.workFlowDetail,
      "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,TaskID",
      "Responsible",
      "ID eq '" + this.detailID + "'"
    )
      //this._Service.getWFDetailWithResponsible(this.props.siteUrl, this.props.workFlowDetail, this.detailID)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,TaskID").expand("Responsible").filter("ID eq '" + this.detailID + "'").get()
      .then(Items => {
        // console.log(Items);
        this.taskID = Items[0].TaskID;
        if (this.state.currentUserEmail === Items[0].Responsible.EMail) {
          this.setState({ access: "", accessDeniedMsgBar: "none", comments: Items[0].ResponsibleComment, });
          if (Items[0].ResponseStatus === "Reviewed" || Items[0].ResponseStatus === "Returned with comments") {
            this.setState({ buttonHidden: "none", statusKey: Items[0].ResponseStatus });
          }
        }
        else {

          this.setState({ access: "none", loaderDisplay: "none", accessDeniedMsgBar: "", invalidMessage: this.noAccess });
          setTimeout(() => {
            this.setState({ accessDeniedMsgBar: 'none', });
            window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
          }, 10000);
        }
      });
    //for binding current reviewers comments in table
    this._Service.getItemSelectExpandFilter(
      this.props.siteUrl,
      this.props.workFlowDetail,
      "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate",
      "Responsible",
      "HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review') "
    )
      //this._Service.getDetailWorkflowReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate")
      //.expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review') ").get()
      .then(currentReviewersItems => {
        console.log("currentReviewersItems", currentReviewersItems);
        if (currentReviewersItems.length > 0) {
          console.log("currentReviewersItems", currentReviewersItems);
          this.setState({
            currentReviewComment: "",
            currentReviewItems: currentReviewersItems,
          });

        }
      });
  }
  //Headerlist items
  private _onPageLoadDataBind = () => {
    // if (this.props.project) {
    //   //header list for project
    //   //var headerItems = "Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,Owner/Title,Owner/ID,Owner/EMail,Revision,WorkflowStatus,Title,DocumentIndexID,RequesterComment,RequestedDate,DueDate,PreviousReviewHeader,DocumentID,SourceDocumentID,DocumentController/Title,DocumentController/EMail,DocumentController/Id,DCCCompletionDate,Workflow";

    //   //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId)
    //   //.select(headerItems).expand("Owner,Approver,Requester,DocumentController").get()
    //   this._Service.getItemSelectExpand(
    //     this.props.siteUrl,
    //     this.props.workflowHeaderListName,
    //     "Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,Owner/Title,Owner/ID,Owner/EMail,Revision,WorkflowStatus,Title,DocumentIndexID,RequesterComment,RequestedDate,DueDate,PreviousReviewHeader,DocumentID,SourceDocumentID,DocumentController/Title,DocumentController/EMail,DocumentController/Id,DCCCompletionDate,Workflow",
    //     "Owner,Approver,Requester,DocumentController"
    //   )
    //     //this._Service.getHeaderItemsDocumentController(this.props.siteUrl, this.props.workflowHeaderListName)
    //     .then(dataitems => {
    //       //console.log(workFlowHeaderItems);
    //       let workFlowHeaderItems: any = dataitems;
    //       let previousheadervalue = workFlowHeaderItems.PreviousReviewHeader;
    //       console.log(workFlowHeaderItems.PreviousReviewHeader);
    //       this.documentIndexId = workFlowHeaderItems.DocumentIndexID;
    //       this.sourceDocumentID = workFlowHeaderItems.SourceDocumentID;
    //       this.dueDateWithoutConversion = workFlowHeaderItems.DueDate;
    //       this.setState({
    //         requestorComment: workFlowHeaderItems.RequesterComment,
    //         requestorDate: moment(workFlowHeaderItems.RequestedDate).format('DD/MM/YYYY'),
    //         dueDate: moment(workFlowHeaderItems.DueDate).format('DD/MM/YYYY'),
    //         DueDate: workFlowHeaderItems.DueDate,
    //         workflowStatus: workFlowHeaderItems.WorkflowStatus,
    //         owner: workFlowHeaderItems.Owner.Title,
    //         ownerEmail: workFlowHeaderItems.Owner.EMail,
    //         ownerID: workFlowHeaderItems.Owner.ID,
    //         revision: workFlowHeaderItems.Revision,
    //         requestor: workFlowHeaderItems.Requester.Title,
    //         requestorEmail: workFlowHeaderItems.Requester.EMail,
    //         approverEmail: workFlowHeaderItems.Approver.EMail,
    //         approverName: workFlowHeaderItems.Approver.Title,
    //         approverId: workFlowHeaderItems.Approver.ID,
    //         documentID: workFlowHeaderItems.DocumentID,
    //         headerListItem: workFlowHeaderItems,
    //         hideproject: false,
    //         documentControllerEmail: workFlowHeaderItems.DocumentController.EMail,
    //         documentControllerName: workFlowHeaderItems.DocumentController.Title,
    //       });
    //       //if ((workFlowHeaderItems.PreviousReviewHeader !== "0" || workFlowHeaderItems.PreviousReviewHeader !== previousheadervalue)) {
    //       if ((workFlowHeaderItems.PreviousReviewHeader !== "0" && workFlowHeaderItems.Workflow === "Review")) {
    //         this.setState({ hideReviewersTable: "", });
    //         this._loadPreviousReturnWithComments(workFlowHeaderItems.PreviousReviewHeader);
    //       }
    //       this._documentIndexListBind(this.documentIndexId);
    //       if (workFlowHeaderItems.DocumentController === null) { this.setState({ ifDccComment: "none", }); }
    //       else {
    //         this._loadPreviousReturnWithComments(workFlowHeaderItems.PreviousReviewHeader);
    //         this._documentIndexListBind(this.documentIndexId);
    //         this.setState({
    //           ifDccComment: " ",
    //           dcc: workFlowHeaderItems.DocumentController.Title,
    //           dccCompletionDate: workFlowHeaderItems.DCCCompletionDate,
    //         });
    //       }
    //     });
    //   this._userMessageSettings();
    //   this._projectInformation();
    //   this._revisionLogChecking();
    // }
    // else
    {
      //header list for qdms 
      //var headerItems = "Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,Owner/Title,Owner/ID,Owner/EMail,Revision,WorkflowStatus,Title,DocumentIndexID,RequesterComment,RequestedDate,DueDate,PreviousReviewHeader,DocumentID,SourceDocumentID,Workflow";
      this._Service.getByIdSelectExpand(
        this.props.siteUrl,
        this.props.workflowHeaderListName,
        this.headerId,
        "Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,Owner/Title,Owner/ID,Owner/EMail,Revision,WorkflowStatus,Title,DocumentIndexID,RequesterComment,RequestedDate,DueDate,PreviousReviewHeader,DocumentID,SourceDocumentID,Workflow",
        "Owner,Approver,Requester"
      )
        //this._Service.getWFHOwnerApproverRequester(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId)
        //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).select(headerItems).expand("Owner,Approver,Requester").get()
        .then(dataitems => {
          // console.log(workFlowHeaderItems);
          // console.log(workFlowHeaderItems.PreviousReviewHeader);
          let workFlowHeaderItems: any = dataitems;
          this.documentIndexId = workFlowHeaderItems.DocumentIndexID;
          this.sourceDocumentID = workFlowHeaderItems.SourceDocumentID;
          this.dueDateWithoutConversion = workFlowHeaderItems.DueDate;
          this.setState({
            requestorComment: workFlowHeaderItems.RequesterComment,
            requestorDate: moment(workFlowHeaderItems.RequestedDate).format('DD/MM/YYYY'),
            dueDate: moment(workFlowHeaderItems.DueDate).format('DD/MM/YYYY'),
            DueDate: workFlowHeaderItems.DueDate,
            workflowStatus: workFlowHeaderItems.WorkflowStatus,
            owner: workFlowHeaderItems.Owner.Title,
            ownerEmail: workFlowHeaderItems.Owner.EMail,
            ownerID: workFlowHeaderItems.Owner.ID,
            revision: workFlowHeaderItems.Revision,
            requestor: workFlowHeaderItems.Requester.Title,
            requestorEmail: workFlowHeaderItems.Requester.EMail,
            approverEmail: workFlowHeaderItems.Approver.EMail,
            approverName: workFlowHeaderItems.Approver.Title,
            approverId: workFlowHeaderItems.Approver.ID,
            documentID: workFlowHeaderItems.DocumentID,
            headerListItem: workFlowHeaderItems,
          });

          if (workFlowHeaderItems.PreviousReviewHeader !== "0") {
            {
              this.setState({ hideReviewersTable: "", });
              this._loadPreviousReturnWithComments(workFlowHeaderItems.PreviousReviewHeader);
            }
          }
          this.setState({ divForDCC: "none", ifDccComment: "none", });
          this._documentIndexListBind(this.documentIndexId);
        });
      this._userMessageSettings();
    }
  }
  public async GetGroupMembers(context: WebPartContext, groupId: string) {
    let users: string[] = [];
    try {
      let response = await this._Service.getGroupMembers(groupId)
      /* let client: MSGraphClient = await context.msGraphClientFactory.getClient();
      let response = await client
        .api(`/groups/${groupId}/members`)
        .version('v1.0')
        .select(['mail', 'displayName'])
        .get(); */
      response.value.map((item: any) => {
        users.push(item);
      });
    } catch (error) {
      console.log('MSGraphService.GetGroupMembers Error: ', error);
    }
    console.log('MSGraphService.GetGroupMembers: ', users, "GroupID:", groupId);
    //cheching current users 
    if (users.length > 0) {
      this._checkingCurrent(users);
    }
    //return
  }
  private _checkingCurrent(userEmail) {

    for (var k in userEmail) {
      if (this.state.currentUserEmail === userEmail[k].mail) {
        this.valid = "Yes";
        this.setState({
          loaderDisplay: "none",
        });
        this._checkingReviewStatus();
        this._onPageLoadDataBind();
        break;
      }
    }
    if (this.valid !== "Yes") {

      this.setState({
        loaderDisplay: "none", access: "none", accessDeniedMsgBar: "", invalidMessage: this.noAccess,
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.props.redirectUrl);
      }, 10000);
    }
  }
  private _LAUrlGetting = async () => {
    const laUrl = await this._Service.getItemFilter(this.props.siteUrl, this.props.requestListName, "Title eq 'QDMS_DocumentPermission_UnderApproval'")
    //const laUrl = await this._Service.getQDMS_DocumentPermission_UnderApproval(this.props.siteUrl, this.props.requestListName)
    //const laUrl = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestListName).items.filter("Title eq 'QDMS_DocumentPermission_UnderApproval'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
  }
  private _LAUrlGettingForUnderReview = async () => {
    const laUrl = await this._Service.getItemFilter(this.props.siteUrl, this.props.requestListName, "Title eq 'QDMS_DocumentPermission_UnderReview'")
    //const laUrl = await this._Service.getQDMS_DocumentPermission_UnderReview(this.props.siteUrl, this.props.requestListName)
    //const laUrl = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestListName).items.filter("Title eq 'QDMS_DocumentPermission_UnderReview'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrlForUnderReview = laUrl[0].PostUrl;
  }

  

  //Adaptive Card
  /* private _LaUrlGettingAdaptive = async () => {

    const laUrl: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestListName).items.get();

    console.log("Posturl" + laUrl);

    for (let i = 0; i < laUrl.length; i++) {

      if (laUrl[i].Title === "Adaptive _Card") {

        this.postUrlForAdaptive = laUrl[i].PostUrl;

      }

    }

  } */

  public _adaptiveCard = async (Workflow, Email, Name, Type) => {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrlForAdaptive;
    var splitted = this.state.documentName.split(".");
    let ext = splitted[splitted.length - 1];
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'Workflow': Workflow,
      'DocumentIndexID': String(this.documentIndexId),
      'SourceDocumentID': String(this.sourceDocumentID),
      'HeaderID': String(this.headerId),
      'DetailID': String(this.newDetailItemID),
      'TaskID': String(this.taskID),
      'ext': ext,
      'Email': Email,
      'Name': Name,
      'Type': Type
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //getting document id
  private _documentIndexListBind(documentIndexId) {
    //document index list with document id
    this._Service.getByIdSelect(this.props.siteUrl, this.props.documentIndex, documentIndexId, "CriticalDocument,DocumentName,SourceDocument")
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(documentIndexId).select("CriticalDocument,DocumentName,SourceDocument").get()
      .then(documentIndexItems => {
        console.log(documentIndexItems);
        this.setState({
          documentIndexItems: documentIndexItems,
          documentName: documentIndexItems.DocumentName,
          criticalDocument: documentIndexItems.CriticalDocument,
          linkToDoc: documentIndexItems.SourceDocument.Url,
        });
      });
    // this.RevisionHistoryUrl = this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexId + "";
    // console.log(this.RevisionHistoryUrl);
  }

  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexId + "");
  }


  public _projectInformation = async () => {
    const projectInformation = await this._Service.getItems(this.props.siteUrl, this.props.projectInformationListName)
    //const projectInformation = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.projectInformationListName).items.get();
    console.log("projectInformation", projectInformation);
    if (projectInformation.length > 0) {
      for (var k in projectInformation) {
        if (projectInformation[k].Key === "ProjectName") {
          this.setState({
            projectName: projectInformation[k].Title,
          });
        }
        if (projectInformation[k].Key === "ProjectNumber") {
          this.setState({
            projectNumber: projectInformation[k].Title,
          });
        }
      }
    }
  }
  private _revisionLogChecking() {
    var today = new Date();
    let date = today.toLocaleString();
    //Updationg DocumentRevisionlog
    // if (this.props.project && this.workFlow === "dcc") {
    //   this._Service.getItemFilter(
    //     this.props.siteUrl,
    //     this.props.documentRevisionLog,
    //     "WorkflowID eq '" + this.headerId + "' and (DocumentIndexId eq '" + this.documentIndexId + "') and (Workflow eq 'DCC Review') and (Status eq 'Under Review')"
    //   )
    //     //this._Service.getDCCReviewUnderReview(this.props.siteUrl, this.props.documentRevisionLog, this.headerId, this.documentIndexId)
    //     //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.filter("WorkflowID eq '" + this.headerId + "' and (DocumentIndexId eq '" + this.documentIndexId + "') and (Workflow eq 'DCC Review') and (Status eq 'Under Review')").get()
    //     .then(ifyes => {
    //       if (ifyes.length > 0) {
    //         this.revisionLogID = ifyes[0].ID;
    //         console.log(ifyes[0].ID);
    //         this.setState({
    //           revisionLogID: ifyes[0].ID,
    //         });
    //       }
    //     });
    // }
    // else 
    {
      this._Service.getItemFilter(
        this.props.siteUrl,
        this.props.documentRevisionLog,
        "WorkflowID eq '" + this.headerId + "' and (DocumentIndexId eq '" + this.documentIndexId + "') and (Workflow eq 'Review') and (Status eq 'Under Review')"
      )
        //this._Service.getWorkflowReviewUnderReview(this.props.siteUrl, this.props.documentRevisionLog, this.headerId, this.documentIndexId)
        //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.filter("WorkflowID eq '" + this.headerId + "' and (DocumentIndexId eq '" + this.documentIndexId + "') and (Workflow eq 'Review') and (Status eq 'Under Review')").get()
        .then(ifyes => {
          if (ifyes.length > 0) {
            this.revisionLogID = ifyes[0].ID;
            console.log(ifyes[0].ID);
            this.setState({
              revisionLogID: ifyes[0].ID,
            });
          }
          else {
            // sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.add({
            //   Status: "Under Review",
            //   LogDate: this.currentDate,
            //   WorkflowID: this.headerId,
            //   DocumentIndexId: this.documentIndexId,
            //   DueDate: this.state.DueDate,
            //   Workflow: "Review",
            //   Revision: this.state.revision,
            //   Title: this.state.documentID,
            // }).then(reviID => {
            //   this.revisionLogID = reviID.data.ID;
            //   this.setState({
            //     revisionLogID: reviID.data.ID,
            //   });
            // });
          }
        });
    }
  }
  //Dropdown Status binding
  public _drpdwnStatus(option: { key: any; text: any }) {
    //console.log(option.key);
    this.setState({ statusKey: option.key, status: option.text });
  }
  //Comment Box
  private _commentBoxChange = (ev: React.FormEvent<HTMLInputElement>, Comment?: string) => {
    this.setState({ comments: Comment || '' });
  }
  //To view response comments in table.
  private async _loadPreviousReturnWithComments(previousReviewHeader) {
    const workflowDetailItems: any[] = await this._Service.getItemSelectExpandFilter(
      this.props.siteUrl,
      this.props.workFlowDetail,
      "Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate",
      "Responsible",
      "HeaderID eq '" + previousReviewHeader + "' and (Workflow eq 'Review')  "
    )
    //const workflowDetailItems: any[] = await this._Service.getResponsibleWithWFReview(this.props.siteUrl, this.props.workFlowDetail, previousReviewHeader)
    //const workflowDetailItems: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate").expand("Responsible").filter("HeaderID eq '" + previousReviewHeader + "' and (Workflow eq 'Review')  ").get();
    this.setState({
      reviewerItems: workflowDetailItems,
    });
    console.log(workflowDetailItems);
    const dccComments: any[] = await this._Service.getItemSelectExpandFilter(
      this.props.siteUrl,
      this.props.workFlowDetail,
      "Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate,ResponsibleComment,ResponseDate",
      "Responsible",
      "HeaderID eq '" + this.headerId + "' and (Workflow eq 'DCC Review')  "
    )
    //const dccComments: any[] = await this._Service.getWorkflowDCCReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
    //const dccComments: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate,ResponsibleComment,ResponseDate").expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'DCC Review')  ").get();
    if (dccComments.length > 0) {
      this.setState({
        dccReviewItems: dccComments,
        dccComment: dccComments[0].ResponsibleComment,
        dccCompletionDate: dccComments[0].ResponsibleComment
      });
      console.log("dccReviewItems", this.state.dccReviewItems);
    }
    // if (this.props.project && this.workFlow === "dcc") {
    //   const dccComments: any[] = await this._Service.getItemSelectExpandFilter(
    //     this.props.siteUrl,
    //     this.props.workFlowDetail,
    //     "Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate,ResponsibleComment,ResponseDate",
    //     "Responsible",
    //     "HeaderID eq '" + previousReviewHeader + "' and (Workflow eq 'DCC Review')  "
    //   )
    //   //const dccComments: any[] = await this._Service.getWorkflowDCCReview(this.props.siteUrl, this.props.workFlowDetail, previousReviewHeader)
    //   //const dccComments: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate,ResponsibleComment,ResponseDate").expand("Responsible").filter("HeaderID eq '" + previousReviewHeader + "' and (Workflow eq 'DCC Review')  ").get();
    //   if (dccComments.length > 0) {
    //     this.setState({
    //       dccReviewItems: dccComments,
    //       dccComment: dccComments[0].ResponsibleComment,
    //       dccCompletionDate: dccComments[0].ResponsibleComment
    //     });
    //     console.log("dccReviewItems", this.state.dccReviewItems);
    //   }
    // }
  }
  //Save as draft
  private _docReviewSaveAsDraft = () => {
    const detailitem = { ResponsibleComment: this.state.comments }
    this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, this.detailID, detailitem)
      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(this.detailID).update({
        ResponsibleComment: this.state.comments,
      }) */
      .then(r => {
        this.setState({
          comments: "",
          statusMessage: { isShowMessage: true, message: this.documentSavedAsDraft, messageType: 4 }
        });
      });
    this.validator.hideMessages();
    setTimeout(() => {
      this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
      window.location.replace(window.location.protocol + "//" + window.location.hostname + "/" + this.props.siteUrl);
      // this.RedirectUrl;
    }, 10000);

  }
  //submit
  private _docReviewSubmit = () => {
    this._revisionLogChecking();
    console.log(this.revisionLogID);
    let reviewStatus;
    let count = 0;
    var today = new Date();
    let date = today.toLocaleString();
    let cancelCount = 0;
    //checking validation
    if (this.state.statusKey !== "Returned with comments") {
      if (this.validator.fieldValid("status")) {
        const detailitem = {
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.currentDate,
        }
        this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, this.detailID, detailitem)
          /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(this.detailID).update({
            ResponsibleComment: this.state.comments,
            ResponseStatus: this.state.status,
            ResponseDate: this.currentDate,
          }) */
          .then(async deleteTask => {
            if (this.taskID !== null) {
              let list = await this._Service.deleteItemById(this.props.siteUrl, this.props.workflowTaskListName, this.taskID)
              //let list = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName);
              //await list.items.getById(this.taskID).delete();
            }
          }).then(detailLIstUpdate => {
            this._Service.getItemSelectFilter(
              this.props.siteUrl,
              this.props.workFlowDetail,
              "ResponseStatus",
              "HeaderID eq " + this.headerId + " and (Workflow eq 'Review')"
            )
              //this._Service.getWFDetailResponseStatus(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
              //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("ResponseStatus").filter("HeaderID eq " + this.headerId + " and (Workflow eq 'Review')").get()
              .then(async ResponseStatus => {
                if (ResponseStatus.length > 0) { //checking all reviewers response status
                  for (var k in ResponseStatus) {
                    if (ResponseStatus[k].ResponseStatus === "Reviewed") {
                      count++;
                    }
                    else if (ResponseStatus[k].ResponseStatus === "Returned with comments") {
                      reviewStatus = "Returned with comments";
                    }
                    else if (ResponseStatus[k].ResponseStatus === "Cancelled") {
                      cancelCount++;
                    }
                    else if (ResponseStatus[k].ResponseStatus === "Under Review") {
                      this.setState({
                        reviewPending: "Yes",
                      });
                    }
                  }
                  //all reviewers reviewed
                  if (ResponseStatus.length === count || (ResponseStatus.length === add(count, cancelCount))) {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headitem = {                   //headerlist
                      WorkflowStatus: "Under Approval",
                      Workflow: "Approval",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                      ({                   //headerlist
                        WorkflowStatus: "Under Approval",
                        Workflow: "Approval",
                        ReviewedDate: this.currentDate,
                      }); */
                    const inditem = {
                      WorkflowStatus: "Under Approval",//docIndex
                      Workflow: "Approval",
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
                      ({
                        WorkflowStatus: "Under Approval",//docIndex
                        Workflow: "Approval",
                      }); */
                    //Updationg DocumentRevisionlog 
                    const logitem = {
                      Status: "Reviewed",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.state.revisionLogID, logitem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.state.revisionLogID).update({
                      Status: "Reviewed",
                      LogDate: this.currentDate,
                    }); */
                    const logdata = {
                      Status: "Under Approval",
                      LogDate: this.currentDate,
                      WorkflowID: this.headerId,
                      DocumentIndexId: this.documentIndexId,
                      DueDate: this.state.DueDate,
                      Workflow: "Approval",
                      Revision: this.state.revision,
                      Title: this.state.documentID,
                    }
                    this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLog, logdata)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.add({
                      Status: "Under Approval",
                      LogDate: this.currentDate,
                      WorkflowID: this.headerId,
                      DocumentIndexId: this.documentIndexId,
                      DueDate: this.state.DueDate,
                      Workflow: "Approval",
                      Revision: this.state.revision,
                      Title: this.state.documentID,
                    }); */
                    //upadting source library without version change.            
                    let bodyArray = [
                      { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
                    ];
                    this._Service.validateUpdateListItem(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, bodyArray)
                    /* sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem
                      (
                        bodyArray,
                      ); */
                    //Task delegation getting user id from hubsite
                    this._Service.getUserIdByEmail(this.state.approverEmail)
                      //sp.web.siteUsers.getByEmail(this.state.approverEmail).get()
                      .then(async user => {
                        console.log('User Id: ', user.Id);
                        this.setState({
                          hubSiteUserId: user.Id,
                        });
                        //Task delegation 
                        // const taskDelegation: any[] = await this._Service.getItemSelectExpandFilter(
                        //   this.props.siteUrl,
                        //   this.props.taskDelegationSettingsListName,
                        //   "DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate",
                        //   "DelegatedFor,DelegatedTo",
                        //   "DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')"
                        // )
                        //const taskDelegation: any[] = await this._Service.getDelegateAndActive(this.props.siteUrl, this.props.taskDelegationSettingsListName, user.Id)
                        //const taskDelegation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationSettingsListName).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')").get();
                        // console.log(taskDelegation);
                        // if (taskDelegation.length > 0) {
                        //   let duedate = moment(this.dueDateWithoutConversion).toDate();
                        //   let toDate = moment(taskDelegation[0].ToDate).toDate();
                        //   let fromDate = moment(taskDelegation[0].FromDate).toDate();
                        //   duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                        //   toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                        //   fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                        //   if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                        //     this.setState({
                        //       approverEmail: taskDelegation[0].DelegatedTo.EMail,
                        //       approverName: taskDelegation[0].DelegatedTo.Title,
                        //       delegatedToId: taskDelegation[0].DelegatedTo.ID,
                        //       delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                        //     });
                        //     //duedate checking

                        //     //detail list adding an item for approval
                        //     this._Service.getUserIdByEmail(taskDelegation[0].DelegatedTo.EMail)
                        //       //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get()
                        //       .then(async DelegatedTo => {
                        //         this.setState({
                        //           delegateToIdInSubSite: DelegatedTo.Id,
                        //         });
                        //         this._Service.getUserIdByEmail(taskDelegation[0].DelegatedFor.EMail)
                        //           //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get()
                        //           .then(async DelegatedFor => {
                        //             this.setState({
                        //               delegateForIdInSubSite: DelegatedFor.Id,
                        //             });
                        //             const detailitem = {
                        //               HeaderIDId: Number(this.headerId),
                        //               Workflow: "Approval",
                        //               Title: this.state.documentName,
                        //               ResponsibleId: DelegatedTo.Id,
                        //               DueDate: this.state.DueDate,
                        //               DelegatedFromId: this.state.approverId,
                        //               ResponseStatus: "Under Approval",
                        //               SourceDocument: {
                        //                 // "__metadata": { type: "SP.FieldUrlValue" },
                        //                 Description: this.state.documentName,
                        //                 Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        //               },
                        //               OwnerId: this.state.ownerID,
                        //             }
                        //             this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detailitem)
                        //               /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
                        //                 ({
                        //                   HeaderIDId: Number(this.headerId),
                        //                   Workflow: "Approval",
                        //                   Title: this.state.documentName,
                        //                   ResponsibleId: DelegatedTo.Id,
                        //                   DueDate: this.state.DueDate,
                        //                   DelegatedFromId: this.state.approverId,
                        //                   ResponseStatus: "Under Approval",
                        //                   SourceDocument: {
                        //                     "__metadata": { type: "SP.FieldUrlValue" },
                        //                     Description: this.state.documentName,
                        //                     Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        //                   },
                        //                   OwnerId: this.state.ownerID,
                        //                 }) */
                        //               .then(async r => {
                        //                 this.setState({ detailIdForApprover: r.data.ID });
                        //                 this.newDetailItemID = r.data.ID;
                        //                 const detailitem = {
                        //                   Link: {
                        //                     // "__metadata": { type: "SP.FieldUrlValue" },
                        //                     Description: this.state.documentName + "-- Approve",
                        //                     Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //                   },
                        //                 }
                        //                 this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
                        //                 /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
                        //                   Link: {
                        //                     "__metadata": { type: "SP.FieldUrlValue" },
                        //                     Description: this.state.documentName + "-- Approve",
                        //                     Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //                   },
                        //                 }); */
                        //                 const headitem = {                   //headerlist
                        //                   ApproverId: DelegatedTo.Id,
                        //                 }
                        //                 this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                        //                 /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                        //                   ({                   //headerlist
                        //                     ApproverId: DelegatedTo.Id,
                        //                   }); */
                        //                 const inditem = {
                        //                   ApproverId: DelegatedTo.Id,
                        //                 }
                        //                 this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
                        //                 /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
                        //                   ({
                        //                     ApproverId: DelegatedTo.Id,
                        //                   }); */
                        //                 //upadting source library without version change.   
                        //                 const sourceitem = {
                        //                   ApproverId: DelegatedTo.Id,
                        //                 }
                        //                 await this._Service.updateItemById(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, sourceitem)
                        //                 /* await sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).update({
                        //                   ApproverId: DelegatedTo.Id,
                        //                 }); */
                        //                 //MY tasks list updation
                        //                 const taskitem = {
                        //                   Title: "Approve '" + this.state.documentName + "'",
                        //                   Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                        //                   DueDate: this.state.DueDate,
                        //                   StartDate: this.currentDate,
                        //                   AssignedToId: taskDelegation[0].DelegatedTo.ID,
                        //                   Workflow: "Approval",
                        //                   Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                        //                   DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                        //                   Source: "QDMS",
                        //                   DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
                        //                   Link: {
                        //                     // "__metadata": { type: "SP.FieldUrlValue" },
                        //                     Description: this.state.documentName + "-- Approve",
                        //                     Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //                   },

                        //                 }
                        //                 await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                        //                   /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                        //                     ({
                        //                       Title: "Approve '" + this.state.documentName + "'",
                        //                       Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                        //                       DueDate: this.state.DueDate,
                        //                       StartDate: this.currentDate,
                        //                       AssignedToId: taskDelegation[0].DelegatedTo.ID,
                        //                       Workflow: "Approval",
                        //                       Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                        //                       DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                        //                       Source: (this.props.project ? "Project" : "QDMS"),
                        //                       DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
                        //                       Link: {
                        //                         "__metadata": { type: "SP.FieldUrlValue" },
                        //                         Description: this.state.documentName + "-- Approve",
                        //                         Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //                       },
  
                        //                     }) */
                        //                   .then(taskId => {
                        //                     this.taskID = taskId.data.ID;
                        //                     const detaitem = {
                        //                       TaskID: taskId.data.ID,
                        //                     }
                        //                     this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detaitem)
                        //                       /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
                        //                         ({
                        //                           TaskID: taskId.data.ID,
                        //                         }) */
                        //                       .then(async aftermail => {
                        //                         this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");
                        //                         //this._adaptiveCard("Approval");
                        //                         // if (!this.props.project)
                        //                         {
                        //                           await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                        //                         }
                        //                         this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                        //                         //Email pending  emailbody to approver                 
                        //                         this.validator.hideMessages();
                        //                         this.setState({
                        //                           comments: "",
                        //                           statusKey: "",
                        //                           approverEmail: "",
                        //                           approverName: "",
                        //                           approverId: "",
                        //                           buttonHidden: "none"
                        //                         });

                        //                       }).then(redirect => {
                        //                         setTimeout(() => {
                        //                           this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                        //                           window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                        //                           //this.RedirectUrl;
                        //                         }, 10000);

                        //                       });//aftermai
                        //                     //notification preference checking  

                        //                   });//taskID
                        //               });//r

                        //           });//DelegatedFor
                        //       });//DelegatedTo
                        //   }
                        //   else {
                        //     const detdata = {
                        //       HeaderIDId: Number(this.headerId),
                        //       Workflow: "Approval",
                        //       Title: this.state.documentName,
                        //       ResponsibleId: this.state.approverId,
                        //       OwnerId: this.state.ownerID,
                        //       DueDate: this.state.DueDate,
                        //       ResponseStatus: "Under Approval",
                        //       SourceDocument: {
                        //         // "__metadata": { type: "SP.FieldUrlValue" },
                        //         Description: this.state.documentName,
                        //         Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        //       },
                        //     }
                        //     this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detdata)
                        //       /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
                        //         ({
                        //           HeaderIDId: Number(this.headerId),
                        //           Workflow: "Approval",
                        //           Title: this.state.documentName,
                        //           ResponsibleId: this.state.approverId,
                        //           OwnerId: this.state.ownerID,
                        //           DueDate: this.state.DueDate,
                        //           ResponseStatus: "Under Approval",
                        //           SourceDocument: {
                        //             "__metadata": { type: "SP.FieldUrlValue" },
                        //             Description: this.state.documentName,
                        //             Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        //           },
                        //         }) */
                        //       .then(async r => {
                        //         this.setState({ detailIdForApprover: r.data.ID });
                        //         this.newDetailItemID = r.data.ID;
                        //         const detitem = {
                        //           Link: {
                        //             // "__metadata": { type: "SP.FieldUrlValue" },
                        //             Description: this.state.documentName + "-- Approve",
                        //             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //           },
                        //         }
                        //         this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
                        //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
                        //           Link: {
                        //             "__metadata": { type: "SP.FieldUrlValue" },
                        //             Description: this.state.documentName + "-- Approve",
                        //             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //           },
                        //         }); */

                        //         //MY tasks list updation
                        //         const taskdata = {
                        //           Title: "Approve '" + this.state.documentName + "'",
                        //           Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                        //           DueDate: this.state.DueDate,
                        //           StartDate: this.currentDate,
                        //           AssignedToId: user.Id,
                        //           Workflow: "Approval",
                        //           Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                        //           Source: "QDMS",
                        //           Link: {
                        //             // "__metadata": { type: "SP.FieldUrlValue" },
                        //             Description: this.state.documentName + "-- Approve",
                        //             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //           },

                        //         }
                        //         await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
                        //           /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                        //             ({
                        //               Title: "Approve '" + this.state.documentName + "'",
                        //               Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                        //               DueDate: this.state.DueDate,
                        //               StartDate: this.currentDate,
                        //               AssignedToId: user.Id,
                        //               Workflow: "Approval",
                        //               Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                        //               Source: (this.props.project ? "Project" : "QDMS"),
                        //               Link: {
                        //                 "__metadata": { type: "SP.FieldUrlValue" },
                        //                 Description: this.state.documentName + "-- Approve",
                        //                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        //               },
  
                        //             }) */
                        //           .then(async taskId => {
                        //             const wfdetailitem = {
                        //               TaskID: taskId.data.ID,
                        //             }
                        //             this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, wfdetailitem)
                        //               /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
                        //                 ({
                        //                   TaskID: taskId.data.ID,
                        //                 }) */
                        //               .then(aftermail => {
                        //                 this.validator.hideMessages();
                        //                 this.setState({
                        //                   statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                        //                   comments: "",
                        //                   statusKey: "",
                        //                   approverEmail: "",
                        //                   approverName: "",
                        //                   approverId: "",
                        //                   buttonHidden: "none"
                        //                 });
                        //                 //Email pending  emailbody to approver  
                        //                 this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");

                        //               })
                        //               .then(redirect => {
                        //                 setTimeout(() => {
                        //                   this.setState({});
                        //                   window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                        //                   //this.RedirectUrl;
                        //                 }, 10000);

                        //               });//aftermai
                        //             //notification preference checking  
                        //             this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                        //             //if (!this.props.project) {
                        //             await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                        //             //}
                        //           });//taskID
                        //       });//r
                        //   }//else no delegation

                        // }

                        // else {
                          const detdata = {
                            HeaderIDId: Number(this.headerId),
                            Workflow: "Approval",
                            Title: this.state.documentName,
                            ResponsibleId: this.state.approverId,
                            DueDate: this.state.DueDate,
                            OwnerId: Number(this.state.ownerID),
                            ResponseStatus: "Under Approval",
                            SourceDocument: {
                              // "__metadata": { type: "SP.FieldUrlValue" },
                              Description: this.state.documentName,
                              Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                            },
                          }
                          this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detdata)
                            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
                              ({
                                HeaderIDId: Number(this.headerId),
                                Workflow: "Approval",
                                Title: this.state.documentName,
                                ResponsibleId: this.state.approverId,
                                DueDate: this.state.DueDate,
                                OwnerId: Number(this.state.ownerID),
                                ResponseStatus: "Under Approval",
                                SourceDocument: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName,
                                  Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                },
                              }) */
                            .then(async r => {
                              this.setState({ detailIdForApprover: r.data.ID });
                              this.newDetailItemID = r.data.ID;
                              const detaildata = {
                                Link: {
                                  // "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },
                              }
                              this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detaildata)
                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },
                              }); */

                              //MY tasks list updation
                              const taskitem = {
                                Title: "Approve '" + this.state.documentName + "'",
                                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                DueDate: this.state.DueDate,
                                StartDate: this.currentDate,
                                AssignedToId: user.Id,
                                Workflow: "Approval",
                                Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                Source: "QDMS",
                                Link: {
                                  // "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },

                              }
                              await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                              
                                /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                                  ({
                                    Title: "Approve '" + this.state.documentName + "'",
                                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                    DueDate: this.state.DueDate,
                                    StartDate: this.currentDate,
                                    AssignedToId: user.Id,
                                    Workflow: "Approval",
                                    Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                    Source: (this.props.project ? "Project" : "QDMS"),
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: this.state.documentName + "-- Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    },
  
                                  }) */
                                .then(async taskId => {
                                  this.taskID = taskId.data.ID;
                                  const detdata = {
                                    TaskID: taskId.data.ID,
                                  }
                                  
                                  this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detdata)

                                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
                                      ({
                                        TaskID: taskId.data.ID,
                                      }) */
                                    .then(async aftermail => {
                                      this.validator.hideMessages();
                                      
                                      this.setState({
                                        statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                        comments: "",
                                        statusKey: "",
                                        approverEmail: "",
                                        approverName: "",
                                        approverId: "",
                                        buttonHidden: "none"
                                      });
                                      //Email pending  emailbody to approver  
                                      this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");


                                    }).then(redirect => {
                                      setTimeout(() => {
                                        this.setState({});
                                        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                        //this.RedirectUrl;
                                      }, 10000);

                                    });//aftermai
                                  //notification preference checking  
                                  this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                  // if (!this.props.project) {
                                  // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                  // }

                                });//taskID
                            });//r
                        // }//else no delegation

                      }).catch(reject => console.error('Error getting Id of user by Email ', reject));
                  }
                  //any of the reviewer returned with comments
                  else if (reviewStatus === "Returned with comments" && this.state.reviewPending === "No") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headitem = {
                      WorkflowStatus: "Returned with comments",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                    /*  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update({
                       WorkflowStatus: "Returned with comments",
                       ReviewedDate: this.currentDate,
                     }); */
                    const inditem = {
                      WorkflowStatus: "Returned with comments",
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update({
                      WorkflowStatus: "Returned with comments",
                    }); */

                    //Updationg DocumentRevisionlog        
                    const logitem = {
                      Status: "Returned with comments",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.revisionLogID, inditem)
                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.revisionLogID).update({
                        Status: "Returned with comments",
                        LogDate: this.currentDate,
                      }) */
                      // //upadting source library without version change.            
                      // let bodyArray = [
                      //   { "FieldName": "WorkflowStatus", "FieldValue": "Returned with comments" }
                      // ];
                      // sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem(
                      //   bodyArray,
                      // )
                      .then(afterHeaderStatusUpdate => {
                        this.triggerDocumentReview(this.sourceDocumentID, "Returned with comments");
                        this._returnWithComments();
                        //mail to document controller if any one reviewer return with comments.
                        // if (this.props.project) {
                        //    this._sendAnEmailUsingMSGraph(this.state.documentControllerEmail, "DocReturn", this.state.documentControllerName, this.newDetailItemID); }
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none",
                        });
                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                        }, 10000);

                      });
                  }
                  //if any review process pending
                  else if (this.state.reviewPending === "Yes") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headitem = {
                      WorkflowStatus: "Under Review",
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update({
                        WorkflowStatus: "Under Review",
                      }) */
                      .then(async after => {
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none"
                        });

                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                          // this.RedirectUrl;
                        }, 10000);

                      });
                  }
                }
              });
          });
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    else {
      if (this.validator.fieldValid("status") && this.validator.fieldValid("comments")) {
        const detitem = {
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.currentDate,
        }
        this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, this.detailID, detitem)
          /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(this.detailID).update({
            ResponsibleComment: this.state.comments,
            ResponseStatus: this.state.status,
            ResponseDate: this.currentDate,
          }) */
          .then(async deleteTask => {
            if (this.taskID !== null) {
              let list = await this._Service.deleteItemById(this.props.siteUrl, this.props.workflowTaskListName, this.taskID)
              //let list = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName);
              //await list.items.getById(this.taskID).delete();
            }
          }).then(detailLIstUpdate => {
            this._Service.getItemSelectFilter(
              this.props.siteUrl,
              this.props.workFlowDetail,
              "ResponseStatus",
              "HeaderID eq " + this.headerId + " and (Workflow eq 'Review')"
            )
              //this._Service.getWFDetailResponseStatus(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
              //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("ResponseStatus").filter("HeaderID eq " + this.headerId + " and (Workflow eq 'Review')").get()
              .then(async ResponseStatus => {
                if (ResponseStatus.length > 0) { //checking all reviewers response status
                  for (var k in ResponseStatus) {
                    if (ResponseStatus[k].ResponseStatus === "Reviewed") {
                      count++;
                    }
                    else if (ResponseStatus[k].ResponseStatus === "Returned with comments") {
                      reviewStatus = "Returned with comments";
                    }
                    else if (ResponseStatus[k].ResponseStatus === "Cancelled") {
                      cancelCount++;
                    }
                    else if (ResponseStatus[k].ResponseStatus === "Under Review") {
                      this.setState({
                        reviewPending: "Yes",
                      });
                    }
                  }
                  //all reviewers reviewed
                  if (ResponseStatus.length === count || (ResponseStatus.length === add(count, cancelCount))) {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headitem = {                   //headerlist
                      WorkflowStatus: "Under Approval",
                      Workflow: "Approval",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                      ({                   //headerlist
                        WorkflowStatus: "Under Approval",
                        Workflow: "Approval",
                        ReviewedDate: this.currentDate,
                      }); */
                    const inditem = {
                      WorkflowStatus: "Under Approval",//docIndex
                      Workflow: "Approval",
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
                      ({
                        WorkflowStatus: "Under Approval",//docIndex
                        Workflow: "Approval",
                      }); */
                    //Updationg DocumentRevisionlog 
                    const logitem = {
                      Status: "Reviewed",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.state.revisionLogID, logitem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.state.revisionLogID).update({
                      Status: "Reviewed",
                      LogDate: this.currentDate,
                    }); */
                    const logdata = {
                      Status: "Under Approval",
                      LogDate: this.currentDate,
                      WorkflowID: this.headerId,
                      DocumentIndexId: this.documentIndexId,
                      DueDate: this.state.DueDate,
                      Workflow: "Approval",
                      Revision: this.state.revision,
                      Title: this.state.documentID,
                    }
                    this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLog, logdata)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.add({
                      Status: "Under Approval",
                      LogDate: this.currentDate,
                      WorkflowID: this.headerId,
                      DocumentIndexId: this.documentIndexId,
                      DueDate: this.state.DueDate,
                      Workflow: "Approval",
                      Revision: this.state.revision,
                      Title: this.state.documentID,
                    }); */
                    //upadting source library without version change.            
                    let bodyArray = [
                      { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
                    ];
                    this._Service.validateUpdateListItem(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, bodyArray)
                    /* sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem
                      (
                        bodyArray,
                      ); */
                    //Task delegation getting user id from hubsite
                    this._Service.getUserIdByEmail(this.state.approverEmail)
                      //sp.web.siteUsers.getByEmail(this.state.approverEmail).get()
                      .then(async user => {
                        console.log('User Id: ', user.Id);
                        this.setState({
                          hubSiteUserId: user.Id,
                        });
                        //Task delegation 
                        const taskDelegation: any[] = await this._Service.getItemSelectExpandFilter(
                          this.props.siteUrl,
                          this.props.taskDelegationSettingsListName,
                          "DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate",
                          "DelegatedFor,DelegatedTo",
                          "DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')"
                        )
                        //const taskDelegation: any[] = await this._Service.getDelegateAndActive(this.props.siteUrl, this.props.taskDelegationSettingsListName, user.Id)
                        //const taskDelegation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationSettingsListName).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')").get();
                        console.log(taskDelegation);
                        if (taskDelegation.length > 0) {
                          let duedate = moment(this.dueDateWithoutConversion).toDate();
                          let toDate = moment(taskDelegation[0].ToDate).toDate();
                          let fromDate = moment(taskDelegation[0].FromDate).toDate();
                          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                          toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                          fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                          if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                            this.setState({
                              approverEmail: taskDelegation[0].DelegatedTo.EMail,
                              approverName: taskDelegation[0].DelegatedTo.Title,
                              delegatedToId: taskDelegation[0].DelegatedTo.ID,
                              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                            });
                            //duedate checking

                            //detail list adding an item for approval
                            this._Service.getUserIdByEmail(taskDelegation[0].DelegatedTo.EMail)
                              //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get()
                              .then(async DelegatedTo => {
                                this.setState({
                                  delegateToIdInSubSite: DelegatedTo.Id,
                                });
                                this._Service.getUserIdByEmail(taskDelegation[0].DelegatedFor.EMail)
                                  //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get()
                                  .then(async DelegatedFor => {
                                    this.setState({
                                      delegateForIdInSubSite: DelegatedFor.Id,
                                    });
                                    const detitem = {
                                      HeaderIDId: Number(this.headerId),
                                      Workflow: "Approval",
                                      Title: this.state.documentName,
                                      ResponsibleId: DelegatedTo.Id,
                                      DueDate: this.state.DueDate,
                                      DelegatedFromId: this.state.approverId,
                                      ResponseStatus: "Under Approval",
                                      SourceDocument: {
                                        // "__metadata": { type: "SP.FieldUrlValue" },
                                        Description: this.state.documentName,
                                        Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                      },
                                      OwnerId: this.state.ownerID,
                                    }
                                    this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detitem)
                                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
                                        ({
                                          HeaderIDId: Number(this.headerId),
                                          Workflow: "Approval",
                                          Title: this.state.documentName,
                                          ResponsibleId: DelegatedTo.Id,
                                          DueDate: this.state.DueDate,
                                          DelegatedFromId: this.state.approverId,
                                          ResponseStatus: "Under Approval",
                                          SourceDocument: {
                                            "__metadata": { type: "SP.FieldUrlValue" },
                                            Description: this.state.documentName,
                                            Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                          },
                                          OwnerId: this.state.ownerID,
                                        }) */
                                      .then(async r => {
                                        this.setState({ detailIdForApprover: r.data.ID });
                                        this.newDetailItemID = r.data.ID;
                                        const detaitem = {
                                          Link: {
                                            // "__metadata": { type: "SP.FieldUrlValue" },
                                            Description: this.state.documentName + "-- Approve",
                                            Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                          },
                                        }
                                        this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detaitem)
                                        /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
                                          Link: {
                                            "__metadata": { type: "SP.FieldUrlValue" },
                                            Description: this.state.documentName + "-- Approve",
                                            Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                          },
                                        }); */
                                        const headitem = {                   //headerlist
                                          ApproverId: DelegatedTo.Id,
                                        }
                                        this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                                        /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                                          ({                   //headerlist
                                            ApproverId: DelegatedTo.Id,
                                          }); */
                                        const inditem = {
                                          ApproverId: DelegatedTo.Id,
                                        }
                                        this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
                                        /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
                                          ({
                                            ApproverId: DelegatedTo.Id,
                                          }); */
                                        //upadting source library without version change. 
                                        const sourceitem = {
                                          ApproverId: DelegatedTo.Id,
                                        }
                                        await this._Service.updateItemById(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, sourceitem)
                                        /* await sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).update({
                                          ApproverId: DelegatedTo.Id,

                                        }); */
                                        //MY tasks list updation
                                        const taskitem = {
                                          Title: "Approve '" + this.state.documentName + "'",
                                          Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                          DueDate: this.state.DueDate,
                                          StartDate: this.currentDate,
                                          AssignedToId: taskDelegation[0].DelegatedTo.ID,
                                          Workflow: "Approval",
                                          Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                          DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                                          Source: "QDMS",
                                          DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
                                          Link: {
                                            // "__metadata": { type: "SP.FieldUrlValue" },
                                            Description: this.state.documentName + "-- Approve",
                                            Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                          },

                                        }
                                        await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                                          /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                                            ({
                                              Title: "Approve '" + this.state.documentName + "'",
                                              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                              DueDate: this.state.DueDate,
                                              StartDate: this.currentDate,
                                              AssignedToId: taskDelegation[0].DelegatedTo.ID,
                                              Workflow: "Approval",
                                              Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                              DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                                              Source: (this.props.project ? "Project" : "QDMS"),
                                              DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
                                              Link: {
                                                "__metadata": { type: "SP.FieldUrlValue" },
                                                Description: this.state.documentName + "-- Approve",
                                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                              },
  
                                            }) */
                                          .then(taskId => {
                                            this.taskID = taskId.data.ID;
                                            const taskitem = {
                                              TaskID: taskId.data.ID,
                                            }
                                            this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, taskitem)
                                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
                                                ({
                                                  TaskID: taskId.data.ID,
                                                }) */
                                              .then(async aftermail => {
                                                this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");
                                                //this._adaptiveCard("Approval");
                                                //if (!this.props.project) {
                                                // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                                //}
                                                this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                                //Email pending  emailbody to approver                 
                                                this.validator.hideMessages();
                                                this.setState({
                                                  // statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                                  comments: "",
                                                  statusKey: "",
                                                  approverEmail: "",
                                                  approverName: "",
                                                  approverId: "",
                                                  buttonHidden: "none"
                                                });

                                              }).then(redirect => {
                                                setTimeout(() => {
                                                  this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                                                  window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                                  //this.RedirectUrl;
                                                }, 10000);

                                              });//aftermai
                                            //notification preference checking  

                                          });//taskID
                                      });//r

                                  });//DelegatedFor
                              });//DelegatedTo
                          }
                          else {
                            const detdata = {
                              HeaderIDId: Number(this.headerId),
                              Workflow: "Approval",
                              Title: this.state.documentName,
                              ResponsibleId: this.state.approverId,
                              OwnerId: this.state.ownerID,
                              DueDate: this.state.DueDate,
                              ResponseStatus: "Under Approval",
                              SourceDocument: {
                                // "__metadata": { type: "SP.FieldUrlValue" },
                                Description: this.state.documentName,
                                Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                              },
                            }
                            this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detdata)
                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
                                ({
                                  HeaderIDId: Number(this.headerId),
                                  Workflow: "Approval",
                                  Title: this.state.documentName,
                                  ResponsibleId: this.state.approverId,
                                  OwnerId: this.state.ownerID,
                                  DueDate: this.state.DueDate,
                                  ResponseStatus: "Under Approval",
                                  SourceDocument: {
                                    "__metadata": { type: "SP.FieldUrlValue" },
                                    Description: this.state.documentName,
                                    Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                  },
                                }) */
                              .then(async r => {
                                this.setState({ detailIdForApprover: r.data.ID });
                                this.newDetailItemID = r.data.ID;
                                const detitem = {
                                  Link: {
                                    // "__metadata": { type: "SP.FieldUrlValue" },
                                    Description: this.state.documentName + "-- Approve",
                                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                  },
                                }
                                this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
                                /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
                                  Link: {
                                    "__metadata": { type: "SP.FieldUrlValue" },
                                    Description: this.state.documentName + "-- Approve",
                                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                  },
                                }); */

                                //MY tasks list updation
                                const taskitem = {
                                  Title: "Approve '" + this.state.documentName + "'",
                                  Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                  DueDate: this.state.DueDate,
                                  StartDate: this.currentDate,
                                  AssignedToId: user.Id,
                                  Workflow: "Approval",
                                  Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                  Source: "QDMS",
                                  Link: {
                                    // "__metadata": { type: "SP.FieldUrlValue" },
                                    Description: this.state.documentName + "-- Approve",
                                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                  },

                                }
                                await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                                  /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                                    ({
                                      Title: "Approve '" + this.state.documentName + "'",
                                      Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                      DueDate: this.state.DueDate,
                                      StartDate: this.currentDate,
                                      AssignedToId: user.Id,
                                      Workflow: "Approval",
                                      Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                      Source: (this.props.project ? "Project" : "QDMS"),
                                      Link: {
                                        "__metadata": { type: "SP.FieldUrlValue" },
                                        Description: this.state.documentName + "-- Approve",
                                        Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                      },
  
                                    }) */
                                  .then(async taskId => {
                                    const detdata = {
                                      TaskID: taskId.data.ID,
                                    }
                                    this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, {
                                      TaskID: taskId.data.ID,
                                    })
                                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
                                        ({
                                          TaskID: taskId.data.ID,
                                        }) */
                                      .then(aftermail => {
                                        this.validator.hideMessages();
                                        this.setState({
                                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                          comments: "",
                                          statusKey: "",
                                          approverEmail: "",
                                          approverName: "",
                                          approverId: "",
                                          buttonHidden: "none"
                                        });
                                        //Email pending  emailbody to approver  
                                        this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");

                                      }).then(redirect => {
                                        setTimeout(() => {
                                          this.setState({});
                                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                          //this.RedirectUrl;
                                        }, 10000);

                                      });//aftermai
                                    //notification preference checking  
                                    this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                    //if (!this.props.project) {
                                    // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                    // }
                                  });//taskID
                              });//r
                          }//else no delegation

                        }

                        else {
                          const detitem = {
                            HeaderIDId: Number(this.headerId),
                            Workflow: "Approval",
                            Title: this.state.documentName,
                            ResponsibleId: this.state.approverId,
                            DueDate: this.state.DueDate,
                            OwnerId: Number(this.state.ownerID),
                            ResponseStatus: "Under Approval",
                            SourceDocument: {
                              // "__metadata": { type: "SP.FieldUrlValue" },
                              Description: this.state.documentName,
                              Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                            },
                          }
                          this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detitem)
                            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
                              ({
                                HeaderIDId: Number(this.headerId),
                                Workflow: "Approval",
                                Title: this.state.documentName,
                                ResponsibleId: this.state.approverId,
                                DueDate: this.state.DueDate,
                                OwnerId: Number(this.state.ownerID),
                                ResponseStatus: "Under Approval",
                                SourceDocument: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName,
                                  Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                },
                              }) */
                            .then(async r => {
                              this.setState({ detailIdForApprover: r.data.ID });
                              this.newDetailItemID = r.data.ID;
                              const detitem = {
                                Link: {
                                  // "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },
                              }

                              this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },
                              }); */

                              //MY tasks list updation
                              const taskitem = {
                                Title: "Approve '" + this.state.documentName + "'",
                                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                DueDate: this.state.DueDate,
                                StartDate: this.currentDate,
                                AssignedToId: user.Id,
                                Workflow: "Approval",
                                Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                Source: "QDMS",
                                Link: {
                                  // "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },

                              }
                              await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                                /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                                  ({
                                    Title: "Approve '" + this.state.documentName + "'",
                                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                    DueDate: this.state.DueDate,
                                    StartDate: this.currentDate,
                                    AssignedToId: user.Id,
                                    Workflow: "Approval",
                                    Priority: (this.state.criticalDocument === true ? "Critical" : ""),
                                    Source: (this.props.project ? "Project" : "QDMS"),
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: this.state.documentName + "-- Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    },
  
                                  }) */
                                .then(async taskId => {
                                  this.taskID = taskId.data.ID;
                                  const detailitem =
                                  {
                                    TaskID: taskId.data.ID,
                                  }
                                  this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
                                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
                                      ({
                                        TaskID: taskId.data.ID,
                                      }). */
                                    .then(async aftermail => {
                                      this.validator.hideMessages();
                                      this.setState({
                                        statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                        comments: "",
                                        statusKey: "",
                                        approverEmail: "",
                                        approverName: "",
                                        approverId: "",
                                        buttonHidden: "none"
                                      });
                                      //Email pending  emailbody to approver  
                                      this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");


                                    }).then(redirect => {
                                      setTimeout(() => {
                                        this.setState({});
                                        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                        //this.RedirectUrl;
                                      }, 10000);

                                    });//aftermai
                                  //notification preference checking  
                                  this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                  // if (!this.props.project) {
                                  // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                  //}

                                });//taskID
                            });//r
                        }//else no delegation

                      }).catch(reject => console.error('Error getting Id of user by Email ', reject));
                  }
                  //any of the reviewer returned with comments
                  else if (reviewStatus === "Returned with comments" && this.state.reviewPending === "No") {
                    this.setState({
                      buttonHidden: "none",
                      statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                    });
                    const headitem = {
                      WorkflowStatus: "Returned with comments",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update({
                      WorkflowStatus: "Returned with comments",
                      ReviewedDate: this.currentDate,
                    }); */
                    const inditem = {
                      WorkflowStatus: "Returned with comments",
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update({
                      WorkflowStatus: "Returned with comments",
                    }); */

                    //Updationg DocumentRevisionlog     
                    const logitem = {
                      Status: "Returned with comments",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.revisionLogID, logitem)
                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.revisionLogID).update({
                        Status: "Returned with comments",
                        LogDate: this.currentDate,
                      }) */
                      // //upadting source library without version change.            
                      // let bodyArray = [
                      //   { "FieldName": "WorkflowStatus", "FieldValue": "Returned with comments" }
                      // ];
                      // sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem(
                      //   bodyArray,
                      // )
                      .then(afterHeaderStatusUpdate => {
                        this.triggerDocumentReview(this.sourceDocumentID, "Returned with comments");
                        this._returnWithComments();
                        //mail to document controller if any one reviewer return with comments.
                        // if (this.props.project) { this._sendAnEmailUsingMSGraph(this.state.documentControllerEmail, "DocReturn", this.state.documentControllerName, this.newDetailItemID); }
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none",
                        });
                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                        }, 10000);

                      });
                  }
                  //if any review process pending
                  else if (this.state.reviewPending === "Yes") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headitem = {
                      WorkflowStatus: "Under Review",
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update({
                        WorkflowStatus: "Under Review",
                      }) */
                      .then(async after => {
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none"
                        });

                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                          // this.RedirectUrl;
                        }, 10000);

                      });
                  }
                }
              });
          });
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
  }
  private _returnWithComments() {
    this._sendAnEmailUsingMSGraph(this.state.requestorEmail, "DocReturn", this.state.requestor, this.newDetailItemID);
    this._sendAnEmailUsingMSGraph(this.state.ownerEmail, "DocReturn", this.state.owner, this.newDetailItemID);

  }
  // private _docDCCReviewSubmit = async () => {
  //   this._revisionLogChecking();
  //   this.setState({
  //     revisionLogID: this.revisionLogID
  //   });
  //   let critical;
  //   var today = new Date();
  //   let date = today.toLocaleString();
  //   if (this.validator.fieldValid("status") && this.validator.fieldValid("comments")) {
  //     this.validator.hideMessages();
  //     const detitem = {
  //       ResponsibleComment: this.state.comments,
  //       ResponseStatus: this.state.status,
  //       ResponseDate: this.currentDate,
  //     }
  //     this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, this.detailID, detitem)
  //       /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(this.detailID).update({
  //         ResponsibleComment: this.state.comments,
  //         ResponseStatus: this.state.status,
  //         ResponseDate: this.currentDate,
  //       }) */
  //       .then(async deleteTask => {
  //         if (this.taskID !== null) {
  //           let list = await this._Service.deleteItemById(this.props.siteUrl, this.props.workflowTaskListName, this.taskID)
  //           //let list = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName);
  //           //await list.items.getById(this.taskID).delete();
  //         }
  //       });
  //     //if dcc review return with comments
  //     if (this.state.status === "Returned with comments") {
  //       this.setState({
  //         buttonHidden: "none",
  //       });
  //       const headdata = {
  //         WorkflowStatus: "Returned with comments",
  //         DCCCompletionDate: this.currentDate,
  //         Workflow: "DCC Review",
  //       }
  //       this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headdata)
  //       /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update({
  //         WorkflowStatus: "Returned with comments",
  //         DCCCompletionDate: this.currentDate,
  //         Workflow: "DCC Review",
  //       }); */
  //       const inditem = {
  //         WorkflowStatus: "Returned with comments",
  //         Workflow: "DCC Review",
  //       }
  //       this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
  //       /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update({
  //         WorkflowStatus: "Returned with comments",
  //         Workflow: "DCC Review",
  //       }); */
  //       //upadting source library without version change.            
  //       // let bodyArray = [
  //       //   { "FieldName": "WorkflowStatus", "FieldValue": "Returned with comments" }, { "FieldName": "Workflow", "FieldValue": "DCC Review" }
  //       // ];
  //       // sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem(
  //       //   bodyArray,
  //       // ).then(afterHeaderStatusUpdate => {
  //       //Updationg DocumentRevisionlog 
  //       const logitem = {
  //         Status: "DCC Review - Returned with comments",
  //         LogDate: this.currentDate,
  //       }
  //       this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.revisionLogID, logitem)
  //       /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.revisionLogID).update({
  //         Status: "DCC Review - Returned with comments",
  //         LogDate: this.currentDate,
  //       }); */
  //       this.triggerDocumentReview(this.sourceDocumentID, "Returned with comments");
  //       this._returnWithComments();
  //       this.validator.hideMessages();
  //       this.setState({
  //         statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
  //         comments: "",
  //         statusKey: "",
  //         buttonHidden: "none",
  //       });
  //       setTimeout(() => {
  //         this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //         window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //         //this.RedirectUrl;
  //       }, 10000);

  //       //  });

  //     }
  //     //if there is reviewers and Reviewed
  //     else {
  //       const headerItemsForDCCSubmit = await this._Service.getByIdSelectExpand(
  //         this.props.siteUrl,
  //         this.props.workflowHeaderListName,
  //         this.headerId,
  //         "Reviewers/ID,Reviewers/Title,Reviewers/EMail",
  //         "Reviewers"
  //       )
  //       //const headerItemsForDCCSubmit = await this._Service.getReviewersData(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId)
  //       //const headerItemsForDCCSubmit = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.select("Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Reviewers").getById(this.headerId).get();
  //       console.log(headerItemsForDCCSubmit);
  //       this.setState({
  //         reviewers: headerItemsForDCCSubmit.ReviewersId,
  //       });
  //       console.log(this.state.reviewers);
  //       if (this.state.reviewers !== null) {
  //         this.setState({
  //           buttonHidden: "none",
  //         });
  //         //for reviewers if exist
  //         //Updationg DocumentRevisionlog 
  //         const logitem = {
  //           Status: "DCC - Reviewed",
  //           LogDate: this.currentDate,
  //         }
  //         this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.state.revisionLogID, logitem)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.state.revisionLogID).update({
  //           Status: "DCC - Reviewed",
  //           LogDate: this.currentDate,
  //         }); */
  //         const logdata = {
  //           Status: "Under Review",
  //           LogDate: this.currentDate,
  //           WorkflowID: this.headerId,
  //           DocumentIndexId: this.documentIndexId,
  //           DueDate: this.state.DueDate,
  //           Workflow: "Review",
  //           Revision: this.state.revision,
  //           Title: this.state.documentID,
  //         }
  //         this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLog, logdata)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.add({
  //           Status: "Under Review",
  //           LogDate: this.currentDate,
  //           WorkflowID: this.headerId,
  //           DocumentIndexId: this.documentIndexId,
  //           DueDate: this.state.DueDate,
  //           Workflow: "Review",
  //           Revision: this.state.revision,
  //           Title: this.state.documentID,
  //         }); */
  //         const headitem = {
  //           WorkflowStatus: "Under Review",
  //           Workflow: "Review",
  //           ReviewedDate: this.currentDate,
  //         }
  //         this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
  //           ({                   //headerlist
  //             WorkflowStatus: "Under Review",
  //             Workflow: "Review",
  //             ReviewedDate: this.currentDate,
  //           }); */
  //         const inditem = {
  //           WorkflowStatus: "Under Review",//docIndex
  //           Workflow: "Review",
  //         }
  //         this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
  //         /*  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
  //            ({
  //              WorkflowStatus: "Under Review",//docIndex
  //              Workflow: "Review",
  //            }); */

  //         //upadting source library without version change.           
  //         // sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).update({
  //         //   WorkflowStatus: "Under Review",
  //         //   Workflow: "Review",
  //         // });
  //         //for reviewers if exist
  //         for (var k = 0; k <= this.state.reviewers.length; k++) {
  //           console.log(this.state.reviewers[k]);
  //           let reviewID = this.state.reviewers[k];
  //           await this._Service.getUserById(parseInt(reviewID))
  //             //await sp.web.siteUsers.getById(parseInt(reviewID)).get()
  //             .then(async user => {
  //               console.log(user);
  //               await this._Service.getUserIdByEmail(user.Email)
  //                 //await sp.web.siteUsers.getByEmail(user.Email).get()
  //                 .then(async hubsieUser => {
  //                   console.log(hubsieUser.Id);
  //                   const taskDelegation: any[] = await this._Service.getItemSelectExpandFilter(
  //                     this.props.siteUrl,
  //                     this.props.taskDelegationSettingsListName,
  //                     "DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate",
  //                     "DelegatedFor,DelegatedTo",
  //                     "DelegatedFor/ID eq '" + hubsieUser.Id + "' and(Status eq 'Active')"
  //                   )
  //                   //const taskDelegation: any[] = await this._Service.getDelegateAndActive(this.props.siteUrl, this.props.taskDelegationSettingsListName, hubsieUser.Id)
  //                   //const taskDelegation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationSettingsListName).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + hubsieUser.Id + "' and(Status eq 'Active')").get();
  //                   console.log(taskDelegation);
  //                   //Check if Task Delegation
  //                   if (taskDelegation.length > 0) {
  //                     let duedate = moment(this.dueDateWithoutConversion).toDate();
  //                     let toDate = moment(taskDelegation[0].ToDate).toDate();
  //                     let fromDate = moment(taskDelegation[0].FromDate).toDate();
  //                     duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
  //                     toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
  //                     fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  //                     if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
  //                       this.setState({
  //                         approverEmail: taskDelegation[0].DelegatedTo.EMail,
  //                         approverName: taskDelegation[0].DelegatedTo.Title,
  //                         delegatedToId: taskDelegation[0].DelegatedTo.ID,
  //                         delegatedFromId: taskDelegation[0].DelegatedFor.ID,
  //                       });

  //                       //Get Delegated To ID
  //                       this._Service.getUserIdByEmail(taskDelegation[0].DelegatedTo.EMail)
  //                         //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get()
  //                         .then(async DelegatedTo => {

  //                           this.setState({
  //                             delegateToIdInSubSite: DelegatedTo.Id,
  //                           });
  //                           //Get Delegated For ID
  //                           this._Service.getUserIdByEmail(taskDelegation[0].DelegatedFor.EMail)
  //                             //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get()
  //                             .then(async DelegatedFor => {

  //                               this.setState({
  //                                 delegateForIdInSubSite: DelegatedFor.Id,
  //                               });
  //                               //detail list adding an item for reviewers
  //                               let index = this.state.reviewers.indexOf(DelegatedFor.Id);
  //                               this.state.reviewers[index] = DelegatedTo.Id;
  //                               const detailitem = {
  //                                 HeaderIDId: Number(this.headerId),
  //                                 Workflow: "Review",
  //                                 Title: this.state.documentName,
  //                                 ResponsibleId: DelegatedTo.Id,
  //                                 DueDate: this.state.DueDate,
  //                                 DelegatedFromId: DelegatedFor.Id,
  //                                 OwnerId: this.state.ownerID,
  //                                 ResponseStatus: "Under Review",
  //                                 SourceDocument: {
  //                                   "__metadata": { type: "SP.FieldUrlValue" },
  //                                   Description: this.state.documentName,
  //                                   Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                                 },
  //                               }
  //                               this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detailitem)
  //                                 /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add({
  //                                   HeaderIDId: Number(this.headerId),
  //                                   Workflow: "Review",
  //                                   Title: this.state.documentName,
  //                                   ResponsibleId: DelegatedTo.Id,
  //                                   DueDate: this.state.DueDate,
  //                                   DelegatedFromId: DelegatedFor.Id,
  //                                   OwnerId: this.state.ownerID,
  //                                   ResponseStatus: "Under Review",
  //                                   SourceDocument: {
  //                                     "__metadata": { type: "SP.FieldUrlValue" },
  //                                     Description: this.state.documentName,
  //                                     Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                                   },
  //                                 }) */
  //                                 .then(async r => {
  //                                   this.setState({ detailIdForApprover: r.data.ID });
  //                                   this.newDetailItemID = r.data.ID;
  //                                   const detailitem = {
  //                                     Link: {
  //                                       "__metadata": { type: "SP.FieldUrlValue" },
  //                                       Description: this.state.documentName + "-- Review",
  //                                       Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                                     },
  //                                   }
  //                                   await this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
  //                                   /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
  //                                     Link: {
  //                                       "__metadata": { type: "SP.FieldUrlValue" },
  //                                       Description: this.state.documentName + "-- Review",
  //                                       Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                                     },
  //                                   }); *///Update link
  //                                   const headitem = {                   //headerlist
  //                                     ReviewersId: { results: this.state.reviewers }
  //                                   }
  //                                   await this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
  //                                   /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
  //                                     ({                   //headerlist
  //                                       ReviewersId: { results: this.state.reviewers }
  //                                     }); */
  //                                   const inditem = {
  //                                     ReviewersId: { results: this.state.reviewers }
  //                                   }
  //                                   await this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
  //                                   /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
  //                                     ({
  //                                       ReviewersId: { results: this.state.reviewers }
  //                                     }); */
  //                                   const sourceitem = {
  //                                     ReviewersId: { results: this.state.reviewers }
  //                                   }
  //                                   await this._Service.updateItemById(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, sourceitem)
  //                                   /* sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).update({
  //                                     ReviewersId: { results: this.state.reviewers }
  //                                   }); */
  //                                   //MY tasks list updation with delegated from
  //                                   const taskdata = {
  //                                     Title: "Review '" + this.state.documentName + "'",
  //                                     Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.currentDate + "'",
  //                                     DueDate: this.state.DueDate,
  //                                     StartDate: this.currentDate,
  //                                     AssignedToId: taskDelegation[0].DelegatedTo.ID,
  //                                     Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                                     DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
  //                                     Source: (this.props.project ? "Project" : "QDMS"),
  //                                     DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
  //                                     Workflow: "Review",
  //                                     Link: {
  //                                       "__metadata": { type: "SP.FieldUrlValue" },
  //                                       Description: this.state.documentName + "-- Review",
  //                                       Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                                     },
  //                                   }
  //                                   await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
  //                                     /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add({
  //                                       Title: "Review '" + this.state.documentName + "'",
  //                                       Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.currentDate + "'",
  //                                       DueDate: this.state.DueDate,
  //                                       StartDate: this.currentDate,
  //                                       AssignedToId: taskDelegation[0].DelegatedTo.ID,
  //                                       Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                                       DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
  //                                       Source: (this.props.project ? "Project" : "QDMS"),
  //                                       DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
  //                                       Workflow: "Review",
  //                                       Link: {
  //                                         "__metadata": { type: "SP.FieldUrlValue" },
  //                                         Description: this.state.documentName + "-- Review",
  //                                         Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                                       },
  //                                     }) */
  //                                     .then(taskId => {
  //                                       const taskitem = {
  //                                         TaskID: taskId.data.ID,
  //                                       }
  //                                       this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, taskitem)
  //                                         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
  //                                           ({
  //                                             TaskID: taskId.data.ID,
  //                                           }) */
  //                                         .then(async aftermail => {
  //                                           //Email pending  emailbody to approver                 
  //                                           this.triggerDocumentUnderReview(this.sourceDocumentID, "Under Review");
  //                                           // this._adaptiveCard("Review");
  //                                           //await this._adaptiveCard("Review",this.state.approverEmail,this.state.approverName,"General");
  //                                           //aftermail
  //                                           this._sendAnEmailUsingMSGraph(DelegatedTo.Email, "DocReview", DelegatedTo.Title, this.newDetailItemID);
  //                                           this.setState({
  //                                             statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
  //                                             comments: "",
  //                                             statusKey: "",
  //                                             approverEmail: "",
  //                                             approverName: "",
  //                                             approverId: "",
  //                                             buttonHidden: "none"
  //                                           });

  //                                         }).then(redirect => {
  //                                           setTimeout(() => {
  //                                             this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //                                             window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //                                             // this.RedirectUrl;
  //                                           }, 10000);
  //                                         });//aftermail


  //                                     });

  //                                 });//r
  //                             });//Delegated For
  //                         });//Delegated To
  //                     }
  //                     else {
  //                       //detail list adding an item for reviewers
  //                       const detailitem = {
  //                         HeaderIDId: Number(this.headerId),
  //                         Workflow: "Review",
  //                         Title: this.state.documentName,
  //                         ResponsibleId: user.Id,
  //                         OwnerId: this.state.ownerID,
  //                         DueDate: this.state.DueDate,
  //                         ResponseStatus: "Under Review",
  //                         SourceDocument: {
  //                           "__metadata": { type: "SP.FieldUrlValue" },
  //                           Description: this.state.documentName,
  //                           Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                         },
  //                       }
  //                       await this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detailitem)
  //                         /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add({
  //                           HeaderIDId: Number(this.headerId),
  //                           Workflow: "Review",
  //                           Title: this.state.documentName,
  //                           ResponsibleId: user.Id,
  //                           OwnerId: this.state.ownerID,
  //                           DueDate: this.state.DueDate,
  //                           ResponseStatus: "Under Review",
  //                           SourceDocument: {
  //                             "__metadata": { type: "SP.FieldUrlValue" },
  //                             Description: this.state.documentName,
  //                             Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                           },

  //                         }) */
  //                         .then(async r => {
  //                           this.setState({ detailIdForApprover: r.data.ID });
  //                           this.newDetailItemID = r.data.ID;
  //                           const detailitem = {
  //                             Link: {
  //                               "__metadata": { type: "SP.FieldUrlValue" },
  //                               Description: this.state.documentName + "-- Review",
  //                               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                             },
  //                           }
  //                           this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
  //                           /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
  //                             Link: {
  //                               "__metadata": { type: "SP.FieldUrlValue" },
  //                               Description: this.state.documentName + "-- Review",
  //                               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                             },
  //                           }); */
  //                           //MY tasks list updation with delegated from
  //                           const taskdata = {
  //                             Title: "Review '" + this.state.documentName + "'",
  //                             Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.currentDate + "'",
  //                             DueDate: this.state.DueDate,
  //                             StartDate: this.currentDate,
  //                             AssignedToId: hubsieUser.Id,
  //                             Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                             Source: (this.props.project ? "Project" : "QDMS"),
  //                             Workflow: "Review",
  //                             Link: {
  //                               "__metadata": { type: "SP.FieldUrlValue" },
  //                               Description: this.state.documentName + "-- Review",
  //                               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                             },
  //                           }
  //                           await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
  //                             /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add({
  //                               Title: "Review '" + this.state.documentName + "'",
  //                               Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.currentDate + "'",
  //                               DueDate: this.state.DueDate,
  //                               StartDate: this.currentDate,
  //                               AssignedToId: hubsieUser.Id,
  //                               Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                               Source: (this.props.project ? "Project" : "QDMS"),
  //                               Workflow: "Review",
  //                               Link: {
  //                                 "__metadata": { type: "SP.FieldUrlValue" },
  //                                 Description: this.state.documentName + "-- Review",
  //                                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                               },
  //                             }) */
  //                             .then(taskId => {
  //                               const detailitem = {
  //                                 TaskID: taskId.data.ID,
  //                               }
  //                               this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
  //                                 /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
  //                                   ({
  //                                     TaskID: taskId.data.ID,
  //                                   }) */
  //                                 .then(aftermail => {
  //                                   //Email pending  emailbody to approver                 
  //                                   this.triggerDocumentUnderReview(this.sourceDocumentID, "Under Review");
  //                                   //this._adaptiveCard("Review");
  //                                   //aftermail
  //                                   this._sendAnEmailUsingMSGraph(user.Email, "DocReview", user.Title, this.newDetailItemID);
  //                                   this.setState({
  //                                     statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
  //                                     comments: "",
  //                                     statusKey: "",
  //                                     approverEmail: "",
  //                                     approverName: "",
  //                                     approverId: "",
  //                                     buttonHidden: "none"
  //                                   });

  //                                 }).then(redirect => {
  //                                   setTimeout(() => {
  //                                     this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //                                     window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //                                     // this.RedirectUrl;
  //                                   }, 10000);
  //                                 });//aftermail

  //                             });
  //                         });//r
  //                     }
  //                   }//IF
  //                   //If no task delegation
  //                   else {
  //                     //detail list adding an item for reviewers
  //                     const detaildata = {
  //                       HeaderIDId: Number(this.headerId),
  //                       Workflow: "Review",
  //                       Title: this.state.documentName,
  //                       ResponsibleId: user.Id,
  //                       OwnerId: this.state.ownerID,
  //                       DueDate: this.state.DueDate,
  //                       ResponseStatus: "Under Review",
  //                       SourceDocument: {
  //                         "__metadata": { type: "SP.FieldUrlValue" },
  //                         Description: this.state.documentName,
  //                         Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                       },
  //                     }
  //                     await this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detaildata)
  //                       /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add({
  //                         HeaderIDId: Number(this.headerId),
  //                         Workflow: "Review",
  //                         Title: this.state.documentName,
  //                         ResponsibleId: user.Id,
  //                         OwnerId: this.state.ownerID,
  //                         DueDate: this.state.DueDate,
  //                         ResponseStatus: "Under Review",
  //                         SourceDocument: {
  //                           "__metadata": { type: "SP.FieldUrlValue" },
  //                           Description: this.state.documentName,
  //                           Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                         },
  //                       }) */
  //                       .then(async r => {
  //                         this.setState({ detailIdForApprover: r.data.ID });
  //                         this.newDetailItemID = r.data.ID;
  //                         const detailitem = {
  //                           Link: {
  //                             "__metadata": { type: "SP.FieldUrlValue" },
  //                             Description: this.state.documentName + "-- Review",
  //                             Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                           },
  //                         }
  //                         this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
  //                         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
  //                           Link: {
  //                             "__metadata": { type: "SP.FieldUrlValue" },
  //                             Description: this.state.documentName + "-- Review",
  //                             Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                           },
  //                         }); */
  //                         //MY tasks list updation with delegated from
  //                         const taskdata = {
  //                           Title: "Review '" + this.state.documentName + "'",
  //                           Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.currentDate + "'",
  //                           DueDate: this.state.DueDate,
  //                           StartDate: this.currentDate,
  //                           AssignedToId: hubsieUser.Id,
  //                           Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                           Source: (this.props.project ? "Project" : "QDMS"),
  //                           Workflow: "Review",
  //                           Link: {
  //                             "__metadata": { type: "SP.FieldUrlValue" },
  //                             Description: this.state.documentName + "-- Review",
  //                             Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                           },
  //                         }
  //                         await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
  //                           /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add({
  //                             Title: "Review '" + this.state.documentName + "'",
  //                             Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.currentDate + "'",
  //                             DueDate: this.state.DueDate,
  //                             StartDate: this.currentDate,
  //                             AssignedToId: hubsieUser.Id,
  //                             Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                             Source: (this.props.project ? "Project" : "QDMS"),
  //                             Workflow: "Review",
  //                             Link: {
  //                               "__metadata": { type: "SP.FieldUrlValue" },
  //                               Description: this.state.documentName + "-- Review",
  //                               Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                             },
  //                           }) */
  //                           .then(taskId => {
  //                             const detailitem = {
  //                               TaskID: taskId.data.ID,
  //                             }
  //                             this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detailitem)
  //                               /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
  //                                 ({
  //                                   TaskID: taskId.data.ID,
  //                                 }) */
  //                               .then(aftermail => {
  //                                 //Email pending  emailbody to approver                 
  //                                 this.triggerDocumentUnderReview(this.sourceDocumentID, "Under Review");
  //                                 //this._adaptiveCard("Review");

  //                                 //aftermail
  //                                 this._sendAnEmailUsingMSGraph(user.Email, "DocReview", user.Title, this.newDetailItemID);
  //                                 this.setState({
  //                                   statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
  //                                   comments: "",
  //                                   statusKey: "",
  //                                   approverEmail: "",
  //                                   approverName: "",
  //                                   approverId: "",
  //                                   buttonHidden: "none"
  //                                 });

  //                               }).then(redirect => {
  //                                 setTimeout(() => {
  //                                   this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //                                   window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //                                   // this.RedirectUrl;
  //                                 }, 10000);
  //                               });//aftermail

  //                           });
  //                       });//r
  //                   }//else
  //                 });//hubsiteuser
  //             });//user
  //         }

  //       }
  //       //if no reviewers  to approve     
  //       else {
  //         this.setState({
  //           buttonHidden: "none",
  //         });
  //         const headitem = {                   //headerlist
  //           WorkflowStatus: "Under Approval",
  //           Workflow: "Approval",
  //           ReviewedDate: this.currentDate,
  //         }
  //         this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
  //           ({                   //headerlist
  //             WorkflowStatus: "Under Approval",
  //             Workflow: "Approval",
  //             ReviewedDate: this.currentDate,
  //           }); */
  //         const inditem = {
  //           WorkflowStatus: "Under Approval",//docIndex
  //           Workflow: "Approval",
  //         }
  //         this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
  //           ({
  //             WorkflowStatus: "Under Approval",//docIndex
  //             Workflow: "Approval",
  //           }); */
  //         //Updationg DocumentRevisionlog 
  //         const logitem = {
  //           Status: "DCC - Reviewed",
  //           LogDate: this.currentDate,
  //         }
  //         this._Service.updateItemById(this.props.siteUrl, this.props.documentRevisionLog, this.state.revisionLogID, logitem)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.getById(this.state.revisionLogID).update({
  //           Status: "DCC - Reviewed",
  //           LogDate: this.currentDate,
  //         }); */
  //         const logdata = {
  //           Status: "Under Approval",
  //           LogDate: this.currentDate,
  //           WorkflowID: this.headerId,
  //           DocumentIndexId: this.documentIndexId,
  //           DueDate: this.state.DueDate,
  //           Workflow: "Approval",
  //           Revision: this.state.revision,
  //           Title: this.state.documentID,
  //         }
  //         this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLog, logdata)
  //         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLog).items.add({
  //           Status: "Under Approval",
  //           LogDate: this.currentDate,
  //           WorkflowID: this.headerId,
  //           DocumentIndexId: this.documentIndexId,
  //           DueDate: this.state.DueDate,
  //           Workflow: "Approval",
  //           Revision: this.state.revision,
  //           Title: this.state.documentID,
  //         }); */
  //         //upadting source library without version change.            
  //         // let bodyArray = [
  //         //   { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
  //         // ];
  //         // sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem
  //         //   (
  //         //     bodyArray,
  //         //   );
  //         //Task delegation getting user id from hubsite
  //         this._Service.getUserIdByEmail(this.state.approverEmail)
  //           //sp.web.siteUsers.getByEmail(this.state.approverEmail).get()
  //           .then(async user => {
  //             console.log('User Id: ', user.Id);
  //             this.setState({
  //               hubSiteUserId: user.Id,
  //             });
  //             //Task delegation 
  //             const taskDelegation: any[] = await this._Service.getItemSelectExpandFilter(
  //               this.props.siteUrl,
  //               this.props.taskDelegationSettingsListName,
  //               "DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate",
  //               "DelegatedFor,DelegatedTo",
  //               "DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')"
  //             )
  //             //const taskDelegation: any[] = await this._Service.getDelegateAndActive(this.props.siteUrl, this.props.taskDelegationSettingsListName, user.Id)
  //             //const taskDelegation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationSettingsListName).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "' and(Status eq 'Active')").get();
  //             console.log(taskDelegation);
  //             if (taskDelegation.length > 0) {
  //               let duedate = moment(this.dueDateWithoutConversion).toDate();
  //               let toDate = moment(taskDelegation[0].ToDate).toDate();
  //               let fromDate = moment(taskDelegation[0].FromDate).toDate();
  //               duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
  //               toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
  //               fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  //               if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
  //                 this.setState({
  //                   approverEmail: taskDelegation[0].DelegatedTo.EMail,
  //                   approverName: taskDelegation[0].DelegatedTo.Title,
  //                   delegatedToId: taskDelegation[0].DelegatedTo.ID,
  //                   delegatedFromId: taskDelegation[0].DelegatedFor.ID,
  //                 });
  //                 //duedate checking

  //                 //detail list adding an item for approval
  //                 this._Service.getUserIdByEmail(taskDelegation[0].DelegatedTo.EMail)
  //                   //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get()
  //                   .then(async DelegatedTo => {
  //                     this.setState({
  //                       delegateToIdInSubSite: DelegatedTo.Id,
  //                     });
  //                     this._Service.getUserIdByEmail(taskDelegation[0].DelegatedFor.EMail)
  //                       //sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get()
  //                       .then(async DelegatedFor => {
  //                         this.setState({
  //                           delegateForIdInSubSite: DelegatedFor.Id,
  //                         });
  //                         const detaildata = {
  //                           HeaderIDId: Number(this.headerId),
  //                           Workflow: "Approval",
  //                           Title: this.state.documentName,
  //                           ResponsibleId: (this.state.delegatedToId !== "" ? this.state.delegateToIdInSubSite : this.state.approverId),
  //                           DueDate: this.state.DueDate,
  //                           OwnerId: this.state.ownerID,
  //                           DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegateForIdInSubSite : parseInt("")),
  //                           ResponseStatus: "Under Approval",
  //                           SourceDocument: {
  //                             "__metadata": { type: "SP.FieldUrlValue" },
  //                             Description: this.state.documentName,
  //                             Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                           },
  //                         }
  //                         this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detaildata)
  //                           /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
  //                             ({
  //                               HeaderIDId: Number(this.headerId),
  //                               Workflow: "Approval",
  //                               Title: this.state.documentName,
  //                               ResponsibleId: (this.state.delegatedToId !== "" ? this.state.delegateToIdInSubSite : this.state.approverId),
  //                               DueDate: this.state.DueDate,
  //                               OwnerId: this.state.ownerID,
  //                               DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegateForIdInSubSite : parseInt("")),
  //                               ResponseStatus: "Under Approval",
  //                               SourceDocument: {
  //                                 "__metadata": { type: "SP.FieldUrlValue" },
  //                                 Description: this.state.documentName,
  //                                 Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                               },
  //                             }) */
  //                           .then(async r => {
  //                             this.setState({ detailIdForApprover: r.data.ID });
  //                             this.newDetailItemID = r.data.ID;
  //                             const detitem = {
  //                               Link: {
  //                                 "__metadata": { type: "SP.FieldUrlValue" },
  //                                 Description: this.state.documentName + "-- Approve",
  //                                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                               },
  //                             }
  //                             this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
  //                             /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
  //                               Link: {
  //                                 "__metadata": { type: "SP.FieldUrlValue" },
  //                                 Description: this.state.documentName + "-- Approve",
  //                                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                               },
  //                             }); */
  //                             const headitem = {
  //                               ApproverId: this.state.delegateToIdInSubSite
  //                             }
  //                             this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
  //                             /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
  //                               ({                   //headerlist
  //                                 ApproverId: this.state.delegateToIdInSubSite
  //                               }); */
  //                             const iditem = {
  //                               ApproverId: this.state.delegateToIdInSubSite
  //                             }
  //                             this._Service.updateItemById(this.props.siteUrl, this.props.documentIndex, this.documentIndexId, inditem)
  //                             /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndex).items.getById(this.documentIndexId).update
  //                               ({
  //                                 ApproverId: this.state.delegateToIdInSubSite
  //                               }); */
  //                             //upadting source library without version change.   
  //                             const sourceitem = {
  //                               ApproverId: this.state.delegateToIdInSubSite,
  //                             }
  //                             this._Service.updateItemById(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, sourceitem)
  //                             /* sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).update({
  //                               ApproverId: this.state.delegateToIdInSubSite,
  //                             }); */
  //                             //MY tasks list updation
  //                             const taskdata = {
  //                               Title: "Approve '" + this.state.documentName + "'",
  //                               Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
  //                               DueDate: this.state.DueDate,
  //                               StartDate: this.currentDate,
  //                               AssignedToId: (this.state.delegatedToId !== "" ? this.state.delegatedToId : user.Id),
  //                               Workflow: "Approval",
  //                               Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                               DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
  //                               Source: (this.props.project ? "Project" : "QDMS"),
  //                               DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegatedFromId : 0),
  //                               Link: {
  //                                 "__metadata": { type: "SP.FieldUrlValue" },
  //                                 Description: this.state.documentName + "-- Approve",
  //                                 Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                               },

  //                             }
  //                             await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
  //                               /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
  //                                 ({
  //                                   Title: "Approve '" + this.state.documentName + "'",
  //                                   Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
  //                                   DueDate: this.state.DueDate,
  //                                   StartDate: this.currentDate,
  //                                   AssignedToId: (this.state.delegatedToId !== "" ? this.state.delegatedToId : user.Id),
  //                                   Workflow: "Approval",
  //                                   Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                                   DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
  //                                   Source: (this.props.project ? "Project" : "QDMS"),
  //                                   DelegatedFromId: (this.state.delegatedToId !== "" ? this.state.delegatedFromId : 0),
  //                                   Link: {
  //                                     "__metadata": { type: "SP.FieldUrlValue" },
  //                                     Description: this.state.documentName + "-- Approve",
  //                                     Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                                   },

  //                                 }) */
  //                               .then(taskId => {
  //                                 const detitem = {
  //                                   TaskID: taskId.data.ID,
  //                                 }
  //                                 this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
  //                                   /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
  //                                     ({
  //                                       TaskID: taskId.data.ID,
  //                                     }) */
  //                                   .then(mailSend => {
  //                                     //notification preference checking  
  //                                     // this._adaptiveCard("Review");
  //                                     this.triggerDocumentReview(this.sourceDocumentID, "Under Approval")

  //                                       .then(aftermail => {
  //                                         //Email pending  emailbody to approver                 
  //                                         this.validator.hideMessages();
  //                                         this.setState({
  //                                           comments: "",
  //                                           statusKey: "",
  //                                           approverEmail: "",
  //                                           approverName: "",
  //                                           approverId: "",
  //                                           buttonHidden: "none"
  //                                         });
  //                                       }).then(redirect => {
  //                                         setTimeout(() => {
  //                                           this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //                                           window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //                                           // this.RedirectUrl;
  //                                         }, 10000);

  //                                       });//aftermail
  //                                     this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
  //                                   });
  //                               });//taskID
  //                           });//r

  //                       });//DelegatedFor
  //                   });//DelegatedTo
  //               }
  //               else {
  //                 const detdata = {
  //                   HeaderIDId: Number(this.headerId),
  //                   Workflow: "Approval",
  //                   Title: this.state.documentName,
  //                   ResponsibleId: this.state.approverId,
  //                   OwnerId: this.state.ownerID,
  //                   DueDate: this.state.DueDate,
  //                   ResponseStatus: "Under Approval",
  //                   SourceDocument: {
  //                     "__metadata": { type: "SP.FieldUrlValue" },
  //                     Description: this.state.documentName,
  //                     Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                   },
  //                 }

  //                 this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detdata)
  //                   /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
  //                     ({
  //                       HeaderIDId: Number(this.headerId),
  //                       Workflow: "Approval",
  //                       Title: this.state.documentName,
  //                       ResponsibleId: this.state.approverId,
  //                       OwnerId: this.state.ownerID,
  //                       DueDate: this.state.DueDate,
  //                       ResponseStatus: "Under Approval",
  //                       SourceDocument: {
  //                         "__metadata": { type: "SP.FieldUrlValue" },
  //                         Description: this.state.documentName,
  //                         Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                       },
  //                     }) */
  //                   .then(async r => {
  //                     this.setState({ detailIdForApprover: r.data.ID });
  //                     this.newDetailItemID = r.data.ID;
  //                     const detitem = {
  //                       Link: {
  //                         "__metadata": { type: "SP.FieldUrlValue" },
  //                         Description: this.state.documentName + "-- Approve",
  //                         Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                       },
  //                     }
  //                     this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
  //                     /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
  //                       Link: {
  //                         "__metadata": { type: "SP.FieldUrlValue" },
  //                         Description: this.state.documentName + "-- Approve",
  //                         Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                       },
  //                     }); */

  //                     //MY tasks list updation
  //                     const taskdata = {
  //                       Title: "Approve '" + this.state.documentName + "'",
  //                       Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
  //                       DueDate: this.state.DueDate,
  //                       StartDate: this.currentDate,
  //                       AssignedToId: user.Id,
  //                       Workflow: "Approval",
  //                       Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                       Source: (this.props.project ? "Project" : "QDMS"),
  //                       Link: {
  //                         "__metadata": { type: "SP.FieldUrlValue" },
  //                         Description: this.state.documentName + "-- Approve",
  //                         Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                       },

  //                     }
  //                     await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
  //                       /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
  //                         ({
  //                           Title: "Approve '" + this.state.documentName + "'",
  //                           Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
  //                           DueDate: this.state.DueDate,
  //                           StartDate: this.currentDate,
  //                           AssignedToId: user.Id,
  //                           Workflow: "Approval",
  //                           Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                           Source: (this.props.project ? "Project" : "QDMS"),
  //                           Link: {
  //                             "__metadata": { type: "SP.FieldUrlValue" },
  //                             Description: this.state.documentName + "-- Approve",
  //                             Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                           },

  //                         }) */
  //                       .then(taskId => {
  //                         const detitem = {
  //                           TaskID: taskId.data.ID,
  //                         }
  //                         this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
  //                           /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
  //                             ({
  //                               TaskID: taskId.data.ID,
  //                             }) */
  //                           .then(mailSend => {
  //                             //notification preference checking  
  //                             //this._adaptiveCard("Review");
  //                             this.triggerDocumentReview(this.sourceDocumentID, "Under Approval")
  //                               .then(aftermail => {
  //                                 //Email pending  emailbody to approver    
  //                                 this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
  //                                 this.validator.hideMessages();
  //                                 this.setState({
  //                                   statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
  //                                   comments: "",
  //                                   statusKey: "",
  //                                   approverEmail: "",
  //                                   approverName: "",
  //                                   approverId: "",
  //                                   buttonHidden: "none"
  //                                 });
  //                               }).then(redirect => {
  //                                 setTimeout(() => {
  //                                   this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //                                   window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //                                   // this.RedirectUrl;
  //                                 }, 10000);
  //                               });//aftermail
  //                           });
  //                       });//taskID
  //                   });//r
  //               }//else no delegation

  //             }

  //             else {
  //               const detdata = {
  //                 HeaderIDId: Number(this.headerId),
  //                 Workflow: "Approval",
  //                 Title: this.state.documentName,
  //                 ResponsibleId: this.state.approverId,
  //                 OwnerId: this.state.ownerID,
  //                 DueDate: this.state.DueDate,
  //                 ResponseStatus: "Under Approval",
  //                 SourceDocument: {
  //                   "__metadata": { type: "SP.FieldUrlValue" },
  //                   Description: this.state.documentName,
  //                   Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                 },
  //               }
  //               this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detdata)
  //                 /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.add
  //                   ({
  //                     HeaderIDId: Number(this.headerId),
  //                     Workflow: "Approval",
  //                     Title: this.state.documentName,
  //                     ResponsibleId: this.state.approverId,
  //                     OwnerId: this.state.ownerID,
  //                     DueDate: this.state.DueDate,
  //                     ResponseStatus: "Under Approval",
  //                     SourceDocument: {
  //                       "__metadata": { type: "SP.FieldUrlValue" },
  //                       Description: this.state.documentName,
  //                       Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
  //                     },
  //                   }) */
  //                 .then(async r => {
  //                   this.setState({ detailIdForApprover: r.data.ID });
  //                   this.newDetailItemID = r.data.ID;
  //                   const detitem = {
  //                     Link: {
  //                       "__metadata": { type: "SP.FieldUrlValue" },
  //                       Description: this.state.documentName + "-- Approve",
  //                       Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                     },
  //                   }
  //                   this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
  //                   /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update({
  //                     Link: {
  //                       "__metadata": { type: "SP.FieldUrlValue" },
  //                       Description: this.state.documentName + "-- Approve",
  //                       Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                     },
  //                   }); */

  //                   //MY tasks list updation
  //                   const taskdata = {
  //                     Title: "Approve '" + this.state.documentName + "'",
  //                     Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
  //                     DueDate: this.state.DueDate,
  //                     StartDate: this.currentDate,
  //                     AssignedToId: user.Id,
  //                     Workflow: "Approval",
  //                     Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                     Source: (this.props.project ? "Project" : "QDMS"),
  //                     Link: {
  //                       "__metadata": { type: "SP.FieldUrlValue" },
  //                       Description: this.state.documentName + "-- Approve",
  //                       Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                     },
  //                   }
  //                   await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskdata)
  //                     /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
  //                       ({
  //                         Title: "Approve '" + this.state.documentName + "'",
  //                         Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
  //                         DueDate: this.state.DueDate,
  //                         StartDate: this.currentDate,
  //                         AssignedToId: user.Id,
  //                         Workflow: "Approval",
  //                         Priority: (this.state.criticalDocument === true ? "Critical" : ""),
  //                         Source: (this.props.project ? "Project" : "QDMS"),
  //                         Link: {
  //                           "__metadata": { type: "SP.FieldUrlValue" },
  //                           Description: this.state.documentName + "-- Approve",
  //                           Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
  //                         },

  //                       }) */
  //                     .then(taskId => {
  //                       const detitem = {
  //                         TaskID: taskId.data.ID,
  //                       }
  //                       this._Service.updateItemById(this.props.siteUrl, this.props.workFlowDetail, r.data.ID, detitem)
  //                         /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.getById(r.data.ID).update
  //                           ({
  //                             TaskID: taskId.data.ID,
  //                           }) */
  //                         .then(mailSend => {
  //                           //notification preference checking  
  //                           // this._adaptiveCard("Approval");
  //                           this.triggerDocumentReview(this.sourceDocumentID, "Under Approval")
  //                             .then(aftermail => {
  //                               //Email pending  emailbody to approver    
  //                               this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
  //                               this.validator.hideMessages();
  //                               this.setState({
  //                                 statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
  //                                 comments: "",
  //                                 statusKey: "",
  //                                 approverEmail: "",
  //                                 approverName: "",
  //                                 approverId: "",
  //                                 buttonHidden: "none"
  //                               });
  //                             }).then(redirect => {
  //                               setTimeout(() => {
  //                                 this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
  //                                 window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //                                 //this.RedirectUrl;
  //                               }, 10000);
  //                             });//aftermail
  //                         });
  //                     });//taskID
  //                 });//r
  //             }//else no delegation

  //           }).catch(reject => console.error('Error getting Id of user by Email ', reject));

  //       }
  //     }
  //   }

  //   else {
  //     this.validator.showMessages();
  //     this.forceUpdate();
  //   }
  // }
  // sending Email
  private async _sendAnEmailUsingMSGraph(email, type, name, detailID): Promise<void> {
    let Subject;
    let Body;
    let link;
    let tableHeader;
    let tableFooter;
    let tableBody = "";
    let finalBody;
    let DocumentLink;
    //console.log(queryVar);
    // const notificationPreference: any[] = await this._Service.getItemSelectExpandFilter(
    //   this.props.siteUrl,
    //   this.props.notificationPrefListName,
    //   "Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail",
    //   "EmailUser",
    //   "EmailUser/EMail eq '" + email + "'"
    // )
    //const notificationPreference: any[] = await this._Service.getEmailUserandPreference(this.props.siteUrl, this.props.notificationPrefListName, email)
    //const notificationPreference: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.notificationPrefListName).items.select("Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail").expand("EmailUser").filter("EmailUser/EMail eq '" + email + "'").get();
    // console.log(notificationPreference);
    // if (notificationPreference.length > 0) {
    //   this.setState({
    //     notificationPreference: notificationPreference[0].Preference,
    //   });
    // }
    // else
     if (this.state.criticalDocument === true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await this._Service.getItemFilter(this.props.siteUrl, this.props.emailNotificationSettings, "Title eq '" + type + "'")
    //const emailNoficationSettings: any[] = await this._Service.getItemTitleFilter(this.props.siteUrl, this.props.emailNotificationSettings, type)
    //const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNotificationSettings).items.filter("Title eq '" + type + "'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;

    if (type === "DocApproval") {
      link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + detailID}>Link</a>`;
      //for binding current reviewers comments in table
      //       if (this.props.project) {
      //         await this._Service.getItemSelectExpandFilter(
      //           this.props.siteUrl,
      //           this.props.workFlowDetail,
      //           "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow",
      //           "Responsible",
      //           "HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review')"
      //         )
      //           //await this._Service.getWorkflowReviewDCCReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
      //           //await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
      //           //.expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review')").get()
      //           .then(currentReviewersItems => {
      //             console.log("currentReviewersItems", currentReviewersItems);
      //             if (currentReviewersItems.length > 0) {
      //               console.log("currentReviewersItems", currentReviewersItems);
      //               this.setState({
      //                 currentReviewComment: "",
      //                 currentReviewItems: currentReviewersItems,
      //               });
      //               currentReviewersItems.map((item) => {
      //                 tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
      //               });
      //             }
      //           }).then(after => {
      //             tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;v">
      //    <tr  style="background-color: #002d71;     color: white;text-align: center;">
      //    <th >Reviewer</th>
      //    <th >Review Date</th>
      //    <th >Response Status</th>
      //    <th >Review Comment</th>
      //    <th >Workflow</th>
      //  </tr>
      //  <tbody style ="width: 100%;
      //  border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

      //             tableFooter = `</tbody>
      //  </table>`;
      //             finalBody = tableHeader + tableBody + tableFooter;
      //           });
      //       }
      //       else
      {
        await this._Service.getItemSelectExpandFilter(
          this.props.siteUrl,
          this.props.workFlowDetail,
          "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow",
          "Responsible",
          "HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review') "
        )
          //await this._Service.getDetailWorkflowReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          //await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
          //.expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review')").get()
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
              });
              currentReviewersItems.map((item) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>
 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
          });
      }
    }
    else if (type === "DocReview") {
      link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + detailID}>Link</a>`;
      //       if (this.props.project) {
      //         await this._Service.getItemSelectExpandFilter(
      //           this.props.siteUrl,
      //           this.props.workFlowDetail,
      //           "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow",
      //           "Responsible",
      //           "HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review') and (ResponseStatus ne 'Under Review') "
      //         )
      //           //await this._Service.getResponseStatusNeUnderReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
      //           //await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
      //           //.expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review') and (ResponseStatus ne 'Under Review') ").get()
      //           .then(currentReviewersItems => {
      //             console.log("currentReviewersItems", currentReviewersItems);
      //             if (currentReviewersItems.length > 0) {
      //               console.log("currentReviewersItems", currentReviewersItems);
      //               this.setState({
      //                 currentReviewComment: "",
      //                 currentReviewItems: currentReviewersItems,
      //               });
      //               currentReviewersItems.map((item) => {
      //                 tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
      //               });
      //             }
      //           }).then(after => {
      //             tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
      //    <tr  style="background-color: #002d71;     color: white;text-align: center;">
      //    <th >Reviewer</th>
      //    <th >Review Date</th>
      //    <th >Response Status</th>
      //    <th >Review Comment</th>
      //    <th >Workflow</th>
      //  </tr>
      //  <tbody style ="width: 100%;
      //  border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

      //             tableFooter = `</tbody>
      //  </table>`;
      //             finalBody = tableHeader + tableBody + tableFooter;
      //           });
      //       }
    }
    //returned with comments mail body
    else if (type === "DocReturn") {
      //       if (this.props.project) {
      //         await this._Service.getItemSelectExpandFilter(
      //           this.props.siteUrl,
      //           this.props.workFlowDetail,
      //           "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,SourceDocument,Workflow",
      //           "Responsible",
      //           "HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review')"
      //         )
      //           //await this._Service.getWorkflowReviewDCCReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
      //           //await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,SourceDocument,Workflow")
      //           //.expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review')").get()
      //           .then(currentReviewersItems => {
      //             console.log("currentReviewersItems", currentReviewersItems);
      //             if (currentReviewersItems.length > 0) {
      //               console.log("currentReviewersItems", currentReviewersItems);
      //               this.setState({
      //                 currentReviewComment: "",
      //                 currentReviewItems: currentReviewersItems,
      //                 linkToDoc: currentReviewersItems[0].SourceDocument.Url,
      //               });
      //               currentReviewersItems.map((item) => {
      //                 tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
      //               });
      //             }
      //           }).then(after => {
      //             tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
      //    <tr  style="background-color: #002d71;     color: white;text-align: center;">
      //    <th >Reviewer</th>
      //    <th >Review Date</th>
      //    <th >Response Status</th>
      //    <th >Review Comment</th>
      //    <th >Workflow</th>
      //  </tr>
      //  <tbody style ="width: 100%;
      //  border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

      //             tableFooter = `</tbody>
      //  </table>`;
      //             finalBody = tableHeader + tableBody + tableFooter;
      //             DocumentLink = `<a href=${this.state.linkToDoc}>Document Link </a>`;
      //           });
      //       }
      //       else
      {
        await this._Service.getItemSelectExpandFilter(
          this.props.siteUrl,
          this.props.workFlowDetail,
          "Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,SourceDocument,Workflow",
          "Responsible",
          "HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review')"
        )
          //await this._Service.getWorkflowReviewDCCReview(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          //await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workFlowDetail).items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,SourceDocument,Workflow")
          //.expand("Responsible").filter("HeaderID eq '" + this.headerId + "' and (Workflow eq 'Review')").get()
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
                linkToDoc: currentReviewersItems[0].SourceDocument.Url,
              });
              currentReviewersItems.map((item) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>

 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
            DocumentLink = `<a href=${this.state.linkToDoc}>Click here </a>`;
          });
      }
    }

    //Replacing the email body with current values
    let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
    let replacedSubjectWithDueDate = replaceString(replacedSubject, '[DueDate]', this.state.dueDate);
    let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
    let replacelink = replaceString(replaceBody, '[Link]', link);
    let var1: any[] = replacelink.split('/');
    let FinalBody = replacelink;
    if (this.state.notificationPreference === "Send all emails") {
      this.status = "Yes";
      //console.log("Send mail for all");                 
    }
    else if (this.state.notificationPreference === "Send mail for critical document" && this.state.criticalDocument === true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    else {
      this.setState({
        statusMessage: { isShowMessage: true, message: strings.DocumentReviewedMsgBar, messageType: 4 },
        comments: "",
        statusKey: "",
      });
    }
    //mail sending
    if (this.status === "Yes") {
      //Check if TextField value is empty or not  
      if (email) {
        //Create Body for Email  
        let emailPostBody: any = {
          "message": {
            "subject": replacedSubjectWithDueDate,
            "body": {
              "contentType": "HTML",
              "content": FinalBody + "<br></br>" + (type === "DocReturn" ? DocumentLink : "") + "<br></br>" + finalBody
            },
            "toRecipients": [
              {
                "emailAddress": {
                  "address": email
                }
              }
            ],
          }
        };
        //Send Email uisng MS Graph  
        this._Service.sendMail(emailPostBody)
        /* this.props.context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient): void => {
            client
              .api('/me/sendMail')
              .post(emailPostBody, (error, response: any, rawResponse?: any) => {
              });
          }); */
      }
    }
  }
 /*  protected async triggerDocumentReview(sourceDocumentID, ResponseStatus) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': ResponseStatus
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  } */

  protected async triggerDocumentReview(sourceDocumentID, ResponseStatus) {
    // const laUrl = await this._Service.DocumentSendForReview(this.props.siteUrl, this.props.requestList);
    const laUrl = await this._Service.getItemFilter(this.props.siteUrl, this.props.requestListName, "Title eq 'QDMS_DocumentPermission_UnderApproval'")
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': ResponseStatus
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }


  protected async triggerDocumentUnderReview(sourceDocumentID, ResponseStatus) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
    console.log(siteUrl);
    const postURL = this.postUrlForUnderReview;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': ResponseStatus,
      'Workflow': "Review"
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //Cancel button click
  private _cancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });
    this.validator.hideMessages();
  }
  //confirm cancel button click
  private _confirmYesCancel = () => {
    window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
    this.setState({
      statusKey: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });

    this.validator.hideMessages();

  }
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
  }
  //access denied msgbar close button click
  private _closeButton = () => {
    window.location.replace(this.props.redirectUrl);
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

  public render(): React.ReactElement<IDocumentReviewProps> {
    const Status: IDropdownOption[] = [
      { key: 'Reviewed', text: 'Reviewed' },
      { key: 'Returned with comments', text: 'Returned with comments' },
    ];

    return (
      <section className={`${styles.documentReview}`}>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        <div style={{ display: this.state.access }}>

          <div className={styles.border}>
            <div className={styles.alignCenter}> {this.props.webPartName}</div>

            <div className={styles.header}>
              <div className={styles.divMetadataCol1}>
                <h3 >Document Details</h3>
                <Link onClick={this._openRevisionHistory} target="_blank" underline style={{ marginLeft: "70%" }}>Revision History</Link>
                {/* <Link onClick={this.RevisionHistoryUrl} target="_blank" underline style={{ marginLeft: "70%" }}>Revision History</Link> */}
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
                <Label >Owner : </Label><div className={styles.divLabel}> {this.state.owner}</div>
              </div>
              <div className={styles.divMetadataCol2}><Label>Due Date :</Label> <div className={styles.divLabel}> {this.state.dueDate}</div></div>
              <div className={styles.divMetadataCol3}><Label>Requested Date :</Label><div className={styles.divLabel}>{this.state.requestorDate} </div></div>
            </div>
            <div className={styles.divMetadata}>
              <div className={styles.divMetadataCol1}>
                <Label >Requester :</Label> <div className={styles.divLabel}>{this.state.requestor}</div>
              </div>
              <div className={styles.divMetadataCol2}><Label>Requester Comment: </Label><div className={styles.divLabel}>{this.state.requestorComment}</div></div>
            </div>
            <div >
              <div style={{ display: this.state.hideReviewersTable }}>
                <Accordion atomic={true}>
                  <AccordionItem title="Previous Review Details">
                    <table className={styles.tableClass}>
                      <tr className={styles.tr}>
                        <th className={styles.th}>Reviewer</th>
                        <th className={styles.th}>Review Date</th>
                        <th className={styles.th}>Review Comment</th>
                      </tr>
                      <tbody className={styles.tbody}>
                        {this.state.reviewerItems.map((item, key) => {
                          return (<tr className={styles.tr}>
                            <td className={styles.th}>{item.Responsible.Title}</td>
                            <td className={styles.th}>{moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                            <td className={styles.th}>{item.ResponsibleComment}</td>
                          </tr>);
                        })}
                      </tbody>
                    </table>
                  </AccordionItem>
                </Accordion>
              </div>
              <div style={{ display: this.state.currentReviewComment }}>
                <Accordion atomic={true}>
                  <AccordionItem title="Reviewers Details">
                    <table className={styles.tableClass}>
                      <tr className={styles.tr}>
                        <th className={styles.th}>Reviewer</th>
                        <th className={styles.th}>Review Date</th>
                        <th className={styles.th}>Response Status</th>
                        <th className={styles.th}>Review Comment</th>
                      </tr>
                      <tbody className={styles.tbody}>
                        {this.state.currentReviewItems.map((item, key) => {
                          return (<tr className={styles.tr}>
                            <td className={styles.th}>{item.Responsible.Title}</td>
                            <td className={styles.th}>{(item.ResponseDate === null) ? "Not Reviewed Yet" : moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                            <td className={styles.th}>{item.ResponseStatus}</td>
                            <td className={styles.th}>{item.ResponsibleComment}</td>
                          </tr>);
                        })}
                      </tbody>
                    </table>
                  </AccordionItem>
                </Accordion>

              </div>
            </div>
            <div className={styles.header}>
              <h3 className="ExampleCard-title title-222"></h3>
            </div>
            <div className={styles.divMetadata}>
              <div style={{ width: "100%", }}>
                <Dropdown
                  placeholder="Select Status"
                  label="Status"
                  options={Status}
                  onChanged={this._drpdwnStatus}
                  selectedKey={this.state.statusKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("status", this.state.statusKey, "required")}{" "}</div>
              </div>
              <div style={{ width: "100%", marginLeft: "12px" }}>
                <TextField label="Comments" required={this.state.statusKey === "Returned with comments"} id="Comments" value={this.state.comments} onChange={this._commentBoxChange} multiline autoAdjustHeight />
                {this.state.statusKey === "Returned with comments" && <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>}
              </div>
            </div>

            {/* Show Message bar for Notification*/}
            <div>
              {this.state.statusMessage.isShowMessage ?
                <MessageBar
                  messageBarType={this.state.statusMessage.messageType}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >{this.state.statusMessage.message}</MessageBar>
                : ''}
            </div>
            <div className={styles.divRow}>
              <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>

              <div className={styles.rgtalign} >
                <PrimaryButton id="b2" className={styles.btn} onClick={this._docReviewSaveAsDraft} style={{ display: this.state.buttonHidden }}>Save as Draft</PrimaryButton >
                <PrimaryButton id="b2" className={styles.btn} onClick={this._docReviewSubmit} style={{ display: this.state.buttonHidden }}>Submit</PrimaryButton >
                <PrimaryButton id="b1" className={styles.btn} onClick={this._cancel}>Cancel</PrimaryButton >
              </div>
            </div>
            {/* Cancel Dialog Box */}
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
          <MessageBar messageBarType={MessageBarType.error}
            onDismiss={this._closeButton}
            isMultiline={false}>
            {this.state.invalidMessage}</MessageBar>

        </div>
      </section>
    );
  }
}
