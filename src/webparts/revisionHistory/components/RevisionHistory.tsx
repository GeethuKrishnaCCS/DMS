import * as React from 'react';
import styles from './RevisionHistory.module.scss';
import type { IRevisionHistoryProps, IRevisionHistoryState } from '../interfaces';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dialog, DialogFooter, DialogType, ITooltipHostStyles, Modal, ProgressIndicator, DefaultButton, FontWeights, getTheme, IconButton, IIconProps, Label, mergeStyleSets, MessageBar, PrimaryButton, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import * as moment from 'moment';
import * as _ from 'lodash';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import replaceString from 'replace-string';
import { add, groupBy } from 'lodash';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient, HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Timeline, TimelineItem } from 'vertical-timeline-component-for-react';
import { DMSService } from '../services';

//for modal popup
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const ReminderTime: IIconProps = { iconName: 'ReminderTime' };
const Comment: IIconProps = { iconName: 'CommentActive' };
const Share: IIconProps = { iconName: 'Share' };
const theme = getTheme();
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',

  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      //borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 284px',
    },
  ],
  header1: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      // borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '10px 20px',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 20px 20px ',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const iconButtonStyles = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};

export default class RevisionHistory extends React.Component<IRevisionHistoryProps, IRevisionHistoryState> {
  private _Service: DMSService;
  private documentIndexID;
  private headerId;
  private documentRevisionLogID;
  private sourceDocumentID;
  private status;
  private newDetailItemID;
  private currentUserEmail;
  private statusForRemainder = "No";
  private statusForCancel = "No";
  private statusForDelegate = "No";
  private currentDate = new Date();
  private dueDateForModel;
  private departmentExist;
  private postUrl;
  private postUrlForUnderReview;
  private postUrlForDelegate;
  private loaderforcancel = "none";
  constructor(props: IRevisionHistoryProps) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      currentUserEmail: "",
      currentUser: 0,
      logItems: [],
      documentName: "",
      documentIndexItems: [],
      workflowDetailItems: [],
      notificationPreference: "",
      criticalDocument: false,
      timelineElement: "",
      dueDate: "",
      workflowStatus: "",
      approverEmail: "",
      approverName: "",
      detailIdForApprover: "",
      requestorEmail: "",
      requestorDate: "",
      ownerEmail: "",
      owner: "",
      approver: "",
      requestor: "",
      documentID: "",
      revision: "",
      hubSiteUserId: "",
      delegatedToId: "",
      delegatedFromId: "",
      delegateToIdInSubSite: "",
      delegateToEmail: "",
      delegateForIdInSubSite: "",
      iframeModalclose: true,
      tableShow: "none",
      tableinTimeLine: "none",
      showModal: false,
      reviewed: "none",
      workflowInitiated: "none",
      showworkflowInitiatedModal: false,
      showReviewModal: false,
      delegateUser: "",
      delagatePeoplePicker: "none",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      cancelledBy: "",
      reviewers: [],
      divForSendRemainder: true,
      divForCancel: true,
      divForDelegation: true,
      workflowInitiatedVoid: "",
      showworkflowInitiatedVoidModal: false,
      delegateToTitle: "",
      delegatingLoader: "none",
      timelineDisplay: "none",
      ownerID: "",
      lengthOfReviwers: "",
      taskOwnerName: "",
      currentUserName: "",
    };
    this._Service = new DMSService(this.props.context);
    this.timelineLoad = this.timelineLoad.bind(this);
    this.documentCreated = this.documentCreated.bind(this);
    this.workFlowStarted = this.workFlowStarted.bind(this);
    this.reviewedDetailsItems = this.reviewedDetailsItems.bind(this);
    this._delegateClick = this._delegateClick.bind(this);
    this._closeModal = this._closeModal.bind(this);
    this.documentExpired = this.documentExpired.bind(this);
    this.sendRemainder = this.sendRemainder.bind(this);
    this.cancelTask = this.cancelTask.bind(this);
    this._delegateSubmit = this._delegateSubmit.bind(this);
    this._delegateClick = this._delegateClick.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._accessGroups = this._accessGroups.bind(this);
    this._checkingCurrent = this._checkingCurrent.bind(this);
    this._gettingGroupID = this._gettingGroupID.bind(this);
    this.GetGroupMembers = this.GetGroupMembers.bind(this);
    this.departmentOrBUIdGetting = this.departmentOrBUIdGetting.bind(this);
    this.GetGroupMembersForProject = this.GetGroupMembersForProject.bind(this);
    this._checkingCurrentForProject = this._checkingCurrentForProject.bind(this);
    this._LAUrlGetting = this._LAUrlGetting.bind(this);
    this._LAUrlGettingForUnderReview = this._LAUrlGettingForUnderReview.bind(this);
    this.triggerDocumentReview = this.triggerDocumentReview.bind(this);
    this.triggerDocumentDelegate = this.triggerDocumentDelegate.bind(this);
    this._LAUrlGettingForDelegate = this._LAUrlGettingForDelegate.bind(this);
    this.triggerDocumentReviewWithWorkFlowStatus = this.triggerDocumentReviewWithWorkFlowStatus.bind(this);
    this.sendEmailForCurrentTaskOwner = this.sendEmailForCurrentTaskOwner.bind(this);
    this.currentUser = this.currentUser.bind(this);
    this.internalTransittalConFab = this.internalTransittalConFab.bind(this);
  }

  public async componentDidMount() {
    this.queryParamGetting();
    this.timelineLoad();
    let groups = await this._Service.getCurrentUser()
    //let groups = await sp.web.currentUser.get();
    console.log(groups.Email);
    this.currentUserEmail = groups.Email;
    this._accessGroups();
    this._LAUrlGetting();
    this._LAUrlGettingForUnderReview();
    this._LAUrlGettingForDelegate();
    this.currentUser();
  }
  //Get collection of SharePoint Groups for the current User    
  private queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let documentIndexID = params.get('did');
    if (documentIndexID != "" && documentIndexID != null) {
      this.documentIndexID = documentIndexID;
      console.log("document index id", this.documentIndexID);
    }
  }
  //Get Access Groups
  private async _accessGroups() {
    let AccessGroup: any[] = [];
    let AccessGroupForCancel: any[] = [];
    let AccessGroupForDelegate: any[] = [];
    let ok = "No";
    if (this.props.project) {
      AccessGroupForCancel = await this._Service.getProject_CancelWF(this.props.siteUrl, this.props.permissionMatrixSettings);
      //AccessGroupForCancel = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.permissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'Project_CancelWF'").get();
      AccessGroup = await this._Service.getProject_SendReminderWFTasks(this.props.siteUrl, this.props.permissionMatrixSettings)
      //AccessGroup = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.permissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'Project_SendReminderWFTasks'").get();
      AccessGroupForDelegate = await this._Service.getProject_DelegateWFTask(this.props.siteUrl, this.props.permissionMatrixSettings);
      //AccessGroupForDelegate = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.permissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'Project_DelegateWFTask'").get();

    }
    else {
      AccessGroup = await this._Service.getQDMS_SendReminderWFTasks(this.props.siteUrl, this.props.permissionMatrixSettings)
      //AccessGroup = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.permissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendReminderWFTasks'").get();
      AccessGroupForCancel = await this._Service.getQDMS_CancelWF(this.props.siteUrl, this.props.permissionMatrixSettings);
      //AccessGroupForCancel = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.permissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_CancelWF'").get();
      AccessGroupForDelegate = await this._Service.getQDMS_DelegateWFTask(this.props.siteUrl, this.props.permissionMatrixSettings)
      //AccessGroupForDelegate = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.permissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_DelegateWFTask'").get();
    }
    let AccessGroupItems: any[] = AccessGroup[0].AccessGroups.split(',');
    let AccessGroupItemsForCancel: any[] = AccessGroupForCancel[0].AccessGroups.split(',');
    let AccessGroupItemsForDelegate: any[] = AccessGroupForDelegate[0].AccessGroups.split(',');
    console.log("AccessGroupItems", AccessGroupItems);
    console.log("AccessGroupItemsForCancel", AccessGroupItemsForCancel);
    console.log("AccessGroupItemsForDelegate", AccessGroupItemsForDelegate);
    if (this.props.project) {
      if (AccessGroupItemsForCancel.length > 0) {
        this.statusForCancel = "Yes";
        this._gettingGroupID(AccessGroupItemsForCancel);
      }
      if (AccessGroupItems.length > 0) {
        this.statusForRemainder = "Yes";
        this._gettingGroupID(AccessGroupItems);
      }

      if (AccessGroupItemsForDelegate.length > 0) {
        this.statusForDelegate = "Yes";
        this._gettingGroupID(AccessGroupItemsForDelegate);

      }
    }
    else {
      if (AccessGroupItems.length > 0) {
        this.statusForRemainder = "Yes";
        this.departmentOrBUIdGetting(AccessGroupItems);
      }
      if (AccessGroupItemsForCancel.length > 0) {
        // this._gettingGroupID(AccessGroupItemsForCancel);
        this.departmentOrBUIdGetting(AccessGroupItemsForCancel);
        this.statusForCancel = "Yes";
      }
      if (AccessGroupItemsForDelegate.length > 0) {
        // this._gettingGroupID(AccessGroupItemsForDelegate);
        this.departmentOrBUIdGetting(AccessGroupItemsForDelegate);
        this.statusForDelegate = "Yes";
      }

    }
  }
  private async departmentOrBUIdGetting(AccessGroupItems) {
    console.log("AccessGroupItems", AccessGroupItems);
    const DocumentIndexItem: any = await this._Service.getDocumentIndexItem(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID)
    //const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).select("DepartmentID,BusinessUnitID").get();
    console.log("DocumentIndexItem", DocumentIndexItem);
    // this._gettingGroupID(AccessGroupItems);
    //cheching if department selected
    if (DocumentIndexItem.DepartmentID != null) {
      //this.departmentExist == "Exists";
      let deptid = parseInt(DocumentIndexItem.DepartmentID);
      const departmentItem: any = await this._Service.getItemById(this.props.siteUrl, this.props.departmentListName, deptid)
      //const departmentItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.departmentListName).items.getById(deptid).get();
      //let AG = DepartmentItem[0].AccessGroups;
      console.log("departmentItem", departmentItem);
      let accessGroupvar = departmentItem.AccessGroups;
      const accessGroupItem: any = await this._Service.getItems(this.props.siteUrl, this.props.accessGroupDetailsListName);
      //const accessGroupItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroupDetailsListName).items.get();
      let accessGroupID;
      console.log(accessGroupItem.length);
      for (let a = 0; a < accessGroupItem.length; a++) {
        if (accessGroupItem[a].Title == accessGroupvar) {
          accessGroupID = accessGroupItem[a].GroupID;
          this.GetGroupMembers(this.props.context, accessGroupID,);
        }
      }

    }
    //if no department
    else {
      //alert("with bussinessUnit");
      if (DocumentIndexItem.BusinessUnitID != null) {
        this.departmentExist == "Exists";
        let bussinessUnitID = parseInt(DocumentIndexItem.BusinessUnitID);
        const bussinessUnitItem: any = await this._Service.getItemById(this.props.siteUrl, this.props.bussinessUnitList, bussinessUnitID)
        //const bussinessUnitItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.bussinessUnitList).items.getById(bussinessUnitID).get();
        console.log("departmentItem", bussinessUnitItem);
        let accessGroupvar = bussinessUnitItem.AccessGroups;
        // alert(accessGroupvar);
        const accessGroupItem: any = await this._Service.getItems(this.props.siteUrl, this.props.accessGroupDetailsListName)
        //const accessGroupItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroupDetailsListName).items.get();
        let accessGroupID;
        console.log(accessGroupItem.length);
        for (let a = 0; a < accessGroupItem.length; a++) {
          if (accessGroupItem[a].Title == accessGroupvar) {
            accessGroupID = accessGroupItem[a].GroupID;
            this.GetGroupMembers(this.props.context, accessGroupID);
          }
        }

      }
    }
  }
  private async _gettingGroupID(AccessGroupItems) {
    let AG;
    for (let a = 0; a < AccessGroupItems.length; a++) {
      AG = AccessGroupItems[a];
      //alert(AG);
      const accessGroupID: any = await this._Service.getAccessGroupID(this.props.siteUrl, this.props.accessGroupDetailsListName, AG)
      //const accessGroupID: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.accessGroupDetailsListName).items.filter("Title eq '" + AG + "'").get();
      let AccessGroupID;
      if (accessGroupID.length > 0) {
        console.log(accessGroupID);
        AccessGroupID = accessGroupID[0].GroupID;
        console.log("AccessGroupID", AccessGroupID);
        this.GetGroupMembersForProject(this.props.context, AccessGroupID, AG);
      }
    }
  }
  public async GetGroupMembers(context: WebPartContext, groupId: string): Promise<any[]> {
    let users: string[] = [];
    try {
      let response = await this._Service.getGroupMembers(groupId);
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
    if (users.length > 0) {
      this._checkingCurrent(users);
    }
    return users;

  }
  public async GetGroupMembersForProject(context: WebPartContext, groupId: string, AG: string): Promise<any[]> {
    let users: string[] = [];
    try {
      let response = await this._Service.getGroupMembers(groupId);
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
    if (users.length > 0) {
      this._checkingCurrentForProject(users, AG);
    }
    return users;
  }
  private _checkingCurrent(userEmail) {
    for (var k in userEmail) {
      if (this.currentUserEmail == userEmail[k].mail) {
        if (this.statusForRemainder == "Yes") {
          // alert("SendRemailder" + this.statusForRemainder);
          this.setState({ divForSendRemainder: false, divForCancel: true, divForDelegation: true });
        }
        if (this.statusForCancel == "Yes") {
          // alert("statusForCancel" + this.statusForRemainder);
          this.setState({ divForCancel: false, });
        }
        if (this.statusForDelegate == "Yes") {
          // alert("statusForDelegate" + this.statusForRemainder);
          this.setState({ divForDelegation: false, });
        }

      }
    }
  }
  private _checkingCurrentForProject(userEmail, AG) {
    console.log(AG, userEmail);
    for (var k in userEmail) {
      if (this.currentUserEmail == userEmail[k].mail) {
        if (AG == "Document Controller") {
          this.setState({ divForSendRemainder: false, divForCancel: true, divForDelegation: true });
        }
        if (AG == "Project Admin") {
          this.setState({ divForCancel: false, });
        }
        if (AG == "Project Admin") {
          this.setState({ divForDelegation: false, });
        }

      }
    }
  }
  private _LAUrlGetting = async () => {
    const laUrl = await this._Service.getQDMS_DocumentPermission_UnderApproval(this.props.siteUrl, this.props.requestLaURL)
    //const laUrl = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestLaURL).items.filter("Title eq 'QDMS_DocumentPermission_UnderApproval'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
  }
  private async timelineLoad() {
    let sorted_State: any[];
    const logItems = await this._Service.getLogItems(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID)
    //const logItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("Title,Status,Modified,Created,Author/ID,Author/Title,Editor/ID,Editor/Title,LogDate,WorkflowID,Revision,DocumentIndex/ID,DocumentIndex/Title,DueDate,Workflow,ID").expand("Author,Editor,DocumentIndex").filter("DocumentIndex eq '" + this.documentIndexID + "'").getAll(5000);
    if (logItems.length > 0) {
      sorted_State = _.orderBy(logItems, 'ID', ['desc']);
      console.log("sorted_State", sorted_State);
      console.log("logItems", logItems);
      this.setState({
        timelineDisplay: "",
        logItems: sorted_State,
        // logItems: logItems,
        documentName: logItems[0].Title,
        dueDate: logItems[0].DueDate,
      });
      const documentIndexItems = await this._Service.getIndexItemsWithOwnerApprover(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID);
      //const documentIndexItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.select("Owner/Title,Owner/ID,Owner/EMail,DocumentName,SourceDocumentID,CriticalDocument,Revision,DocumentID,Approver/Title,Approver/ID").expand("Owner,Approver").filter("ID eq '" + this.documentIndexID + "'").get();
      console.log("Document Index items", documentIndexItems);
      this.sourceDocumentID = documentIndexItems[0].SourceDocumentID;
      this.setState({
        documentName: documentIndexItems[0].DocumentName,
        criticalDocument: documentIndexItems[0].CriticalDocument,
        owner: documentIndexItems[0].Owner.Title,
        ownerEmail: documentIndexItems[0].Owner.EMail,
        ownerID: documentIndexItems[0].Owner.ID,
        documentID: documentIndexItems[0].DocumentID,
        revision: documentIndexItems[0].Revision,
        approverName: documentIndexItems[0].Approver.Title,
        approver: documentIndexItems[0].Approver.ID,
        documentIndexItems: documentIndexItems
      });
    }
    else {
      this.setState({ timelineDisplay: "none", statusMessage: { isShowMessage: true, message: "No History Available", messageType: 1 } });
    }
  }
  protected async triggerDocumentReview(sourceDocumentID) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
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
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  protected async triggerDocumentReviewWithWorkFlowStatus(sourceDocumentID, ResponseStatus) {
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


  }
  private _LAUrlGettingForUnderReview = async () => {
    const laUrl = await this._Service.getQDMS_DocumentPermission_UnderReview(this.props.siteUrl, this.props.requestLaURL)
    //const laUrl = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestLaURL).items.filter("Title eq 'QDMS_DocumentPermission_UnderReview'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrlForUnderReview = laUrl[0].PostUrl;
  }
  private _LAUrlGettingForDelegate = async () => {
    const laUrl = await this._Service.getQDMS_DocumentPermission_Delegate(this.props.siteUrl, this.props.requestLaURL)
    //const laUrl = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.requestLaURL).items.filter("Title eq 'QDMS_DocumentPermission_Delegate'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrlForDelegate = laUrl[0].PostUrl;
  }
  protected async triggerDocumentDelegate(sourceDocumentID, delegateToEmail, delegateForEmail, link, type, headerId, indexID) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
    console.log(siteUrl);
    const postURL = this.postUrlForDelegate;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'DelegateToEmail': delegateToEmail,
      'DelegateForEmail': delegateForEmail,
      'Link': link,
      'Workflow': type,
      'HeaderID': headerId,
      'IndexID': indexID
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
      this.sendEmailForCurrentTaskOwner(delegateForEmail, "TaskDelegationNotifyOwner", this.state.taskOwnerName);
      this.SendAnEmailForDelagation(delegateToEmail, type, this.state.delegateToTitle, link).then(msgdisplay => {
        setTimeout(() => {
          this.setState({ delegatingLoader: "none", statusMessage: { isShowMessage: true, message: "Task Delegated to  " + this.state.delegateToTitle, messageType: 4 }, });
        }, 5000);
      }).then(modalClose => {
        // this.setState({
        //   delagatePeoplePicker: "none",
        // });
      });
    }
    else {
    }
  }
  protected async triggerDocumentUnderReview(sourceDocumentID) {
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
      'WorkflowStatus': "Under Review"
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //WORKFLOW STARTED
  private workFlowStartedVoid = (item: any) => {

    return (

      <TimelineItem key="001"
        dateText={moment(item.Created).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.workflowStartedDateColor }}
        dateInnerStyle={{ background: this.props.workflowStartedDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.workflowStartedContentColor,
          padding: '20px',
          borderRadius: '8px',
          // boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              {/* Requested By&nbsp;&nbsp;   : &nbsp;&nbsp;&nbsp; {this.state.requestor} */}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Requested Date &nbsp;&nbsp; : &nbsp;&nbsp;&nbsp; {moment(item.Created).format("DD/MM/YYYY hh:mm a")}
            </div>

            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Due  Date&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{moment(item.DueDate).format("DD/MM/YYYY")}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              {/* Approver&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{ this.state.approverName} */}
            </div>
            <br></br>
            <PrimaryButton text="Details" onClick={() => this.voidIntitiatedModel(item.ID)}></PrimaryButton>
          </p>
        </p>
      </TimelineItem>
    );
  }
  private voidIntitiatedModel = async (ID) => {
    const workflowId = await this._Service.getDocumentRevisionLog(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID, ID)
    //const workflowId = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("WorkflowID").filter("DocumentIndex eq '" + this.documentIndexID + "'and (ID eq '" + ID + "')").get();
    // alert(workflowId[0].WorkflowID)  ;

    this._Service.getWorkflowHeaderWithApproverRequester(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).select("Requester/ID,Requester/Title,Approver/ID,Approver/Title,RequestedDate,DueDate").expand("Approver,Requester").get()
      .then(headerItemsFromList => {
        console.log("headerItemsFromList", headerItemsFromList);
        this.setState({
          workflowInitiatedVoid: "",
          showworkflowInitiatedVoidModal: true,
          requestor: headerItemsFromList.Requester.Title,
          approverName: headerItemsFromList.Approver.Title,
          requestorDate: moment(headerItemsFromList.RequestedDate).format('DD/MM/YYYY, h:mm a'),

        });
      });
  }
  private documentCreated1 = async (ID) => {
    const workflowId = await this._Service.getDocumentRevisionLog(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID, ID)
    //const workflowId = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("WorkflowID").filter("DocumentIndex eq '" + this.documentIndexID + "'and (ID eq '" + ID + "')").get();
    // alert(workflowId[0].WorkflowID)  ;
    var headerItems1 = "Reviewers/ID,Reviewers/Title";
    await this._Service.getItemById(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
      //await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).get()
      .then(ReviewerIDFromList => {
        this.setState({
          lengthOfReviwers: ReviewerIDFromList.ReviewersId
        });

        this._Service.getWorkflowHeaderItem(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
          //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).select("Requester/ID,Requester/Title,Approver/ID,Approver/Title,Reviewers/ID,Reviewers/Title,RequestedDate,DueDate").expand("Approver,Requester,Reviewers").get()
          .then(headerItemsFromList => {
            console.log("headerItemsFromList", headerItemsFromList);
            this.setState({
              workflowInitiated: "",
              showworkflowInitiatedModal: true,
              requestor: headerItemsFromList.Requester.Title,
              approverName: headerItemsFromList.Approver.Title,
              requestorDate: moment(headerItemsFromList.RequestedDate).format('DD/MM/YYYY, h:mm a'),
              //  reviewers: headerItemsFromList.Reviewers,
              reviewers: ReviewerIDFromList.ReviewersId == null ? [] : headerItemsFromList.Reviewers,

            });
          });
      });
  }
  private reviewedDetailsItemsDCC = async (ID) => {
    const workflowId = await this._Service.getFlowDataInDocumentRevisionLog(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID, ID)
    //const workflowId = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("WorkflowID,ID,DueDate").filter("DocumentIndex eq '" + this.documentIndexID + "' and (ID eq '" + ID + "')").get();
    // alert(workflowId[0].WorkflowID);    
    this.headerId = workflowId[0].WorkflowID;
    this.documentRevisionLogID = workflowId[0].ID;
    this.dueDateForModel = workflowId[0].DueDate;
    const workflowDetailsItems = await this._Service.getDetailsWorkflow_DCCReview(this.props.siteUrl, this.props.workflowDetailsListName, workflowId[0].WorkflowID)
    //const workflowDetailsItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Workflow").expand("Responsible,Editor").filter("HeaderID eq '" + workflowId[0].WorkflowID + "' and (Workflow eq 'DCC Review') ").get();
    console.log("Workflow detail items of header id", workflowDetailsItems);
    this.setState({
      reviewed: "",
      showReviewModal: true,
      workflowDetailItems: workflowDetailsItems,
      // dueDate:workflowDetailsItems['DueDate'],
      timelineElement: "Review",
      workflowStatus: workflowDetailsItems['ResponseStatus'],
    });
    for (var k in workflowDetailsItems) {
      if (workflowDetailsItems[k].ResponseStatus == "Cancelled") {
        this.setState({
          cancelledBy: workflowDetailsItems[k].Editor.Title,
        });
      }
    }
    this._Service.getWorkflowHeaderWithApproverRequester(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,RequestedDate,DueDate").expand("Approver,Requester").get()
      .then(headerItemsFromList => {
        console.log("headerItemsFromList", headerItemsFromList);
        this.setState({
          requestor: headerItemsFromList.Requester.Title,
          requestorEmail: headerItemsFromList.Requester.EMail,
          approverEmail: headerItemsFromList.Approver.EMail,
          approverName: headerItemsFromList.Approver.Title,
          requestorDate: moment(headerItemsFromList.RequestedDate).format('DD/MM/YYYY, h:mm a'),
          approver: headerItemsFromList.Approver.ID,
          dueDate: moment(headerItemsFromList.DueDate).format('DD/MM/YYYY, h:mm a')
        });
      });
  }
  private reviewedDetailsItems = async (ID) => {
    const workflowId = await this._Service.getFlowDataInDocumentRevisionLog(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID, ID)
    //const workflowId = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("WorkflowID,ID,DueDate").filter("DocumentIndex eq '" + this.documentIndexID + "' and (ID eq '" + ID + "')").get();
    // alert(workflowId[0].WorkflowID);    
    this.headerId = workflowId[0].WorkflowID;
    this.documentRevisionLogID = workflowId[0].ID;
    this.dueDateForModel = workflowId[0].DueDate;
    const workflowDetailsItems = await this._Service.getDetailsWorkflow_Review(this.props.siteUrl, this.props.workflowDetailsListName, workflowId[0].WorkflowID)
    //const workflowDetailsItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title").expand("Responsible,Editor").filter("HeaderID eq '" + workflowId[0].WorkflowID + "' and (Workflow eq 'Review') ").get();
    console.log("Workflow detail items of header id", workflowDetailsItems);
    this.setState({
      reviewed: "",
      showReviewModal: true,
      workflowDetailItems: workflowDetailsItems,
      // dueDate:workflowDetailsItems['DueDate'],
      timelineElement: "Review",
      workflowStatus: workflowDetailsItems['ResponseStatus'],
    });
    for (var k in workflowDetailsItems) {
      if (workflowDetailsItems[k].ResponseStatus == "Cancelled") {
        this.setState({
          cancelledBy: workflowDetailsItems[k].Editor.Title,
        });
      }
    }
    var headerItems = "Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,RequestedDate,DueDate";
    this._Service.getWorkflowHeaderWithApproverRequester(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).select(headerItems).expand("Approver,Requester").get()
      .then(headerItemsFromList => {
        console.log("headerItemsFromList", headerItemsFromList);
        this.setState({
          requestor: headerItemsFromList.Requester.Title,
          requestorEmail: headerItemsFromList.Requester.EMail,
          approverEmail: headerItemsFromList.Approver.EMail,
          approverName: headerItemsFromList.Approver.Title,
          requestorDate: moment(headerItemsFromList.RequestedDate).format('DD/MM/YYYY, h:mm a'),
          approver: headerItemsFromList.Approver.ID,
          dueDate: moment(headerItemsFromList.DueDate).format('DD/MM/YYYY, h:mm a')
        });
      });
  }
  private underApprovalVoid = async (ID) => {
    const workflowId = await this._Service.getFlowDataInDocumentRevisionLog(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID, ID)
    //const workflowId = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("WorkflowID,ID,DueDate").filter("DocumentIndex eq '" + this.documentIndexID + "' and (ID eq '" + ID + "')").get();
    // alert(workflowId[0].WorkflowID);    
    this.headerId = workflowId[0].WorkflowID;
    this.dueDateForModel = workflowId[0].DueDate;
    this.documentRevisionLogID = workflowId[0].ID;
    const workflowDetailsItems = await this._Service.getDetailsWorkflow_Void(this.props.siteUrl, this.props.workflowDetailsListName, workflowId[0].WorkflowID)
    //const workflowDetailsItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Workflow").expand("Responsible,Editor").filter("HeaderID eq '" + workflowId[0].WorkflowID + "' and (Workflow eq 'Void') ").get();
    console.log("Workflow detail items of header id", workflowDetailsItems);
    this.setState({
      reviewed: "",
      showReviewModal: true,
      workflowDetailItems: workflowDetailsItems,
      // dueDate:workflowDetailsItems['DueDate'],
      timelineElement: "Approval",
      workflowStatus: workflowDetailsItems['ResponseStatus'],
    });
    for (var k in workflowDetailsItems) {
      if (workflowDetailsItems[k].ResponseStatus == "Cancelled") {
        this.setState({
          cancelledBy: workflowDetailsItems[k].Editor.Title,
        });
      }
    }
    this._Service.getWorkflowHeaderWithApproverRequester(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,RequestedDate,DueDate").expand("Approver,Requester").get()
      .then(headerItemsFromList => {
        console.log("headerItemsFromList", headerItemsFromList);
        this.setState({
          requestor: headerItemsFromList.Requester.Title,
          requestorEmail: headerItemsFromList.Requester.EMail,
          approverEmail: headerItemsFromList.Approver.EMail,
          approverName: headerItemsFromList.Approver.Title,
          requestorDate: moment(headerItemsFromList.RequestedDate).format('DD/MM/YYYY, h:mm a'),
          approver: headerItemsFromList.Approver.ID,
          dueDate: moment(headerItemsFromList.DueDate).format('DD/MM/YYYY, h:mm a'),
        });
      });

  }
  private approvedDetailsItems = async (ID) => {
    const workflowId = await this._Service.getFlowDataInDocumentRevisionLog(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentIndexID, ID)
    //const workflowId = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.select("WorkflowID,ID,DueDate").filter("DocumentIndex eq '" + this.documentIndexID + "' and (ID eq '" + ID + "')").get();
    // alert(workflowId[0].WorkflowID)  ;
    this.headerId = workflowId[0].WorkflowID;
    this.documentRevisionLogID = workflowId[0].ID;
    this.dueDateForModel = workflowId[0].DueDate;
    const workflowDetailsItems = await this._Service.getWorkflowApproval(this.props.siteUrl, this.props.workflowDetailsListName, workflowId[0].WorkflowID)
    //const workflowDetailsItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.select("Responsible/Title,Responsible/ID,Responsible/EMail,ResponseDate,ResponseStatus,ResponsibleComment,DueDate,ID,TaskID,Editor/Title,Link,Workflow").expand("Responsible,Editor").filter("HeaderID eq '" + workflowId[0].WorkflowID + "' and (Workflow eq 'Approval') ").get();
    console.log("Workflow detail items of header id", workflowDetailsItems);
    this.setState({
      reviewed: "",
      showReviewModal: true,
      workflowDetailItems: workflowDetailsItems,
      //dueDate: workflowId[0].DueDate,
      timelineElement: "Approval",
    });
    for (var k in workflowDetailsItems) {
      if (workflowDetailsItems[k].ResponseStatus == "Cancelled") {
        this.setState({
          cancelledBy: workflowDetailsItems[k].Editor.Title,
        });
      }
    }
    this._Service.getWorkflowHeaderWithApproverRequester(this.props.siteUrl, this.props.workflowHeaderListName, workflowId[0].WorkflowID)
      //sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(workflowId[0].WorkflowID).select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,Approver/EMail,RequestedDate,DueDate").expand("Approver,Requester").get()
      .then(headerItemsFromList => {
        console.log("headerItemsFromList", headerItemsFromList);
        this.setState({
          requestor: headerItemsFromList.Requester.Title,
          requestorEmail: headerItemsFromList.Requester.EMail,
          approverEmail: headerItemsFromList.Approver.EMail,
          approverName: headerItemsFromList.Approver.Title,
          requestorDate: moment(headerItemsFromList.RequestedDate).format('DD/MM/YYYY, h:mm a'),
          approver: headerItemsFromList.Approver.ID,
          dueDate: moment(headerItemsFromList.DueDate).format('DD/MM/YYYY, h:mm a')
        });
      });
  }
  private _closeModal = (): void => {
    this.setState({ iframeModalclose: false, showModal: false, showReviewModal: false, delagatePeoplePicker: "none", showworkflowInitiatedModal: false, showworkflowInitiatedVoidModal: false, workflowInitiatedVoid: "none" });
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
  private async currentUser() {
    this._Service.getCurrentUser()
      //sp.web.currentUser.get()
      .then(currentUser => {
        this.setState({
          currentUser: currentUser.Id,
          currentUserEmail: currentUser.Email,
          currentUserName: currentUser.Title,
        });
        console.log(this.state.currentUser);
      });

  }
  private documentExpiryRevoked = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentApprovalDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentApprovalContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Revoke Date&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{moment(item.Created).format("DD/MM/YYYY hh:mm a")}
            </div>
            <br></br>
          </p>
        </p>
      </TimelineItem>
    );
  }
  private directPublish = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentApprovalDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentApprovalContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: (item.Status != "Published") ? "none" : "" }}>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Published By&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp; {item.Editor.Title}
              </div>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Published Date&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp; {moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
              </div>
            </div>
            <br></br>
          </p>
        </p>
      </TimelineItem>
    );
  }
  private documentExpired = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentApprovalDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentApprovalContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Expired Date &nbsp;&nbsp; :&nbsp;&nbsp;&nbsp; {moment(item.Created).format("DD/MM/YYYY hh:mm a")}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>

            </div>
            <br></br>

          </p>
        </p>
      </TimelineItem>
    );
  }
  private Cancelled = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentApprovalDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentApprovalContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: '#61b8ff' }}>{item.Workflow + " " + item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Cancelled By  &nbsp;&nbsp;: &nbsp;&nbsp;&nbsp; {item.Editor.Title}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Cancelled Date&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;{moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
            </div>
            <br></br>
          </p>
        </p>
      </TimelineItem>
    );
  }
  private dccReview = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentApprovalDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentApprovalContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: '#61b8ff' }}>{item.Status + " " + "DCC Review"}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>
          <div style={{ color: '#61b8ff', display: (item.Status != "Under Review") ? "" : "none" }}>
            <p style={{ fontSize: '12px' }}>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Reviewed By &nbsp;&nbsp; : &nbsp;&nbsp;&nbsp; {item.Editor.Title}
              </div>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Reviewed Date&nbsp;&nbsp;: &nbsp;&nbsp;&nbsp;  {moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
              </div>
              <br></br>
            </p>
          </div>
          <PrimaryButton text="Details" onClick={() => this.reviewedDetailsItemsDCC(item.ID)}></PrimaryButton>

        </p>
      </TimelineItem>
    );
  }
  private documentVoided = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentVoidDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentVoidContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              {/* Approver &nbsp;&nbsp;  : &nbsp;&nbsp;&nbsp; {item.Editor.Title} */}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              {/* Approved Date &nbsp;&nbsp;: &nbsp;&nbsp;&nbsp; {moment(item.Modified).format("DD/MM/YYYY hh:mm a")} */}
            </div>
            <br></br>
            <PrimaryButton text="Details" onClick={() => this.underApprovalVoid(item.ID)}></PrimaryButton>
          </p>
        </p>
      </TimelineItem>
    );
  }
  //PUBLISHED
  private approved = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentApprovalDateColor }}
        dateInnerStyle={{ background: this.props.documentApprovalDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentApprovalContentColor,
          padding: '20px',
          borderRadius: '8px',
          boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: (item.Status != "Published") ? "none" : "" }}>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Published By&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp; {item.Editor.Title}
              </div>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Published Date&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp; {moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
              </div>
            </div>
            <br></br>
            <PrimaryButton text="Details" onClick={() => this.approvedDetailsItems(item.ID)}></PrimaryButton>
          </p>

        </p>
      </TimelineItem>
    );
  }
  //REVIEWED
  private reviewed = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentReviewedDateColor }}
        dateInnerStyle={{ background: this.props.documentReviewedDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentReviewedContentColor,
          padding: '20px',
          borderRadius: '8px',
          // boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: (item.Status != "Reviewed") ? "none" : "" }}>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Reviewed By &nbsp;&nbsp; :&nbsp;&nbsp;&nbsp; {item.Editor.Title}
              </div>
              <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
                Reviewed Date &nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{moment(item.Modified).format("DD/MM/YYYY hh:mm a")}
              </div>
              <br></br>
            </div>
            <PrimaryButton text="Details" onClick={() => this.reviewedDetailsItems(item.ID)}></PrimaryButton>
          </p>
        </p>
      </TimelineItem>
    );
  }
  //WORKFLOW STARTED
  private workFlowStarted = (item: any) => {

    return (

      <TimelineItem key="001"
        dateText={moment(item.Created).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.workflowStartedDateColor }}
        dateInnerStyle={{ background: this.props.workflowStartedDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.workflowStartedContentColor,
          padding: '20px',
          borderRadius: '8px',
          // boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff' }}>{this.state.documentName}</h4>
        <p>

          <p style={{ fontSize: '12px' }}>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              {/* Requested By&nbsp;&nbsp;   : &nbsp;&nbsp;&nbsp; {this.state.requestor} */}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Requested Date &nbsp;&nbsp; : &nbsp;&nbsp;&nbsp; {moment(item.Created).format("DD/MM/YYYY hh:mm a")}
            </div>

            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              Due  Date&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{moment(item.DueDate).format("DD/MM/YYYY")}
            </div>
            <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
              {/* Approver&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{ this.state.approverName} */}
            </div>
            <br></br>
            <PrimaryButton text="Details" onClick={() => this.documentCreated1(item.ID)}></PrimaryButton>
          </p>
        </p>
      </TimelineItem>
    );
  }
  //DOCUMENT CREATED
  private documentCreated = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Created).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentCreatedDateColor }}
        dateInnerStyle={{ background: this.props.documentCreatedDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentCreatedContentColor,
          padding: '20px',
          borderRadius: '8px',
          // boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff', wordBreak: "break-all" }}>{this.state.documentName}</h4>
        <p> <p style={{ fontSize: '12px' }}>
          <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
            Created By&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{item.Author.Title}
          </div>
          <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
            Created Date &nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;{moment(item.Created).format("DD/MM/YYYY hh:mm a")}
          </div>

          <br></br>

        </p>
        </p>
      </TimelineItem>
    );
  }
  //INTERNALLY TRANSMITTEDtO CUSTRUCTION/fABRICATION
  private internalTransittalConFab = (item: any) => {
    return (
      <TimelineItem key="001"
        dateText={moment(item.Created).format("DD/MM/YYYY hh:mm a")}
        style={{ color: this.props.documentCreatedDateColor }}
        dateInnerStyle={{ background: this.props.documentCreatedDateColor, color: '#000' }}
        bodyContainerStyle={{
          background: this.props.documentCreatedContentColor,
          padding: '20px',
          borderRadius: '8px',
          // boxShadow: '0.5rem 0.5rem 2rem 0 rgba(0, 0, 0, 0.2)',
        }}>
        <h3 style={{ color: this.props.statusColor }}>{item.Status}</h3>
        <h4 style={{ color: '#61b8ff', wordBreak: "break-all" }}>{this.state.documentName}</h4>
        <p> <p style={{ fontSize: '12px' }}>
          <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
            Transmitted By&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp;{item.Author.Title}
          </div>
          <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
            Transmitted Date &nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;{moment(item.Created).format("DD/MM/YYYY hh:mm a")}
          </div>

          <br></br>

        </p>
        </p>
      </TimelineItem>
    );
  }
  //cancel task for under approvel and under review
  public cancelTask = async (item: any, key: any) => {
    let ReviewedCount = 0;
    let cancelCount = 0;
    let UnderReview;
    let ReturnedWithComments;
    var today = new Date();
    let date = today.toLocaleString();
    if (this.state.timelineElement == "Review") {
      this.loaderforcancel = "";
      const resdata = {
        ResponseStatus: "Cancelled",
      }
      this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, item.ID, resdata)
        /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(item.ID).update({
          ResponseStatus: "Cancelled",
        }) */
        .then(async taskDelete => {
          if (item.TaskID != null) {
            this._Service.deleteItemById(this.props.siteUrl, this.props.workflowTaskListName, item.TaskID)
            /* let list = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName);
            await list.items.getById(item.TaskID).delete(); */
          }
          const reviewersResponseStatus = await this._Service.getReviewersResponseStatus(this.props.siteUrl, this.props.workflowDetailsListName, this.headerId);
          //const reviewersResponseStatus = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.select("ResponseStatus").filter("HeaderID eq " + this.headerId + " and (Workflow eq 'Review')").get();
          console.log(reviewersResponseStatus.length);
          for (var k in reviewersResponseStatus) {
            if (reviewersResponseStatus[k].ResponseStatus == "Reviewed") { ReviewedCount++; }
            else if (reviewersResponseStatus[k].ResponseStatus == "Under Review") { UnderReview = "Yes"; }
            else if (reviewersResponseStatus[k].ResponseStatus == "Returned with comments") { ReturnedWithComments = "Yes"; }
            else if (reviewersResponseStatus[k].ResponseStatus == "Cancelled") { cancelCount++; }
          }
          //all are reviewed 
          if (reviewersResponseStatus.length == add(ReviewedCount, cancelCount)) {
            console.log(add(ReviewedCount, cancelCount));
            const headeritem = {
              WorkflowStatus: "Under Approval",//headerlist
              Workflow: "Approval",
              ReviewedDate: date,
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headeritem)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
              ({
                WorkflowStatus: "Under Approval",//headerlist
                Workflow: "Approval",
                ReviewedDate: date,
              }); */
            const inditem = {
              WorkflowStatus: "Under Approval",//docIndex
              Workflow: "Approval"
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID, inditem)
            /*  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).update
               ({
                 WorkflowStatus: "Under Approval",//docIndex
                 Workflow: "Approval"
               }); */
            //Updationg DocumentRevisionlog 
            const logitem = {
              Status: "Cancelled",
              LogDate: date,
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentRevisionLogID, logitem)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.getById(this.documentRevisionLogID).update({
              Status: "Cancelled",
              LogDate: date,
            }); */
            const logdata = {
              Status: "Under Approval",
              LogDate: date,
              WorkflowID: this.headerId,
              DocumentIndexId: this.documentIndexID,
              DueDate: this.dueDateForModel,
              Workflow: "Approval",
              Revision: this.state.revision,
              Title: this.state.documentID,
            }
            this._Service.createNewItem(this.props.siteUrl, this.props.DocumentRevisionLog, logdata)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.add({
              Status: "Under Approval",
              LogDate: date,
              WorkflowID: this.headerId,
              DocumentIndexId: this.documentIndexID,
              DueDate: this.dueDateForModel,
              Workflow: "Approval",
              Revision: this.state.revision,
              Title: this.state.documentID,
            }); */
            //upadting source library without version change.            
            let bodyArray = [
              { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
            ];
            this._Service.validateUpdateListItem(this.props.siteUrl, this.props.sourceDocuments, this.sourceDocumentID, bodyArray)
            /* sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocuments).items.getById(this.sourceDocumentID).validateUpdateListItem
              (
                bodyArray,
              ); */
            this._Service.getUserIdByEmail(this.state.approverEmail)
              //sp.web.siteUsers.getByEmail(this.state.approverEmail).get()
              .then(async user => {
                console.log('User Id: ', user.Id);
                this.setState({
                  hubSiteUserId: user.Id,
                });
                //Task delegation 
                const taskDelegation: any[] = await this._Service.getTaskDelegationData(this.props.siteUrl, this.props.taskDelegationListName, user.Id)
                //const taskDelegation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationListName).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
                console.log(taskDelegation);
                if (taskDelegation.length > 0) {
                  let duedate = moment(this.state.dueDate).toDate();
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
                  }//duedate checking

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
                          const wfitem = {
                            HeaderIDId: Number(this.headerId),
                            Workflow: "Approval",
                            Title: this.state.documentName,
                            ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                            DueDate: this.dueDateForModel,
                            DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                            ResponseStatus: "Under Approval",
                            OwnerId: this.state.ownerID,
                            SourceDocument: {
                              "__metadata": { type: "SP.FieldUrlValue" },
                              Description: this.state.documentName,
                              Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                            },

                          }
                          this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsListName, wfitem)
                            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.add
                              ({
                                HeaderIDId: Number(this.headerId),
                                Workflow: "Approval",
                                Title: this.state.documentName,
                                ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                                DueDate: this.dueDateForModel,
                                DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                                ResponseStatus: "Under Approval",
                                OwnerId: this.state.ownerID,
                                SourceDocument: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName,
                                  Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                },
      
                              }) */
                            .then(async r => {
                              this.setState({ detailIdForApprover: r.data.ID });
                              this.newDetailItemID = r.data.ID;
                              const wfdetail = {
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: "Link to Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },
                              }
                              this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, r.data.ID, wfdetail)
                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(r.data.ID).update({
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: "Link to Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },
                              }); */
                              const wfheader = {                   //headerlist
                                ApproverId: this.state.delegateToIdInSubSite
                              }
                              this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, wfheader)
                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                                ({                   //headerlist
                                  ApproverId: this.state.delegateToIdInSubSite
                                }); */
                              const inditem = {
                                ApproverId: this.state.delegateToIdInSubSite,
                              }
                              this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID, inditem)
                              /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).update
                                ({
                                  ApproverId: this.state.delegateToIdInSubSite
                                }); */
                              //upadting source library without version change.
                              const sourceitem = {
                                ApproverId: this.state.delegateToIdInSubSite,
                              }
                              this._Service.updateItemById(this.props.siteUrl, "/SourceDocuments", this.sourceDocumentID, sourceitem)
                              /* await sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).update({
                                ApproverId: this.state.delegateToIdInSubSite,

                              }); */
                              //MY tasks list updation
                              const taskitem = {
                                Title: "Approve '" + this.state.documentName + "'",
                                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                DueDate: this.dueDateForModel,
                                StartDate: this.currentDate,
                                AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                                Workflow: "Approval",
                                // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                                DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                                Source: (this.props.project ? "Project" : "QDMS"),
                                DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: "Link to Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },

                              }
                              this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                                /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                                  ({
                                    Title: "Approve '" + this.state.documentName + "'",
                                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                    DueDate: this.dueDateForModel,
                                    StartDate: this.currentDate,
                                    AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                                    Workflow: "Approval",
                                    // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                                    DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                                    Source: (this.props.project ? "Project" : "QDMS"),
                                    DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: "Link to Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    },
  
                                  }) */
                                .then(taskId => {
                                  const detailitem = {
                                    TaskID: taskId.data.ID,
                                  }
                                  this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, r.data.ID, detailitem)
                                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(r.data.ID).update
                                      ({
                                        TaskID: taskId.data.ID,
                                      }) */
                                    .then(afterTask => {
                                      this.triggerDocumentReviewWithWorkFlowStatus(this.sourceDocumentID, "Under Approval");
                                    }).then(msgBar => {
                                      this.SendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                    }).then(mail => {
                                      //modalclosing
                                      this.loaderforcancel = "none";
                                      this.setState({ delegatingLoader: "none", statusMessage: { isShowMessage: true, message: "Review Cancelled for" + this.state.approverName, messageType: 4 }, });
                                      setTimeout(() => {
                                        this.setState({
                                          showReviewModal: false,
                                        });
                                      }, 3000);
                                    });
                                });//taskID
                            });//r

                        });//DelegatedFor
                    });//DelegatedTo
                }

                else {
                  const detitem = {
                    HeaderIDId: Number(this.headerId),
                    Workflow: "Approval",
                    Title: this.state.documentName,
                    ResponsibleId: this.state.approver,
                    DueDate: this.dueDateForModel,
                    ResponseStatus: "Under Approval",
                    OwnerId: this.state.ownerID,
                    SourceDocument: {
                      "__metadata": { type: "SP.FieldUrlValue" },
                      Description: this.state.documentName,
                      Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                  }
                  this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsListName, detitem)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.add
                      ({
                        HeaderIDId: Number(this.headerId),
                        Workflow: "Approval",
                        Title: this.state.documentName,
                        ResponsibleId: this.state.approver,
                        DueDate: this.dueDateForModel,
                        ResponseStatus: "Under Approval",
                        OwnerId: this.state.ownerID,
                        SourceDocument: {
                          "__metadata": { type: "SP.FieldUrlValue" },
                          Description: this.state.documentName,
                          Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        },
                      }) */
                    .then(async r => {
                      this.setState({ detailIdForApprover: r.data.ID });
                      this.newDetailItemID = r.data.ID;
                      const detailitem = {
                        Link: {
                          "__metadata": { type: "SP.FieldUrlValue" },
                          Description: "Link to Approve",
                          Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        },
                      }
                      this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, r.data.ID, detailitem)
                      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(r.data.ID).update({
                        Link: {
                          "__metadata": { type: "SP.FieldUrlValue" },
                          Description: "Link to Approve",
                          Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        },
                      }); */

                      //MY tasks list updation
                      const taskitem = {
                        Title: "Approve '" + this.state.documentName + "'",
                        Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                        DueDate: this.dueDateForModel,
                        StartDate: this.currentDate,
                        AssignedToId: user.Id,
                        Workflow: "Approval",
                        Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                        Source: (this.props.project ? "Project" : "QDMS"),
                        Link: {
                          "__metadata": { type: "SP.FieldUrlValue" },
                          Description: "Link to Approve",
                          Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                        },

                      }
                      this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                        /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                          ({
                            Title: "Approve '" + this.state.documentName + "'",
                            Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                            DueDate: this.dueDateForModel,
                            StartDate: this.currentDate,
                            AssignedToId: user.Id,
                            Workflow: "Approval",
                            Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                            Source: (this.props.project ? "Project" : "QDMS"),
                            Link: {
                              "__metadata": { type: "SP.FieldUrlValue" },
                              Description: "Link to Approve",
                              Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                            },
  
                          }) */
                        .then(taskId => {
                          const detailitem = {
                            TaskID: taskId.data.ID,
                          }
                          this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, r.data.ID, detailitem)
                          /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(r.data.ID).update
                            ({
                              TaskID: taskId.data.ID,
                            }); */
                          this.triggerDocumentReviewWithWorkFlowStatus(this.sourceDocumentID, "Under Approval");
                          //notification preference checking                                 
                          this.SendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID).then(mailSuccess => {
                          }).then(modalClose => {
                            this.loaderforcancel = "none";
                            this.setState({ delegatingLoader: "none", statusMessage: { isShowMessage: true, message: "Review Cancelled for" + this.state.approverName, messageType: 4 }, });
                            setTimeout(() => {
                              this.setState({
                                showReviewModal: false,
                              });
                            }, 3000);
                          });

                        });//taskID
                    });//r
                }//else no delegation


              });
            //if end 
          }
          //if last response status is Return with comments
          else if (ReturnedWithComments == "Yes" && UnderReview !== "Yes") {
            //alert("Returned with Comments  email  to originator");
            const headitem = {
              WorkflowStatus: "Returned with comments",
            }
            this._Service.updateItemById(this.props.siteUrl, "WorkflowHeader", this.headerId, headitem)
            /* sp.web.getList(this.props.siteUrl + "/Lists/WorkflowHeader").items.getById(this.headerId).update({
              WorkflowStatus: "Returned with comments",
            }); */
            const indexitem = {
              WorkflowStatus: "Returned with comments",
            }
            this._Service.updateItemById(this.props.siteUrl, "DocumentIndex", this.documentIndexID, indexitem)
            /* sp.web.getList(this.props.siteUrl + "/Lists/DocumentIndex").items.getById(this.documentIndexID).update({
              WorkflowStatus: "Returned with comments",
            }); */
            //Updationg DocumentRevisionlog 
            const logitem = {
              Status: "Returned with comments",
              LogDate: date,
              Workflow: "Review",
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentRevisionLogID, logitem)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.getById(this.documentRevisionLogID).update({
              Status: "Returned with comments",
              LogDate: date,
              Workflow: "Review",
            }); */
            //upadting source library without version change.            
            let bodyArray = [
              { "FieldName": "WorkflowStatus", "FieldValue": "Returned with comments" }
            ];
            this._Service.validateUpdateListItem(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, bodyArray)
              /* sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem(
                bodyArray,
              ) */
              .then(afterHeaderStatusUpdate => {
                this.triggerDocumentReviewWithWorkFlowStatus(this.sourceDocumentID, "Returned with comments");
                this.returnWithComments();
              }).then(msg => {
                this.loaderforcancel = "none";
                this.setState({ delegatingLoader: "none", statusMessage: { isShowMessage: true, message: "Review Cancelled for" + item.Responsible.Title, messageType: 4 }, });
              });
          }
          else if (ReviewedCount == 0 && UnderReview != "Yes" && ReturnedWithComments != "Yes" && cancelCount == 1) {
            const wfheaditem = {
              WorkflowStatus: "Under Approval",//headerlist
              Workflow: "Approval",
              ReviewedDate: date,
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, wfheaditem)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
              ({
                WorkflowStatus: "Under Approval",//headerlist
                Workflow: "Approval",
                ReviewedDate: date,
              }); */
            const inddata = {
              WorkflowStatus: "Under Approval",//docIndex
              Workflow: "Approval"
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID, inddata)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).update
              ({
                WorkflowStatus: "Under Approval",//docIndex
                Workflow: "Approval"
              }); */
            //Updationg DocumentRevisionlog 
            const logdata = {
              Status: "Cancelled",
              LogDate: date,
            }
            this._Service.updateItemById(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentRevisionLogID, logdata)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.getById(this.documentRevisionLogID).update({
              Status: "Cancelled",
              LogDate: date,
            }); */
            const newlog = {
              Status: "Under Approval",
              LogDate: date,
              WorkflowID: this.headerId,
              DocumentIndexId: this.documentIndexID,
              DueDate: this.state.dueDate,
              Workflow: "Approval",
              Revision: this.state.revision,
              Title: this.state.documentID,
            }
            this._Service.createNewItem(this.props.siteUrl, this.props.DocumentRevisionLog, newlog)
            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.add({
              Status: "Under Approval",
              LogDate: date,
              WorkflowID: this.headerId,
              DocumentIndexId: this.documentIndexID,
              DueDate: this.state.dueDate,
              Workflow: "Approval",
              Revision: this.state.revision,
              Title: this.state.documentID,
            }); */
            //upadting source library without version change.            
            let bodyArray = [
              { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
            ];
            this._Service.validateUpdateListItem(this.props.siteUrl, this.props.sourceDocuments, this.sourceDocumentID, bodyArray)
            /* sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocuments).items.getById(this.sourceDocumentID).validateUpdateListItem
              (
                bodyArray,
              ); */
            this._Service.getUserIdByEmail(this.state.approverEmail)
              //sp.web.siteUsers.getByEmail(this.state.approverEmail).get()
              .then(async user => {
                console.log('User Id: ', user.Id);
                this.setState({
                  hubSiteUserId: user.Id,
                });
                //Task delegation 
                const taskDelegation: any[] = await this._Service.getTaskDelegationData(this.props.siteUrl, this.props.taskDelegationListName, user.Id)
                //const taskDelegation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.taskDelegationListName).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
                console.log(taskDelegation);
                if (taskDelegation.length > 0) {
                  let duedate = moment(this.state.dueDate).toDate();
                  let ToDate = moment(taskDelegation[0].ToDate).toDate();
                  let FromDate = moment(taskDelegation[0].FromDate).toDate();
                  duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                  ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
                  FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
                  if (duedate >= FromDate && duedate <= ToDate) {
                    this.setState({
                      approverEmail: taskDelegation[0].DelegatedTo.EMail,
                      approverName: taskDelegation[0].DelegatedTo.Title,
                      delegatedToId: taskDelegation[0].DelegatedTo.ID,
                      delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                    });
                  }
                }
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
                        const detaildata = {
                          HeaderIDId: Number(this.headerId),
                          Workflow: "Approval",
                          ResponseStatus: "Under Approval",
                          Title: this.state.documentName,
                          ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                          DueDate: this.state.dueDate,
                          DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                          OwnerId: this.state.ownerID,
                          SourceDocument: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: this.state.documentName,
                            Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                          },
                        }
                        this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsListName, detaildata)
                          /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.add
                            ({
                              HeaderIDId: Number(this.headerId),
                              Workflow: "Approval",
                              ResponseStatus: "Under Approval",
                              Title: this.state.documentName,
                              ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                              DueDate: this.state.dueDate,
                              DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                              OwnerId: this.state.ownerID,
                              SourceDocument: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: this.state.documentName,
                                Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                              },
                            }) */
                          .then(async r => {
                            this.setState({ detailIdForApprover: r.data.ID });
                            this.newDetailItemID = r.data.ID;
                            const detailitem = {
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: "Link to Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              },
                            }
                            this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, r.data.ID, detailitem)
                            /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(r.data.ID).update({
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: "Link to Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              },
                            }); */
                            //MY tasks list updation
                            const taskitem = {
                              Title: "Approve '" + this.state.documentName + "'",
                              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                              DueDate: this.state.dueDate,
                              StartDate: date,
                              AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : this.state.hubSiteUserId),
                              Workflow: "Approval",
                              Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                              DelegatedOn: (this.state.delegatedToId !== "" ? date : " "),
                              Source: (this.props.project ? "Project" : "QDMS"),
                              DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: "Link to Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              },

                            }
                            await this._Service.createNewItem(this.props.siteUrl, this.props.workflowTaskListName, taskitem)
                              /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.add
                                ({
                                  Title: "Approve '" + this.state.documentName + "'",
                                  Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                  DueDate: this.state.dueDate,
                                  StartDate: date,
                                  AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : this.state.hubSiteUserId),
                                  Workflow: "Approval",
                                  Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                                  DelegatedOn: (this.state.delegatedToId !== "" ? date : " "),
                                  Source: (this.props.project ? "Project" : "QDMS"),
                                  DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                                  Link: {
                                    "__metadata": { type: "SP.FieldUrlValue" },
                                    Description: "Link to Approve",
                                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                  },
  
                                }) */
                              .then(taskId => {
                                const detailitem =
                                {
                                  TaskID: taskId.data.ID,
                                }

                                this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, r.data.ID, detailitem)
                                /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(r.data.ID).update
                                  ({
                                    TaskID: taskId.data.ID,
                                  }); */
                                this.triggerDocumentReviewWithWorkFlowStatus(this.sourceDocumentID, "Under Approval");
                                //notification preference checking                                 
                                this.SendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                this._closeModal();
                              });
                          });
                      });
                  });
              });
            //if end 
          }
        });
    }
    else if (this.state.timelineElement == "Approval") {
      //alert("inside");
      this.loaderforcancel = "";
      console.log("Approval part in cancel task");
      //alert(item.ID);
      const detaildata = {
        ResponseStatus: "Cancelled",
      }
      this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, item.ID, detaildata)
        /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(item.ID).update({
          ResponseStatus: "Cancelled",
        }) */
        .then(async taskDelete => {
          if (item.TaskID != null) {
            let list = await this._Service.deleteItemById(this.props.siteUrl, this.props.workflowTaskListName, item.TaskID)
            /* let list = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName);
            await list.items.getById(item.TaskID).delete(); */
          }
        });
      const headitem = {
        WorkflowStatus: "Cancelled",
      }
      this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update({
        WorkflowStatus: "Cancelled",
      }); */
      const inddata = {
        WorkflowStatus: "Cancelled",//docIndex
        Workflow: "Approval"
      }

      this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID, inddata)
      /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).update
        ({
          WorkflowStatus: "Cancelled",//docIndex
          Workflow: "Approval"
        }); */
      //upadting source library without version change.            
      let bodyArray = [
        { "FieldName": "WorkflowStatus", "FieldValue": "Cancelled" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
      ];
      this._Service.validateUpdateListItem(this.props.siteUrl, "SourceDocuments", this.sourceDocumentID, bodyArray)
      /* sp.web.getList(this.props.siteUrl + "/SourceDocuments").items.getById(this.sourceDocumentID).validateUpdateListItem
        (
          bodyArray,
        ); */
      const canceldata = {
        Status: "Cancelled",
        LogDate: date,
      }
      this._Service.updateItemById(this.props.siteUrl, this.props.DocumentRevisionLog, this.documentRevisionLogID, canceldata)
        /*  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLog).items.getById(this.documentRevisionLogID).update({
           Status: "Cancelled",
           LogDate: date,
         }) */
        .then(triggerFlow => {
          this.triggerDocumentReviewWithWorkFlowStatus(this.sourceDocumentID, "Cancelled");
        }).then(redirect => {
          this.loaderforcancel = "none";
          this.setState({ delegatingLoader: "none", statusMessage: { isShowMessage: true, message: "WorkFlow Cancelled", messageType: 4 }, });
          setTimeout(() => {
            this.setState({
              showModal: false,
              showReviewModal: false,
            });
          }, 3000);

        }).then(cancelApproval => {
          this.emailForCancel();
        });
    }

  }
  // email for cancelling
  private emailForCancel() {
    this.SendAnEmailUsingMSGraph(this.state.requestorEmail, "DocCancel", this.state.requestor, this.newDetailItemID);
    this.SendAnEmailUsingMSGraph(this.state.ownerEmail, "DocCancel", this.state.owner, this.newDetailItemID);
  }
  private returnWithComments() {
    this.SendAnEmailUsingMSGraph(this.state.requestorEmail, "DocReturn", this.state.requestor, this.newDetailItemID);
    this.SendAnEmailUsingMSGraph(this.state.ownerEmail, "DocReturn", this.state.owner, this.newDetailItemID);
  }
  // sending Email
  private async SendAnEmailUsingMSGraph(email, type, name, detailID): Promise<void> {
    let Subject;
    let Body;
    let link;

    //console.log(queryVar);
    const notificationPreference: any[] = await this._Service.getNotificationPref(this.props.siteUrl, this.props.notificationPrefListName, email)
    //const notificationPreference: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.notificationPrefListName).items.select("Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail").expand("EmailUser").filter("EmailUser/EMail eq '" + email + "'").get();
    // console.log(notificationPreference);
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference == "Send all emails") {
        this.status = "Yes";
        //console.log("Send mail for all");                 
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.criticalDocument == true) {
        //console.log("Send mail for critical document");
        this.status = "Yes";
      }
      else {
        this.status = "No";
      }
    }
    else if (this.state.criticalDocument == true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await this._Service.getEmailNoficationSettings(this.props.siteUrl, this.props.emailNoficationSettings, type)
    //const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNoficationSettings).items.filter("Title eq '" + type + "'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;

    if (type == "DocApproval") {
      link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + detailID}>Link</a>`;

    }
    else {
      link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + detailID}>Link</a>`;

    }
    //Replacing the email body with current values
    let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
    let replacedSubjectWithDueDate = replaceString(replacedSubject, '[DueDate]', this.state.dueDate);
    let replaceRequester = replaceString(Body, '[Sir/Madam],', name + "<br></br>");
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
    let replacelink = replaceString(replaceBody, '[Link]', link);
    let var1: any[] = replacelink.split('/');
    // alert(var1[0]);
    let FinalBody = replacelink;


    //mail sending
    if (this.status == "Yes") {
      //Check if TextField value is empty or not  
      if (email) {
        //Create Body for Email  
        let emailPostBody: any = {
          "message": {
            "subject": replacedSubjectWithDueDate,
            "body": {
              "contentType": "HTML",
              "content": FinalBody
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
        this._Service.sendMail(emailPostBody);
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
  private sendRemainder = async (item: any, key: any) => {
    setTimeout(() => {
      this.setState({ statusMessage: { isShowMessage: true, message: "Remainder Send ", messageType: 4 }, });
    }, 5000);
    this.setState({
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
    });
    let Subject;
    let Body;
    let link;
    // //Email Notification Settings.
    const emailNoficationSettings: any[] = await this._Service.getRevisionHistory(this.props.siteUrl, this.props.emailNoficationSettings)
    //const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNoficationSettings).items.filter("PageName eq 'RevisionHistory'").get();
    console.log("Notifications", emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    //Replacing the email body with current values
    let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
    let finalsubject = replaceString(replacedSubject, '[Review/Approval]', "<b>" + (this.state.timelineElement == "Review") ? "Review" : "Approval" + "</b>");
    let replaceRequester = replaceString(Body, '[Sir/Madam],', item.Responsible.Title);
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
    let replacereviewApprove = replaceString(replaceBody, '[Review/Approve]', "<B>" + (this.state.timelineElement == "Review") ? "Review" : "Approval" + "</B>");
    let duedate = moment(this.dueDateForModel).format("DD/MM/YYYY");
    let replacedate = replaceString(replacereviewApprove, '[DueDate]', duedate);
    let FinalBody = replacedate;
    console.log(FinalBody);
    //Check if TextField value is empty or not  
    if (item.Responsible.EMail) {
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": finalsubject,

          "body": {
            "contentType": "HTML",
            "content": FinalBody

          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": item.Responsible.EMail
              }
            }
          ],
        }
      };
      //Send Email uisng MS Graph  
      this._Service.sendMail(emailPostBody);
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
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
    console.log(items[0].id);
    let selectedUsers: any[] = [];
    for (let item in items) {
      selectedUsers.push(items[0].id);
    }
    this.setState({
      delegatedToId: items[0].id,
      delegateToEmail: items[0].secondaryText,
      delegateToTitle: items[0].text
    });
  }
  public _delegateClick = (item: any, key: any) => {
    if (item.ResponseStatus == "Under Review" || item.ResponseStatus == "Under Approval") {
      this.setState({
        delagatePeoplePicker: "",
      });
    }

  }
  public _delegateSubmit = async (item: any, key: any) => {
    console.log(item);
    var today = new Date();
    let date = today.toLocaleString();
    let type;
    let link;
    console.log("delegateTo.ID", this.state.delegatedToId);
    console.log("Responsible.ID", item.ResponsibleId);
    console.log("Responsible.Title", item.Responsible.Title);
    console.log("WorkFlow", item.Workflow);
    this.setState({
      delagatePeoplePicker: "none",
      delegatingLoader: "",
      taskOwnerName: item.Responsible.Title,
    });
    const detialitem = {
      ResponsibleId: this.state.delegatedToId,
      DelegatedFromId: item.Responsible.ID,
    }
    this._Service.updateItemById(this.props.siteUrl, this.props.workflowDetailsListName, item.ID, detialitem)
    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsListName).items.getById(item.ID).update({
      ResponsibleId: this.state.delegatedToId,
      DelegatedFromId: item.Responsible.ID,
    }); */
    //getting hub site id
    this._Service.getUserIdByEmail(this.state.delegateToEmail)
      //sp.web.siteUsers.getByEmail(this.state.delegateToEmail).get()
      .then(async delegateToIdInHubsite => {
        console.log("delegateTo ID in hubsite", delegateToIdInHubsite.Id);
        this._Service.getUserIdByEmail(item.Responsible.EMail)
          //sp.web.siteUsers.getByEmail(item.Responsible.EMail).get()
          .then(async responsibleIdInHubsite => {
            console.log("responsible ID in hubsite", responsibleIdInHubsite.Id);
            const taskitem = {
              AssignedToId: delegateToIdInHubsite.Id,
              DelegatedOn: date,
              DelegatedFromId: responsibleIdInHubsite.Id,
              StartDate: this.currentDate,
            }
            await this._Service.updateItemById(this.props.siteUrl, this.props.workflowTaskListName, item.TaskID, taskitem)
            /* await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowTaskListName).items.getById(item.TaskID).update
              ({
                AssignedToId: delegateToIdInHubsite.Id,
                DelegatedOn: date,
                DelegatedFromId: responsibleIdInHubsite.Id,
                StartDate: this.currentDate,
              }); */
            if (item.ResponseStatus == "Under Approval") {
              type = "DocApproval";
              link = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + item.ID;
              this.triggerDocumentDelegate(this.sourceDocumentID, this.state.delegateToEmail, item.Responsible.EMail, link, type, this.headerId, this.documentIndexID).then(afterTrigger => {
                const headitem = {
                  ApproverId: this.state.delegatedToId
                }
                this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headitem)
                /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                  ({                   //headerlist
                    ApproverId: this.state.delegatedToId
                  }); */
                const inditem = {
                  ApproverId: this.state.delegatedToId
                }
                this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID, inditem)
                /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).update
                  ({
                    ApproverId: this.state.delegatedToId
                  }); */
              });
            }
            else if (item.ResponseStatus == "Under Review") {
              if (item.Workflow == "DCC Review") {
                type = "DCCReview";
                link = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + item.ID;
                this.triggerDocumentDelegate(this.sourceDocumentID, this.state.delegateToEmail, item.Responsible.EMail, link, type, this.headerId, this.documentIndexID).
                  then(afterTrigger => {
                    const headdata = {                   //headerlist
                      DocumentControllerId: this.state.delegatedToId
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.workflowHeaderListName, this.headerId, headdata)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderListName).items.getById(this.headerId).update
                      ({                   //headerlist
                        DocumentControllerId: this.state.delegatedToId
                      }); */
                    const inddata = {
                      DocumentControllerId: this.state.delegatedToId
                    }
                    this._Service.updateItemById(this.props.siteUrl, this.props.documentIndexListName, this.documentIndexID, inddata)
                    /* sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexListName).items.getById(this.documentIndexID).update
                      ({
                        DocumentControllerId: this.state.delegatedToId
                      }); */
                  });
              }
              else {
                type = "DocReview";
                link = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + item.ID;
                this.triggerDocumentDelegate(this.sourceDocumentID, this.state.delegateToEmail, item.Responsible.EMail, link, type, this.headerId, this.documentIndexID);
              }
            }
          });
      });
  }
  //mail for delegation
  private async SendAnEmailForDelagation(email, type, name, Link): Promise<void> {
    // alert(name);
    let Subject;
    let Body;
    let link;
    link = `<a href=${Link}>Link</a>`;
    console.log(link);
    //console.log(queryVar);
    const notificationPreference: any[] = await this._Service.getNotificationPref(this.props.siteUrl, this.props.notificationPrefListName, email)
    //const notificationPreference: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.notificationPrefListName).items.select("Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail").expand("EmailUser").filter("EmailUser/EMail eq '" + email + "'").get();
    // console.log(notificationPreference);
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference == "Send all emails") {
        this.status = "Yes";
        //console.log("Send mail for all");                 
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.criticalDocument == true) {
        //console.log("Send mail for critical document");
        this.status = "Yes";
      }
      else {
        this.status = "No";
      }
    }
    else if (this.state.criticalDocument == true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    //Email Notification Settings.
    if (type == "DCCReview") {
      type = "DocReview";
    }
    const emailNoficationSettings: any[] = await this._Service.getEmailNoficationSettings(this.props.siteUrl, this.props.emailNoficationSettings, type)
    //const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNoficationSettings).items.filter("Title eq '" + type + "'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    //Replacing the email body with current values
    let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
    let duedate = moment(this.dueDateForModel).format("DD/MM/YYYY");
    let replacedSubjectWithDueDate = replaceString(replacedSubject, '[DueDate]', duedate);
    let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
    let replacelink = replaceString(replaceBody, '[Link]', link);
    let var1: any[] = replacelink.split('/');
    // alert(var1[0]);
    let FinalBody = replacelink;
    //mail sending
    if (this.status == "Yes") {
      //Check if TextField value is empty or not  
      if (email) {
        //Create Body for Email  
        let emailPostBody: any = {
          "message": {
            "subject": replacedSubjectWithDueDate,
            "body": {
              "contentType": "HTML",
              "content": FinalBody
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

  //mail for delegation
  private async sendEmailForCurrentTaskOwner(email, type, name): Promise<void> {
    // alert(name);
    let Subject;
    let Body;
    let link;
    console.log(link);
    //console.log(queryVar);
    const notificationPreference: any[] = await this._Service.getNotificationPref(this.props.siteUrl, this.props.notificationPrefListName, email)
    //const notificationPreference: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.notificationPrefListName).items.select("Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail").expand("EmailUser").filter("EmailUser/EMail eq '" + email + "'").get();
    // console.log(notificationPreference);
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference == "Send all emails") {
        this.status = "Yes";
        //console.log("Send mail for all");                 
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.criticalDocument == true) {
        //console.log("Send mail for critical document");
        this.status = "Yes";
      }
      else {
        this.status = "No";
      }
    }
    else if (this.state.criticalDocument == true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await this._Service.getEmailNoficationSettings(this.props.siteUrl, this.props.emailNoficationSettings, type)
    //const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNoficationSettings).items.filter("Title eq '" + type + "'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    console.log(Body);
    //Replacing the email body with current values
    let c = Subject;
    let replaceRequester = replaceString(Body, '[Sir/Madam]', this.state.taskOwnerName);
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
    let replacewithToUser = replaceString(replaceBody, '[ToUser]', this.state.delegateToTitle);
    let replacewithCurrentUser = replaceString(replacewithToUser, '[CurrentUser]', this.state.currentUserName);
    // alert(var1[0]);
    // let FinalBody = replacelink;
    //mail sending
    if (this.status == "Yes") {
      //Check if TextField value is empty or not  
      if (email) {
        //Create Body for Email  
        let emailPostBody: any = {
          "message": {
            "subject": Subject,
            "body": {
              "contentType": "HTML",
              "content": replacewithCurrentUser
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

  public render(): React.ReactElement<IRevisionHistoryProps> {

    return (
      <section className={`${styles.revisionHistory}`}>
        <div>
          <div style={{ textAlign: "center", fontWeight: "bold", display: this.state.timelineDisplay }}  >
            <Timeline lineColor={this.props.timeLineColor}>

              {this.state.logItems.map((items, key) => {
                if (items.Status == "Cancelled") {
                  return (
                    this.Cancelled(items)
                  );
                }
                else if (items.Status == "Document Expired") {
                  return (
                    this.documentExpired(items)
                  );
                }
                else if (items.Status == "DCC Review - Returned with comments" || items.Status == "DCC - Reviewed" || items.Status == "Under Review" && items.Workflow == "DCC Review") {
                  return (
                    this.dccReview(items)
                  );
                }
                else if (items.Status == "Document Expiry Revoked") {
                  return (
                    this.documentExpiryRevoked(items)
                  );
                }
                else if (items.Status == "Document Archived" || items.Status == "Void Under Approval" && items.Workflow == "Void") {
                  return (
                    this.documentVoided(items)
                  );
                }
                else if (items.Status == "Published" && items.Workflow == null) {
                  return (
                    this.directPublish(items)
                  );
                }
                else if ((items.Status == "Published" || items.Status == "Under Approval" || items.Status == "Returned with comments" || items.Status == "Rejected") && items.Workflow == "Approval") {
                  return (
                    this.approved(items)
                  );
                }
                else if ((items.Status == "Reviewed" || items.Status == "Under Review" || items.Status == "Returned with comments") && items.Workflow == "Review") {
                  return (
                    this.reviewed(items)
                  );
                }
                else if (items.Status == "Document Void Initiated") {
                  return (
                    this.workFlowStartedVoid(items)
                  );
                }
                else if (items.Status == "Workflow Initiated" || items.Status == "Document Void Initiated") {
                  return (
                    this.workFlowStarted(items)
                  );
                }
                else if (items.Status == "Document Created") {
                  return (
                    this.documentCreated(items)
                  );
                }
                else if (items.Status == this.props.internalTransittalConFab || items.Status == "Internally transmitted to Construction/Fabrication") {
                  return (
                    this.internalTransittalConFab(items)
                  );
                }
              })}
            </Timeline>


            <div style={{ display: this.state.reviewed }}>
              <Modal
                isOpen={this.state.showReviewModal}
                onDismiss={this._closeModal}
                containerClassName={contentStyles.container}>
                <div className={contentStyles.header1}>
                  {(this.state.timelineElement == "Review") ?
                    <><div style={{ textAlign: "left", width: "50%", fontSize: "20px", fontWeight: "bold" }}>Review Details</div><div style={{ textAlign: "right", width: "50%", fontSize: "14px" }}> DueDate: {moment(this.dueDateForModel).format("DD/MM/YYYY")} </div></> :
                    <><div style={{ textAlign: "left", width: "50%", fontSize: "20px", fontWeight: "bold" }}>Approval Details </div><div style={{ textAlign: "right", width: "50%", fontSize: "14px" }}> DueDate : {moment(this.dueDateForModel).format("DD/MM/YYYY")} </div></>
                  }
                  <IconButton
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={this._closeModal}
                    styles={iconButtonStyles}
                  />
                </div>
                <div style={{ padding: "0 20px " }}>
                  <div>
                    {this.state.statusMessage.isShowMessage ?
                      <MessageBar
                        messageBarType={this.state.statusMessage.messageType}
                        isMultiline={false}
                        dismissButtonAriaLabel="Close"
                      >{this.state.statusMessage.message}</MessageBar>
                      : ''}
                    <div style={{ display: this.state.delegatingLoader }}>
                      <Spinner size={SpinnerSize.large} label="Delegating....." labelPosition="left" />
                    </div>
                    <div style={{ display: this.loaderforcancel }}>
                      <Spinner size={SpinnerSize.large} label="Canceling....." labelPosition="left" />
                    </div>
                  </div>
                  <table className={styles.tableModal}>
                    <tr style={{ background: "#f4f4f4" }}>
                      {(this.state.timelineElement == "Review") ?
                        <><th style={{ padding: "5px 10px" }}>Reviewer</th><th style={{ padding: "5px 10px" }}>Reviewed Date</th></> :
                        <><th style={{ padding: "5px 10px" }}>Approver</th><th style={{ padding: "5px 10px" }}>Approved Date</th></>}
                      <th style={{ padding: "5px 10px" }}>Status</th>
                      <th style={{ padding: "5px 10px" }}>Comments</th>
                      <th style={{ padding: "5px 10px" }} hidden={this.state.divForSendRemainder}>Reminder</th>
                      <th style={{ padding: "5px 10px" }} hidden={this.state.divForCancel}>Cancel</th>
                      <th style={{ padding: "5px 10px" }} hidden={this.state.divForDelegation}>Delegate</th>

                    </tr>
                    {this.state.workflowDetailItems.map((wfItems, key) => {
                      return (
                        <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                          <td style={{ padding: "5px 10px" }}>{wfItems.Responsible.Title}</td>
                          <td style={{ padding: "5px 10px" }}>{(wfItems.ResponseDate == null) ? "" : moment(wfItems.ResponseDate).format("DD/MM/YYYY hh:mm")}</td>
                          <td style={{ padding: "5px 10px" }}>{(wfItems.ResponseStatus !== null) ? wfItems.ResponseStatus : "Pending"}</td>
                          <td style={{ padding: "5px 10px" }}>
                            <TooltipHost
                              content={(wfItems.ResponsibleComment !== null) ? wfItems.ResponsibleComment : "No Comments"}
                              // This id is used on the tooltip itself, not the host
                              // (so an element with this id only exists when the tooltip is shown)                              
                              calloutProps={calloutProps}
                              styles={hostStyles}
                            >
                              <IconButton iconProps={Comment} title=" " ariaLabel=" " />
                            </TooltipHost>
                          </td>

                          <td style={{ padding: "5px 10px" }} hidden={this.state.divForSendRemainder}><TooltipHost
                            content="Send Reminder"
                            // This id is used on the tooltip itself, not the host
                            // (so an element with this id only exists when the tooltip is shown)                              
                            calloutProps={calloutProps}
                            styles={hostStyles}
                          >
                            <IconButton iconProps={ReminderTime} title=" " ariaLabel=" " disabled={(wfItems.ResponseStatus == "Under Review" || wfItems.ResponseStatus == "Under Approval") ? false : true} onClick={() => this.sendRemainder(wfItems, key)} />
                          </TooltipHost>
                          </td>
                          <td style={{ padding: "5px 10px", display: (wfItems.ResponseStatus == "Void Under Approval") ? "none" : "" }} hidden={this.state.divForCancel}><TooltipHost
                            content={(wfItems.ResponseStatus == "Cancelled") ? "Cancelled by : " + wfItems.Editor.Title : "Cancel"}
                            // (so an element with this id only exists when the tooltip is shown)                              
                            calloutProps={calloutProps}
                            styles={hostStyles}
                          >
                            <IconButton iconProps={cancelIcon} title=" " ariaLabel=" " disabled={(wfItems.ResponseStatus == "Under Review" || wfItems.ResponseStatus == "Under Approval") ? false : true} onClick={() => this.cancelTask(wfItems, key)} />
                          </TooltipHost>
                          </td>
                          <td style={{ padding: "5px 10px" }} hidden={this.state.divForDelegation}><TooltipHost
                            content="Share"
                            // This id is used on the tooltip itself, not the host
                            // (so an element with this id only exists when the tooltip is shown)                              
                            calloutProps={calloutProps}
                            styles={hostStyles}
                          >
                            <IconButton iconProps={Share} title=" " ariaLabel=" " disabled={(wfItems.ResponseStatus == "Under Review" || wfItems.ResponseStatus == "Under Approval") ? false : true} onClick={() => this._delegateClick(wfItems, key)} />
                          </TooltipHost>
                          </td>
                          <td style={{ padding: "5px 10px" }}> <div style={{ display: (wfItems.ResponseStatus == "Under Review" || wfItems.ResponseStatus == "Under Approval") ? this.state.delagatePeoplePicker : "none" }}>
                            <div style={{ display: "flex" }}>
                              <PeoplePicker
                                context={this.props.context as any}
                                // placeholder="Delegate to"
                                titleText="Delegate to "
                                personSelectionLimit={1}
                                groupName={""} // Leave this blank in case you want to filter from all users    
                                showtooltip={true}
                                disabled={false}
                                ensureUser={true}
                                // onChange={this._getPeoplePickerItems}
                                //selectedItems={this._getPeoplePickerItems}
                                //defaultSelectedUsers={[this.state.approver]}
                                showHiddenInUI={false}
                                //isRequired={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={1000}
                              />
                              <div style={{ marginLeft: "20px" }}>
                                <PrimaryButton text="Delegate" onClick={() => this._delegateSubmit(wfItems, key)} />
                              </div>
                            </div>
                          </div>
                          </td>
                        </tr>
                      );
                    })}
                  </table>
                </div>
              </Modal>
            </div>
            <div style={{ display: this.state.workflowInitiated }}>
              <Modal
                isOpen={this.state.showworkflowInitiatedModal}
                onDismiss={this._closeModal}
                containerClassName={contentStyles.container}>
                <div className={contentStyles.header1}>
                  <div style={{ display: "block" }}><div style={{ textAlign: "left", width: "100%", fontSize: "10px" }}> </div><div style={{ textAlign: "center", width: "100%", fontSize: "20px", fontWeight: "bold" }}>Workflow Initiated </div></div>
                  <IconButton
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={this._closeModal}
                    styles={iconButtonStyles}
                  />
                </div>
                <div style={{ padding: "0 20px 20px" }}>
                  <div style={{ width: "100%", marginBottom: "10px" }}>Requestor :{this.state.requestor}</div>
                  <div style={{ width: "100%", marginBottom: "10px" }}>Requested Date :{this.state.requestorDate}</div>
                  <div style={{ width: "100%", marginBottom: "10px" }}>Approver : {this.state.approverName}</div>
                  <table className={styles.tableModal} style={{ display: this.state.lengthOfReviwers == null ? "none" : "" }}>
                    <tr>
                      <th>Reviewers </th>
                    </tr>
                    {this.state.reviewers.map((reviewers, key) => {
                      return (
                        <tr>
                          <td>{reviewers.Title}</td>
                        </tr>
                      );
                    })}
                  </table>
                </div>
              </Modal>
            </div>
            <div style={{ display: this.state.workflowInitiatedVoid }}>
              <Modal
                isOpen={this.state.showworkflowInitiatedVoidModal}
                onDismiss={this._closeModal}
                containerClassName={contentStyles.container}>
                <div className={contentStyles.header1}>
                  <div style={{ display: "block" }}><div style={{ textAlign: "left", width: "100%", fontSize: "10px" }}> </div><div style={{ textAlign: "center", width: "100%", fontSize: "20px", fontWeight: "bold" }}>Workflow Initiated </div></div>
                  <IconButton
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={this._closeModal}
                    styles={iconButtonStyles}
                  />
                </div>
                <div style={{ padding: "0 20px 20px" }}>
                  <div style={{ width: "100%", marginBottom: "10px" }}>Requester :{this.state.requestor}</div>
                  <div style={{ width: "100%", marginBottom: "10px" }}>Requested Date :{this.state.requestorDate}</div>
                  <div style={{ width: "100%", marginBottom: "10px" }}>Approver : {this.state.approverName}</div>

                </div>
              </Modal>
            </div>
          </div>
          <div>{this.state.statusMessage.isShowMessage ?
            <MessageBar
              messageBarType={this.state.statusMessage.messageType}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >{this.state.statusMessage.message}</MessageBar>
            : ''}</div>

        </div>
      </section>
    );
  }
}
