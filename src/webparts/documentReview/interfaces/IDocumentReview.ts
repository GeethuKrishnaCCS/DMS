import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentReviewProps {
  description: string;
  webPartName: string;
  project: string;
  redirectUrl: string;
  siteUrl: string;
  workflowHeaderListName: string;
  context: WebPartContext;
  notificationPrefListName: string;
  hubSiteUrl: string;
  hubsite: string;
  emailNotificationSettings: string;
  userMessageSettings: string;
  documentIndex: string;
  workFlowDetail: string;
  documentApprovalSitePage: string;
  documentRevisionLog: string;
  documentReviewSitePage: string;
  workflowTaskListName: string;
  taskDelegationSettingsListName: string;
  accessGroups: string;
  departmentList: string;
  accessGroupDetailsList: string;
  projectInformationListName: string;
  bussinessUnitList: string;
  masterListName: string;
}

export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}

export interface IDocumentReviewState {
  statusMessage: IMessage;
  currentUser: number;
  status: string;
  statusKey: string;
  comments: string;
  reviewerItems: any[];
  access: string;
  accessDeniedMsgBar: string;
  documentIndexItems: any[];
  documentID: string;
  linkToDoc: string;
  documentName: string;
  revision: string;
  owner: string;
  requestor: string;
  documentControllerName: string;
  requestorComment: string;
  dueDate: any;
  requestorDate: string;
  workflowStatus: string;
  hideReviewersTable: string;
  detailListID: number;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  approverEmail: string;
  requestorEmail: string;
  ownerEmail: string;
  documentControllerEmail: string;
  ownerID: any;
  headerListItem: any[];
  notificationPreference: string;
  criticalDocument: boolean;
  approverName: string;
  userMessageSettings: any[];
  currentUserEmail: string;
  invalidMessage: string;
  pageLoadItems: any[];
  buttonHidden: string;
  approverId: any;
  reviewPending: string;
  DueDate: Date;
  detailIdForApprover: any;
  hubSiteUserId: any;
  delegatedFromId: any;
  delegatedToId: any;
  divForDCC: string;
  divForReview: string;
  ifDccComment: string;
  dcc: string;
  dccComment: string;
  dccCompletionDate: string;
  revisionLogID: any;
  delegateToIdInSubSite: any;
  delegateForIdInSubSite: any;
  noAccess: string;
  invalidQueryParam: string;
  projectName: string;
  projectNumber: string;
  hideproject: boolean;
  reviewers: any[];
  dccReviewItems: any[];
  currentReviewComment: string;
  currentReviewItems: any[];
  loaderDisplay: string;

}

export interface IDocumentReviewWebPartProps {
  description: string;
  webPartName: string;
  project: string;
  redirectUrl: string;
  siteUrl: string;
  workflowHeaderListName: string;
  context: WebPartContext;
  notificationPrefListName: string;
  hubSiteUrl: string;
  hubsite: string;
  emailNotificationSettings: string;
  userMessageSettings: string;
  documentIndex: string;
  workFlowDetail: string;
  documentApprovalSitePage: string;
  documentRevisionLog: string;
  documentReviewSitePage: string;
  workflowTaskListName: string;
  taskDelegationSettingsListName: string;
  accessGroups: string;
  departmentList: string;
  accessGroupDetailsList: string;
  projectInformationListName: string;
  bussinessUnitList: string;
  masterListName: string;
}