import { IMessage } from "../interfaces";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISendRequestProps {
    context: WebPartContext;
    siteUrl: string;
    redirectUrl: string;
    hubUrl: string;
    userMessageSettings: string;
    documentIndexList: string;
    project: string;
    notificationPreference: string;
    emailNotification: string;
    workflowHeaderList: string;
    transmittalCodeSettingsList: string;
    workflowDetailsList: string;
    documentRevisionLogList: string;
    workflowTasksList: string;
    sourceDocumentLibrary: string;
    revisionLevelList: string;
    taskDelegationSettings: string;
    revisionHistoryPage: string;
    documentApprovalPage: string;
    documentReviewPage: string;
    accessGroups: string;
    departmentList: string;
    accessGroupDetailsList: string;
    hubsite: string;
    projectInformationListName: string;
    businessUnitList: string;
    webpartHeader: string;
    siteAddress: string;
    requestList: string;
    sourceDocumentLibraryView: string;
}

export interface ISendRequestState {
    statusMessage: IMessage;
    documentID: string;
    linkToDoc: string;
    documentName: string;
    revision: any;
    ownerName: string;
    currentUser: string;
    hideProject: boolean;
    revisionLevel: any[];
    revisionLevelvalue: any;
    dcc: any;
    reviewer: any;
    dueDate: any;
    approver: any;
    comments: any;
    cancelConfirmMsg: string;
    confirmDialog: boolean;
    saveDisable: string;
    requestSend: string;
    statusKey: string;
    access: any;
    accessDeniedMsgBar: any;
    reviewers: any[];
    currentUserReviewer: any[];
    ownerId: any;
    delegatedToId: any;
    delegateToIdInSubSite: any;
    delegateForIdInSubSite: any;
    reviewerEmail: any;
    reviewerName: any;
    delegatedFromId: any;
    detailIdForReviewer: any;
    approverEmail: any;
    approverName: any;
    hubSiteUserId: any;
    detailIdForApprover: any;
    criticalDocument: any;
    dccReviewerName: any;
    dccReviewerEmail: any;
    dccReviewer: any;
    revisionLevelArray: any[];
    revisionCoding: any;
    projectName: any;
    projectNumber: any;
    acceptanceCodeId: any;
    transmittalRevision: any;
    reviewersName: any[];
    hideLoading: boolean;
    sameRevision: any;
    loaderDisplay: string;
    businessUnitID: any;
    departmentId: any;
    validApprover: string;
    hideCreateLoading: string;
    
}

export interface ISendRequestWebPartProps {
    context: WebPartContext;
    siteUrl: string;
    redirectUrl: string;
    hubUrl: string;
    userMessageSettings: string;
    documentIndexList: string;
    project: string;
    notificationPreference: string;
    emailNotification: string;
    workflowHeaderList: string;
    transmittalCodeSettingsList: string;
    workflowDetailsList: string;
    documentRevisionLogList: string;
    workflowTasksList: string;
    sourceDocumentLibrary: string;
    revisionLevelList: string;
    taskDelegationSettings: string;
    revisionHistoryPage: string;
    documentApprovalPage: string;
    documentReviewPage: string;
    accessGroups: string;
    departmentList: string;
    accessGroupDetailsList: string;
    hubsite: string;
    projectInformationListName: string;
    businessUnitList: string;
    webpartHeader: string;
    siteAddress: string;
    requestList: string;
    sourceDocumentLibraryView: string;
}