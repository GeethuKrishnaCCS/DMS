import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IMessage } from "./IMessage";

export interface IEditDocumentProps {
    description: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
    context: WebPartContext;
    siteUrl: string;
    redirectUrl: string;
    userMessageSettings: string;
    documentIndexList: string;
    notificationPreference: string;
    emailNotification: string;
    businessUnit: string;
    department: string;
    category: string;
    subCategory: string;
    publisheddocumentLibrary: string;
    documentIdSettings: string;
    documentIdSequenceSettings: string;
    sourceDocumentLibrary: string;
    revisionHistoryPage: string;
    siteAddress: string;
    sourceDocumentViewLibrary: string;
    documentRevisionLogList: string;
    backUrl: string;
    revokePage: string;
    legalEntity: string;
    departmentList: string;
    businessUnitList: string;
    requestList: string;
    webpartHeader: string;
    QDMSUrl: string;
    
}

export interface IEditDocumentState {
    statusMessage: IMessage;
    title: any;
    hideProject: string;
    businessUnitID: any;
    departmentId: any;
    categoryId: any;
    subCategoryKey: any;
    createDocument: boolean;
    directPublishCheck: boolean;
    criticalDocument: boolean;
    replaceDocumentCheckbox: boolean;
    templateDocument: boolean;
    templateDocuments: any;
    approvalDate: any;
    dateValid: string;
    uploadOrTemplateRadioBtn: any;
    publishOptionKey: any;
    hideDoc: string;
    hidePublish: string;
    hideExpiry: string;
    expiryCheck: any;
    expiryDate: any;
    expiryLeadPeriod: any;
    cancelConfirmMsg: string;
    confirmDialog: boolean;
    saveDisable: boolean;
    createDocumentView: string;
    createDocumentProject: string;
    businessUnitOption: any[];
    departmentOption: any[];
    categoryOption: any[];
    businessUnitCode: any;
    departmentCode: any;
    subCategoryArray: any[];
    subCategoryId: any;
    reviewers: any[];
    approver: any;
    approverEmail: any;
    approverName: any;
    owner: any;
    ownerEmail: any;
    ownerName: any;
    templateId: any;
    publishOption: any;
    incrementSequenceNumber: any;
    documentid: any;
    documentName: any;
    businessUnit: any;
    category: any;
    subCategory: any;
    department: any;
    newDocumentId: any;
    sourceDocumentId: any;
    templateKey: any;
    dcc: any;
    dccEmail: any;
    dccName: any;
    revisionCoding: any;
    revisionLevel: any;
    transmittalCheck: boolean;
    externalDocument: boolean;
    revisionCodingId: any;
    revisionLevelId: any;
    revisionLevelArray: any[];
    revisionSettingsArray: any[];
    categoryCode: any;
    projectName: any;
    projectNumber: any;
    messageBar: string;
    projectEditDocumentView: string;
    qdmsEditDocumentView: string;
    revokeExpiryView: string;
    hideCreate: string;
    reviewersName: any;
    approvalDateEdit: Date;
    currentRevision: any;
    previousRevisionItemID: any;
    revisionItemID: any;
    newRevision: any;
    hideloader: boolean;
    legalEntityOption: any[];
    legalEntityId: any;
    legalEntity: any;
    updateDisable: boolean;
    hideLoading: boolean;
    workflowStatus: any;
    accessDeniedMessageBar: string;
    titleReadonly: boolean;
    leadmsg: string;
    invalidQueryParam: string;
    hidebutton: string;
    isdocx: string;
    nodocx: string;
    insertdocument: string;
    loaderDisplay: string;
    checkdirect: string;
    hideDirect: string;
    validApprover: string;
    createDocumentCheckBoxDiv: string;
    replaceDocument: string;
    hideSelectTemplate: string;
    validDocType: string;
    checkrename: string;
    subContractorNumber: string;
    customerNumber: string;
    linkToDoc: string;
    hideCreateLoading: string;
    norefresh: string;
    upload: boolean;
    template: boolean;
    hideupload: string;
    sourceId: string;
    hidetemplate: string;
    hidesource: string;
    showReviewModal: boolean;
    DueDate: Date,
    sendForReview: boolean,
    dueDateMadatory: string;
    comments: string;
    mydoc: File | null;
    
}

export interface IEditDocumentWebPartProps {
    description: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
    context: WebPartContext;
    siteUrl: string;
    redirectUrl: string;
    hubUrl: string;
    hubsite: string;
    userMessageSettings: string;
    documentIndexList: string;
    notificationPreference: string;
    emailNotification: string;
    businessUnit: string;
    department: string;
    category: string;
    subCategory: string;
    publisheddocumentLibrary: string;
    documentIdSettings: string;
    documentIdSequenceSettings: string;
    sourceDocumentLibrary: string;
    revisionHistoryPage: string;
    siteAddress: string;
    sourceDocumentViewLibrary: string;
    documentRevisionLogList: string;
    revisionLevelList: string;
    revisionSettingsList: string;
    projectInformationListName: string;
    backUrl: string;
    transmittalHistory: string;
    revokePage: string;
    legalEntity: string;
    permissionMatrix: string;
    departmentList: string;
    accessGroupDetailsList: string;
    businessUnitList: string;
    requestList: string;
    webpartHeader: string;
    QDMSUrl: string;
}