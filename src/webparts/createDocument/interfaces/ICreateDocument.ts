import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICreateDocumentProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  description: string;
  context: WebPartContext;
  hasTeamsContext: boolean;
  siteUrl: string;
  webpartHeader: string;
  documentIndexList: string;
  sourceDocumentLibrary: string;
  businessUnit: string;
  department: string;
  category: string;
  subCategory: string;
  legalEntity: string;
  userMessageSettings: string;
  publisheddocumentLibrary: string;
  requestList: string;
  documentIdSettings: string;
  documentIdSequenceSettings: string;
  documentRevisionLogList: string;
  revisionHistoryPage: string;
  revokePage: string;
  notificationPreference: string;
  emailNotification: string;
  directPublish: boolean;
  hubUrl: string;
  QDMSUrl: string;
}
export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}
export interface ICreateDocumentState {
  statusMessage: IMessage;
  title: any;
  approvalDate: any;
  loaderDisplay: string;
  businessUnitOption: any[];
  departmentOption: any[];
  categoryOption: any[];
  legalEntityOption: any[];
  owner: any;
  ownerEmail: any;
  ownerName: any;
  documentName: any;
  saveDisable: boolean;
  businessUnitID: any;
  departmentId: any;
  categoryId: any;
  businessUnit: any;
  businessUnitCode: any;
  departmentCode: any;
  department: any;
  subCategoryArray: any[];
  subCategoryId: any;
  category: any;
  subCategory: any;
  categoryCode: any;
  legalEntityId: any;
  legalEntity: any;
  approver: any;
  approverEmail: any;
  approverName: any;
  reviewers: any[];
  validApprover: string;
  hideDoc: string;
  createDocument: boolean;
  hideDirect: string;
  upload: boolean;
  checkdirect: string;
  insertdocument: string;
  hidePublish: string;
  directPublishCheck: boolean;
  hideupload: string;
  template: boolean;
  hidesource: string;
  hidetemplate: string;
  templateDocuments: any;
  isdocx: string;
  nodocx: string;
  sourceId: string;
  templateId: any;
  templateKey: any;
  approvalDateEdit: Date;
  publishOption: any;
  hideExpiry: string;
  expiryCheck: any;
  expiryDate: any;
  expiryLeadPeriod: any;
  leadmsg: string;
  criticalDocument: boolean;
  templateDocument: boolean;
  hideLoading: boolean;
  hideCreateLoading: string;
  norefresh: string;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  hideloader: boolean;
  documentid: any;
  incrementSequenceNumber: any;
  sourceDocumentId: any;
  newDocumentId: any;
  newRevision: any;
  messageBar: string;
  dateValid: string;
  uploadOrTemplateRadioBtn: any;
  showReviewModal: boolean;
  DueDate: Date,
  sendForReview: boolean,
  dueDateMadatory: string;
  comments: string;

}
export interface ICreateDocumentWebPartProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  description: string;
  context: WebPartContext;
  hasTeamsContext: boolean;
  siteUrl: string;
  webpartHeader: string;
  documentIndexList: string;
  sourceDocumentLibrary: string;
  businessUnit: string;
  department: string;
  category: string;
  subCategory: string;
  legalEntity: string;
  userMessageSettings: string;
  publisheddocumentLibrary: string;
  requestList: string;
  documentIdSettings: string;
  documentIdSequenceSettings: string;
  documentRevisionLogList: string;
  revisionHistoryPage: string;
  revokePage: string;
  notificationPreference: string;
  emailNotification: string;
  directPublish: boolean;
  hubUrl: string;
  QDMSUrl: string;
}
