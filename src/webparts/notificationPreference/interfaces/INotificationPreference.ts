import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INotificationPreferenceProps {
  description: string;
  hubSiteUrl: string;
  notificationPrefListName: string;
  noEmail: string;
  sendForCriticalDocuments: string;
  sendForAllDocuments: string;
  userMessageSettings: string;
  defaultPreferenceText: string;
  noEmailText: string;
  sendForCriticalDocumentsText: string;
  sendForAllDocumentsText: string;
  context: WebPartContext;
}

export interface INotificationPreferenceState {
  notificationPreferenceKey: string;
  notificationPreferenceValue: string;
  currentUserId: any;
  currentUserLoginName: any;
  showMessage: any;
  message: any;
  defaultPreference: any;
  currentPreferenceItemID: any;
  messageMode: any;
}

export interface INotificationPreferenceWebPartProps {
  description: string;
  hubSiteUrl: string;
  notificationPrefListName: string;
  noEmail: string;
  sendForCriticalDocuments: string;
  sendForAllDocuments: string;
  userMessageSettings: string;
  defaultPreferenceText: string;
  noEmailText: string;
  sendForCriticalDocumentsText: string;
  sendForAllDocumentsText: string;
}