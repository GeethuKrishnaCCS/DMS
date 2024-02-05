import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IPaplDmsProps {
    description: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
    libraryConfig: any[];
    siteUrl: string;
    context: WebPartContext;
}
export interface IPaplDmsWebPartProps {
    description: string;
    listandlibraries: any;
    siteUrl: string;
    context: WebPartContext;
}
export interface IPaplDmsState {
    title: string;
    departmentId: string;
    departmentItems: any[];
    categoryItems: any[];
    categoryId: string;
    auditFrequencyId: string;
    auditFrequencyItems: any[];
    attachedUrl: string;
    approverId: Number;
    reviewerId: string;
    reviewerName: string;
    approverName: string;
    reviewerEmail: string;
    departmentCode: string;
    fileUploaded: any[];
    uploadToggle: boolean;
    confirmDialog: boolean;
    showReviewModal: boolean;
    comments: string;
    DueDate: Date;
    sendForReview: boolean;
    statusMessage: IMessage;
    buttonHide: boolean;
    documentName: string;
    currentUserId: Number;
    incrementSequenceNumber: any;
    categoryTitle: string;
    reviewers: any[];
    approvers: any[];
    emailNotification: {
        subject: string;
        body: string;
    };
}
export interface IMessage {
    isShowMessage: boolean;
    messageType: number;
    message: string;
}
//# sourceMappingURL=IPaplDmsProps.d.ts.map