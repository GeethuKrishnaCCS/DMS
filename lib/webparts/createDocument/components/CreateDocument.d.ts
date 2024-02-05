import * as React from 'react';
import { ICreateDocumentProps, ICreateDocumentState } from '../interfaces';
export default class CreateDocument extends React.Component<ICreateDocumentProps, ICreateDocumentState> {
    private _Service;
    private validator;
    private siteUrl;
    private currentEmail;
    private currentId;
    private currentUser;
    private today;
    private createDocument;
    private directPublish;
    private getSelectedReviewers;
    private myfile;
    private isDocument;
    private permissionpostUrl;
    private documentNameExtension;
    private revokeUrl;
    private Timeout;
    private documentIndexID;
    private revisionHistoryUrl;
    private postUrl;
    constructor(props: ICreateDocumentProps);
    componentWillMount: () => void;
    componentDidMount(): Promise<void>;
    _bindData(): Promise<void>;
    private _userMessageSettings;
    _titleChange: (ev: React.FormEvent<HTMLInputElement>, title?: string) => void;
    _departmentChange(option: {
        key: any;
        text: any;
    }): Promise<void>;
    _categoryChange(option: {
        key: any;
        text: any;
    }): Promise<void>;
    _subCategoryChange(option: {
        key: any;
        text: any;
    }): void;
    _legalEntityChange(option: {
        key: any;
        text: any;
    }): void;
    _selectedOwner: (items: any[]) => void;
    _selectedReviewers: (items: any[]) => void;
    _selectedApprover: (items: any[]) => Promise<void>;
    _onCreateDocChecked: (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => Promise<void>;
    private onUploadOrTemplateRadioBtnChange;
    private _onUploadCheck;
    private _onTemplateCheck;
    _add(e: any): void;
    _sourcechange(option: {
        key: any;
        text: any;
    }): Promise<void>;
    _templatechange(option: {
        key: any;
        text: any;
    }): Promise<void>;
    private _onDirectPublishChecked;
    _checkdirectPublish(type: any): Promise<void>;
    _onApprovalDatePickerChange: (date: Date) => void;
    _publishOptionChange(option: {
        key: any;
        text: any;
    }): void;
    _onExpiryDateChecked: (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => void;
    _onExpDatePickerChange: (date?: Date) => void;
    _expLeadPeriodChange: (ev: React.FormEvent<HTMLInputElement>, expiryLeadPeriod: string) => void;
    _onCriticalChecked: (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => void;
    _onTemplateChecked: (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => void;
    _onCreateDocument(): Promise<void>;
    _documentidgeneration(): Promise<void>;
    _incrementSequenceNumber(incrementvalue: any, sequenceNumber: any): void;
    _documentCreation(): Promise<void>;
    _createDocumentIndex(): void;
    _addSourceDocument(): Promise<void>;
    protected _triggerPermission(sourceDocumentID: any): Promise<void>;
    protected _triggerSendForReview(sourceDocumentID: any, documentIndexId: any): Promise<void>;
    protected _publish(): Promise<void>;
    _revisionCoding: () => Promise<void>;
    _publishUpdate(): Promise<void>;
    _sendMail: (emailuser: any, type: any, name: any) => Promise<void>;
    private _onCancel;
    private _confirmYesCancel;
    private _confirmNoCancel;
    private _dialogCloseButton;
    private dialogStyles;
    private dialogContentProps;
    private modalProps;
    private _onFormatDate;
    private _closeModal;
    _DueDateChange: (date: Date) => void;
    _commentChange: (ev: React.FormEvent<HTMLInputElement>, comments?: string) => void;
    private _onSendForReview;
    onConfirmReview: () => Promise<void>;
    render(): React.ReactElement<ICreateDocumentProps>;
}
//# sourceMappingURL=CreateDocument.d.ts.map