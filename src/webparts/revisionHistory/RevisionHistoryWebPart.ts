import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'RevisionHistoryWebPartStrings';
import RevisionHistory from './components/RevisionHistory';
import { IRevisionHistoryProps, IRevisionHistoryWebPartProps } from './interfaces';

export default class RevisionHistoryWebPart extends BaseClientSideWebPart<IRevisionHistoryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IRevisionHistoryProps> = React.createElement(
      RevisionHistory,
      {
        description: this.properties.description,
        timeLineColor: this.properties.timeLineColor,
        context: this.context,
        //project: this.properties.project,
        hubSiteUrl: this.properties.hubSiteUrl,
        hubSite: this.properties.hubSite,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        emailNoficationSettings: this.properties.emailNoficationSettings,
        DocumentRevisionLog: this.properties.DocumentRevisionLog,
        notificationPrefListName: this.properties.notificationPrefListName,
        documentApprovalDateColor: this.properties.documentApprovalDateColor,
        documentCreatedDateColor: this.properties.documentCreatedDateColor,
        documentReviewedDateColor: this.properties.documentReviewedDateColor,
        workflowStartedDateColor: this.properties.workflowStartedDateColor,
        documentApprovalContentColor: this.properties.documentApprovalContentColor,
        documentCreatedContentColor: this.properties.documentCreatedContentColor,
        documentReviewedContentColor: this.properties.documentReviewedContentColor,
        workflowStartedContentColor: this.properties.workflowStartedContentColor,
        documentVoidDateColor: this.properties.documentVoidDateColor,
        documentVoidContentColor: this.properties.documentVoidContentColor,
        documentApprovalSitePage: this.properties.documentApprovalSitePage,
        documentReviewSitePage: this.properties.documentReviewSitePage,
        statusColor: this.properties.statusColor,
        workflowDetailsListName: this.properties.workflowDetailsListName,
        workflowHeaderListName: this.properties.workflowHeaderListName,
        sourceDocuments: this.properties.sourceDocuments,
        documentIndexListName: this.properties.documentIndexListName,
        workflowTaskListName: this.properties.workflowTaskListName,
        taskDelegationListName: this.properties.taskDelegationListName,
        permissionMatrixSettings: this.properties.permissionMatrixSettings,
        accessGroupDetailsListName: this.properties.accessGroupDetailsListName,
        departmentListName: this.properties.departmentListName,
        bussinessUnitList: this.properties.bussinessUnitList,
        requestLaURL: this.properties.requestLaURL,
        internalTransittalConFab: this.properties.internalTransittalConFab
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                //PropertyPaneToggle('project', {
                //  label: 'Project',
                //  onText: 'On',
                //  offText: 'Off'
                // }),
                PropertyPaneTextField('timeLineColor', {
                  label: "Time Line Color"
                }),
                PropertyPaneTextField('statusColor', {
                  label: "Status Heading Color"
                }),
                PropertyPaneTextField('documentApprovalSitePage', {
                  label: "Document Approval SitePage"
                }), PropertyPaneTextField('documentReviewSitePage', {
                  label: "Document Review SitePage"
                }),
                PropertyPaneTextField('hubSite', {
                  label: "HubSitePage"
                }),

                PropertyPaneTextField('hubSiteUrl', {
                  label: "Hub Site Url"
                }),
              ],
            },
            {
              groupName: "List Names",
              groupFields: [
                PropertyPaneTextField('requestLaURL', {
                  label: "LA Url listname"
                }),
                PropertyPaneTextField('workflowDetailsListName', {
                  label: "Work details listname"
                }),
                PropertyPaneTextField('workflowHeaderListName', {
                  label: "Workflow Header list "
                }),
                PropertyPaneTextField('documentIndexListName', {
                  label: "Document Index List Name"
                }),
                PropertyPaneTextField('DocumentRevisionLog', {
                  label: "Document Revision Log list name"
                }),
                PropertyPaneTextField('notificationPrefListName', {
                  label: "Notification Preferance List Name"
                }),
                PropertyPaneTextField('emailNoficationSettings', {
                  label: "Email Nofication Settings List Name"
                }),
                PropertyPaneTextField('workflowTaskListName', {
                  label: "Work flow task listname"
                }),
                PropertyPaneTextField('taskDelegationListName', {
                  label: "Task Delegation List Name"
                }),
                PropertyPaneTextField('departmentListName', {
                  label: "Department List Name"
                }),
                PropertyPaneTextField('permissionMatrixSettings', {
                  label: "Permission Matrix Settings List Name"
                }),
                PropertyPaneTextField('accessGroupDetailsListName', {
                  label: "Access Group Details List Name"
                }),
                PropertyPaneTextField('bussinessUnitList', {
                  label: "Bussiness Unit List"
                }),
              ],
            },
            {
              groupName: "Date Color",
              groupFields: [
                PropertyPaneTextField('documentCreatedDateColor', {
                  label: "Document Created"
                }),
                PropertyPaneTextField('workflowStartedDateColor', {
                  label: "Workflow Started"
                }),
                PropertyPaneTextField('documentReviewedDateColor', {
                  label: "Reviewed"
                }),
                PropertyPaneTextField('documentApprovalDateColor', {
                  label: "Approved"
                }),
                PropertyPaneTextField('documentVoidDateColor', {
                  label: "Void Approval"
                }),
              ],
            },
            {
              groupName: "Content Color",
              groupFields: [
                PropertyPaneTextField('documentCreatedContentColor', {
                  label: "Document Created"
                }),
                PropertyPaneTextField('workflowStartedContentColor', {
                  label: "Workflow Started"
                }),
                PropertyPaneTextField('documentReviewedContentColor', {
                  label: "Reviewed"
                }),
                PropertyPaneTextField('documentApprovalContentColor', {
                  label: "Published"
                }),
                PropertyPaneTextField('documentVoidContentColor', {
                  label: "Void Approval"
                }),
                PropertyPaneTextField('internalTransittalConFab', {
                  label: "Internlly Transittted Content"
                }),
              ],
            }
          ]
        }
      ]
    };
  }
}
