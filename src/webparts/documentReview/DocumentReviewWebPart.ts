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

import * as strings from 'DocumentReviewWebPartStrings';
import DocumentReview from './components/DocumentReview';
import { IDocumentReviewProps, IDocumentReviewWebPartProps } from './interfaces';

export default class DocumentReviewWebPart extends BaseClientSideWebPart<IDocumentReviewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDocumentReviewProps> = React.createElement(
      DocumentReview,
      {
        description: this.properties.description,
        webPartName: this.properties.webPartName,
        project: this.properties.project,
        redirectUrl: this.properties.redirectUrl,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        workflowHeaderListName: this.properties.workflowHeaderListName,
        context: this.context,
        notificationPrefListName: this.properties.notificationPrefListName,
        hubSiteUrl: this.properties.hubSiteUrl,
        hubsite: this.properties.hubsite,
        emailNotificationSettings: this.properties.emailNotificationSettings,
        userMessageSettings: this.properties.userMessageSettings,
        documentIndex: this.properties.documentIndex,
        workFlowDetail: this.properties.workFlowDetail,
        documentApprovalSitePage: this.properties.documentApprovalSitePage,
        documentRevisionLog: this.properties.documentRevisionLog,
        documentReviewSitePage: this.properties.documentReviewSitePage,
        workflowTaskListName: this.properties.workflowTaskListName,
        taskDelegationSettingsListName: this.properties.taskDelegationSettingsListName,
        accessGroups: this.properties.accessGroups,
        departmentList: this.properties.departmentList,
        accessGroupDetailsList: this.properties.accessGroupDetailsList,
        projectInformationListName: this.properties.projectInformationListName,
        bussinessUnitList: this.properties.bussinessUnitList,
        masterListName: this.properties.masterListName,
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
                PropertyPaneToggle('project', {
                  label: 'Project',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('webPartName', {
                  label: "WebPart Name"
                }),
                PropertyPaneTextField('workflowHeaderListName', {
                  label: "Workflow Header List Name"
                }),
                PropertyPaneTextField('notificationPrefListName', {
                  label: "Notification Preference List Name"
                }),
                PropertyPaneTextField('workflowTaskListName', {
                  label: "workflow Task List Name"
                }),
                PropertyPaneTextField('taskDelegationSettingsListName', {
                  label: "Task Delegation Settings List Name"
                }),
                PropertyPaneTextField('emailNotificationSettings', {
                  label: "Email Notification Settings List Name"
                }),
                PropertyPaneTextField('projectInformationListName', {
                  label: "Project Information List Name"
                }),
                PropertyPaneTextField('userMessageSettings', {
                  label: "User Message Settings List Name"
                }),
                PropertyPaneTextField('documentIndex', {
                  label: "Document Index List Name"
                }),
                PropertyPaneTextField('workFlowDetail', {
                  label: "Work Flow Detail List Name"
                }),
                PropertyPaneTextField('documentRevisionLog', {
                  label: "Document RevisionLogList Name"
                }),
                PropertyPaneTextField('hubsite', {
                  label: "Hub Site Name"
                }),
                PropertyPaneTextField('masterListName', {
                  label: "Master List Name"
                }),
                PropertyPaneTextField('documentApprovalSitePage', {
                  label: "Document Approval SitePage"
                }), PropertyPaneTextField('documentReviewSitePage', {
                  label: "Document Review SitePage"
                }),
                PropertyPaneTextField('accessGroups', {
                  label: "Access Groups"
                }),
                PropertyPaneTextField('departmentList', {
                  label: "Department List"
                }), PropertyPaneTextField('accessGroupDetailsList', {
                  label: "Access Group Details List"
                }), PropertyPaneTextField('bussinessUnitList', {
                  label: "Bussiness Unit List"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
