import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CreateDocumentWebPartStrings';
import CreateDocument from './components/CreateDocument';
import { ICreateDocumentProps, ICreateDocumentWebPartProps } from './interfaces';

export default class CreateDocumentWebPart extends BaseClientSideWebPart<ICreateDocumentWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICreateDocumentProps> = React.createElement(
      CreateDocument,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        context: this.context,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        webpartHeader: this.properties.webpartHeader,
        documentIndexList: this.properties.documentIndexList,
        sourceDocumentLibrary: this.properties.sourceDocumentLibrary,
        businessUnit: this.properties.businessUnit,
        department: this.properties.department,
        category: this.properties.category,
        subCategory: this.properties.subCategory,
        legalEntity: this.properties.legalEntity,
        userMessageSettings: this.properties.userMessageSettings,
        publisheddocumentLibrary: this.properties.publisheddocumentLibrary,
        requestList: this.properties.requestList,
        documentIdSettings: this.properties.documentIdSettings,
        documentIdSequenceSettings: this.properties.documentIdSequenceSettings,
        documentRevisionLogList: this.properties.documentRevisionLogList,
        revisionHistoryPage: this.properties.revisionHistoryPage,
        revokePage: this.properties.revokePage,
        notificationPreference: this.properties.notificationPreference,
        emailNotification: this.properties.emailNotification,
        directPublish: this.properties.directPublish,
        QDMSUrl: this.properties.QDMSUrl,
        hubUrl: this.properties.hubUrl,
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
                PropertyPaneTextField('webpartHeader', {
                  label: 'webpartHeader'
                }),
                PropertyPaneTextField('documentIndexList', {
                  label: 'Document Index List'
                }),
                PropertyPaneTextField('sourceDocumentLibrary', {
                  label: 'Source Document Library'
                }),
                PropertyPaneTextField('businessUnit', {
                  label: 'Business Unit'
                }),
                PropertyPaneTextField('department', {
                  label: 'Department List'
                }),
                PropertyPaneTextField('category', {
                  label: 'Category List'
                }),
                PropertyPaneTextField('subCategory', {
                  label: 'Sub-Category List'
                }),
                PropertyPaneTextField('legalEntity', {
                  label: 'Legal Entity List'
                }),
                PropertyPaneTextField('userMessageSettings', {
                  label: 'User Message Settings List'
                }),
                PropertyPaneTextField('publisheddocumentLibrary', {
                  label: 'Published Document Library'
                }),
                PropertyPaneTextField('requestList', {
                  label: 'Request List'
                }),
                PropertyPaneTextField('documentIdSettings', {
                  label: 'DocumentId Settings List'
                }),
                PropertyPaneTextField('documentIdSequenceSettings', {
                  label: 'DocumentId Sequence Settings List'
                }),
                PropertyPaneTextField('documentRevisionLogList', {
                  label: 'Document RevisionLog List'
                }),
                PropertyPaneTextField('revisionHistoryPage', {
                  label: 'Revision History Page'
                }),
                PropertyPaneTextField('revokePage', {
                  label: 'Revoke Page'
                }),
                PropertyPaneTextField('notificationPreference', {
                  label: 'Notification Preference List'
                }),
                PropertyPaneTextField('emailNotification', {
                  label: 'Email Notification List'
                }),
                PropertyPaneTextField('QDMSUrl', {
                  label: 'QDMSUrl'
                }),
                PropertyPaneTextField('hubUrl', {
                  label: 'HubUrl'
                }),
                PropertyPaneCheckbox('directPublish', {
                  text: 'Direct Publish'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
