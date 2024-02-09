import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'NotificationPreferenceWebPartStrings';
import NotificationPreference from './components/NotificationPreference';
import { INotificationPreferenceProps, INotificationPreferenceWebPartProps } from './interfaces';

export default class NotificationPreferenceWebPart extends BaseClientSideWebPart<INotificationPreferenceWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<INotificationPreferenceProps> = React.createElement(
      NotificationPreference,
      {
        description: this.properties.description,
        hubSiteUrl: this.properties.hubSiteUrl,
        notificationPrefListName: this.properties.notificationPrefListName,
        noEmail: this.properties.noEmail,
        sendForCriticalDocuments: this.properties.sendForCriticalDocuments,
        sendForAllDocuments: this.properties.sendForAllDocuments,
        userMessageSettings: this.properties.userMessageSettings,
        defaultPreferenceText: this.properties.defaultPreferenceText,
        noEmailText: this.properties.noEmailText,
        sendForCriticalDocumentsText: this.properties.sendForCriticalDocumentsText,
        sendForAllDocumentsText: this.properties.sendForAllDocumentsText,
        context: this.context
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
                PropertyPaneTextField('hubSiteUrl', {
                  label: "Hubsite URL:"
                }),
                PropertyPaneTextField('notificationPrefListName', {
                  label: "Notification preference settings list name:"
                }),
                //Keys are setting as per list choice values
                PropertyPaneLabel('noEmail', {
                  text: "No Email"
                }),
                PropertyPaneLabel('sendForCriticalDocuments', {
                  text: "Send mail for critical document"
                }),
                PropertyPaneLabel('sendForAllDocuments', {
                  text: "Send all emails"
                }),
                //Check box text values
                PropertyPaneTextField('noEmailText', {
                  label: "No Email text"
                }),
                PropertyPaneTextField('sendForCriticalDocumentsText', {
                  label: "Send for critical text:"
                }),
                PropertyPaneTextField('sendForAllDocumentsText', {
                  label: "Send for all text:"
                }),
                PropertyPaneTextField('userMessageSettings', {
                  label: "user Message Settings list name:"
                }),
                PropertyPaneTextField('defaultPreferenceText', {
                  label: "Inital text for default preference:"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
