import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { loadTheme } from '@fluentui/react';

import * as strings from 'SignatureWebPartWebPartStrings';
import SignatureWebPart from './components/SignatureWebPart';
import type { DocumentDisplayMode, ISignatureWebPartProps } from './components/ISignatureWebPartProps';

export interface ISignatureWebPartWebPartProps {
  description: string;
  siteUrl?: string;
  taskListName?: string;
  displayMode?: DocumentDisplayMode;
}

export default class SignatureWebPartWebPart extends BaseClientSideWebPart<ISignatureWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element = React.createElement(
      SignatureWebPart as React.ComponentType<ISignatureWebPartProps>,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        siteUrl: this.properties.siteUrl,
        taskListName: this.properties.taskListName,
        displayMode: this.properties.displayMode || 'detailsList',
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient
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
      semanticColors,
      palette
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
      this.domElement.style.setProperty('--themeDarkAlt', palette.themeDarkAlt || null);
      this.domElement.style.setProperty('--neutralLight', palette.neutralLight || null);
      this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || null);
      this.domElement.style.setProperty('--neutralTertiary', palette.neutralTertiary || null);
      this.domElement.style.setProperty('--white', palette.white || null);
    }

    // Sync the site theme into Fluent's global theme so components reading getTheme() (e.g. icon colors) match it.
    loadTheme({ palette, semanticColors });
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
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel
                }),
                PropertyPaneTextField('taskListName', {
                  label: strings.TaskListNameFieldLabel
                }),
                PropertyPaneDropdown('displayMode', {
                  label: strings.DisplayModeFieldLabel,
                  selectedKey: this.properties.displayMode || 'detailsList',
                  options: [
                    { key: 'detailsList', text: strings.DisplayModeOptionDetailsList },
                    { key: 'cards', text: strings.DisplayModeOptionCards }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
