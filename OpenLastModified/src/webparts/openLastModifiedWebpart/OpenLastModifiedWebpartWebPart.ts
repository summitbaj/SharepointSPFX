import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'OpenLastModifiedWebpartWebPartStrings';
import OpenLastModifiedWebpart from './components/OpenLastModifiedWebpart';
import { IOpenLastModifiedWebpartProps } from './components/IOpenLastModifiedWebpartProps';

export interface IOpenLastModifiedWebpartWebPartProps {
  spHttpClient: SPHttpClient;
  description: string;
  imgSrc:string;
  siteURL:string;
  folderRelativeURL: string;
}

export default class OpenLastModifiedWebpartWebPart extends BaseClientSideWebPart<IOpenLastModifiedWebpartWebPartProps> {



  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  public render(): void {
    const element: React.ReactElement<IOpenLastModifiedWebpartProps> = React.createElement(
      OpenLastModifiedWebpart,
      {
        description: this.properties.description,
        imgSrc : this.properties.imgSrc,
        siteURL: this.properties.siteURL,
        folderRelativeURL: this.properties.folderRelativeURL,
        spHttpClient :this.context.spHttpClient,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
                PropertyPaneTextField('imgSrc', {
                  label: strings.ImgSrcLabel,
                  description:'Example: https://example.com/image.jpg'
                }),
                PropertyPaneTextField('siteURL', {
                  label: strings.SiteURLLabel,
                  description: 'Example: https://<your-tenant>.sharepoint.com/sites/<site-name>/'
                }),
                PropertyPaneTextField('folderRelativeURL', {
                  label: strings.FolderLocation,
                  description: 'Example: /sites/<site-name>/<folder-name>'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
