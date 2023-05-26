import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import Map from './components/Map';
import { IMapProps } from './components/Map';

export interface ICustomWebMapProps {
  listUrl: string;
  apiKey: string;
  listName: string;
}

export default class CustomWebMapWebPart extends BaseClientSideWebPart<ICustomWebMapProps> {
  public render(): void {
    const element: React.ReactElement<IMapProps> = React.createElement(Map, {
      listUrl: this.properties.listUrl,
      apiKey: this.properties.apiKey,
      listName: this.properties.listName
    });
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the SharePoint List URL, Google Maps API Key and List Name'
          },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('listUrl', {
                  label: 'SharePoint Site URL'
                }),
                PropertyPaneTextField('apiKey', {
                  label: 'Google Maps API Key'
                }),
                PropertyPaneTextField('listName', {
                  label: 'Enter List Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
