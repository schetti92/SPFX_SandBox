import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WpspfxCurdWebPartStrings';
import WpspfxCurd from './components/WpspfxCurd';
import { IWpspfxCurdProps } from './components/IWpspfxCurdProps';

import { getSP } from './pnpjsConfig';
//import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IWpspfxCurdWebPartProps {
  listName: string;
}

export default class WpspfxCurdWebPart extends BaseClientSideWebPart<IWpspfxCurdWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWpspfxCurdProps> = React.createElement(
      WpspfxCurd,
      {
        listName: this.properties.listName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
