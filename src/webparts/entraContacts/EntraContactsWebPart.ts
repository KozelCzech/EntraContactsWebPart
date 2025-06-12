import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EntraContactsWebPartStrings';
import EntraContacts from './components/EntraContacts';
import { IEntraContactsProps } from './components/IEntraContactsProps';

import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IEntraContactsWebPartProps {
  description: string;
}

export default class EntraContactsWebPart extends BaseClientSideWebPart<IEntraContactsWebPartProps> {

  private _graphClient: MSGraphClientV3;


  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      this.context.msGraphClientFactory.getClient('3').then((client: MSGraphClientV3) => {
        this._graphClient = client;
        resolve();
      }).catch((error) => {
        reject(error);
      });
    });  
  }


  public render(): void {
    const element: React.ReactElement<IEntraContactsProps> = React.createElement(
      EntraContacts,
      {
        graphClient: this._graphClient,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
