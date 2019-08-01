import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'CallAzureFunctionWebPartStrings';
import CallAzureFunction from './components/CallAzureFunction';
import { ICallAzureFunctionProps } from './components/ICallAzureFunctionProps';
import { AadHttpClient } from '@microsoft/sp-http';

export interface ICallAzureFunctionWebPartProps {
  description: string;
}

export default class CallAzureFunctionWebPart extends BaseClientSideWebPart<ICallAzureFunctionWebPartProps> {
  private aadHttpClient: AadHttpClient;
  public render(): void {
    const element: React.ReactElement<ICallAzureFunctionProps > = React.createElement(
      CallAzureFunction,
      {
        description: this.properties.description,
        aadHttpClient:this.aadHttpClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit():Promise<void>{
    this.aadHttpClient =await this.context.aadHttpClientFactory.getClient("https://cc-lob-app.azurewebsites.net");
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
