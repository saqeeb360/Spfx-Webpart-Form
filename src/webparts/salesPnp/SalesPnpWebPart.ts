import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SalesPnpWebPartStrings';
import SalesPnp from './components/SalesPnp';
import { ISalesPnpProps } from './components/ISalesPnpProps';
import { sp } from '@pnp/sp/presets/all';

export interface ISalesPnpWebPartProps {
  description: string;
}

export default class SalesPnpWebPart extends BaseClientSideWebPart<ISalesPnpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISalesPnpProps> = React.createElement(
      SalesPnp,
      {
        description: this.properties.description,
        context : this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit():Promise<void>{
    console.log("onInit Called!! ");
    return super.onInit().then((_) =>{
      sp.setup({spfxContext : this.context});
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
