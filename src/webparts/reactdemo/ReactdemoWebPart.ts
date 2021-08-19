import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactdemoWebPartStrings';
import Reactdemo from './components/Reactdemo';
import { IReactdemoProps } from './components/IReactdemoProps';

export interface IReactdemoWebPartProps {
  description: string;
}

export default class ReactdemoWebPart extends BaseClientSideWebPart<IReactdemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactdemoProps> = React.createElement(
      Reactdemo,
      {
        description: this.properties.description,
        webURL:this.context.pageContext.web.absoluteUrl,
        context:this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
