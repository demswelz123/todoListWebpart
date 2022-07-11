import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TodolistwebpartWebPartStrings';
import Todolistwebpart from './components/Todolistwebpart';
import { ITodolistwebpartProps } from './components/ITodolistwebpartProps';

export interface ITodolistwebpartWebPartProps {
  context: any;
  description: string;
}

export default class TodolistwebpartWebPart extends BaseClientSideWebPart<ITodolistwebpartWebPartProps> {

  public onInit() {
    return new Promise<void>((resolve, _reject) => {
        sp.setup({
            spfxContext: this.context
        });
        resolve(undefined);
    });
}
  public render(): void {
    const element: React.ReactElement<ITodolistwebpartProps > = React.createElement(
      Todolistwebpart,
      {
        description: this.properties.description
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
