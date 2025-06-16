// src/webparts/xyea/XyeaWebPart.ts

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'XyeaWebPartStrings';
import Xyea from './components/Xyea';
import { IXyeaProps } from './components/IXyeaProps';

export interface IXyeaWebPartProps {
  description: string;
}

export default class XyeaWebPart extends BaseClientSideWebPart<IXyeaWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IXyeaProps> = React.createElement(
      Xyea,
      {
        context: this.context,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
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