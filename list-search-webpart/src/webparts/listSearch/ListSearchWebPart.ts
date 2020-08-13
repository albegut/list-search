import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {IListSearchProps} from './components/IListSearchProps';
import ListSearch from './components/ListSearch';


import * as strings from 'ListSearchWebPartStrings';

export interface IListSearchWebPartProps {
  ListName: string;
}

export default class ListSearchWebPart extends BaseClientSideWebPart<IListSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListSearchProps> = React.createElement(
      ListSearch,
      {
        ListName: this.properties.ListName,
      }
    );
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: strings.ListFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
