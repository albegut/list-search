import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IListSearchProps } from './components/IListSearchProps';
import ListSearch from './components/ListSearch';
import * as strings from 'ListSearchWebPartStrings';
import { IListConfigProps } from './model/IListConfigProps';

export interface IListSearchWebPartProps {
  ListName: string;
  collectionData: Array<IListConfigProps>;
  ShowListName : boolean;
  ListNameTitle: string;
}

export default class ListSearchWebPart extends BaseClientSideWebPart<IListSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListSearchProps> = React.createElement(
      ListSearch,
      {
        ListName: this.properties.ListName,
        collectionData: this.properties.collectionData,
        ShowListName : this.properties.ShowListName,
        ListNameTitle: this.properties.ListNameTitle,
        Context: this.context
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
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: "ListSoruceField",
                      title: "List",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "SoruceField",
                      title: "Source field",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "TargetField",
                      title: "Target field",
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ]
                }),
                PropertyPaneToggle('ShowListName', {
                  label: strings.ListFieldLabel
                }),
                PropertyPaneTextField('ListNameTitle', {
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
