import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IListSearchProps } from './components/IListSearchProps';
import ListSearch from './components/ListSearch';
import * as strings from 'ListSearchWebPartStrings';
import { IListConfigProps } from './model/IListConfigProps';
import { PropertyFieldSitePicker, IPropertyFieldSite, } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IListSearchWebPartProps {
  ListName: string;
  collectionData: Array<IListConfigProps>;
  ShowListName: boolean;
  ListNameTitle: string;
  ListNameOrder: number;
  ShowSiteTitle: boolean;
  SiteNameTitle: string;
  SiteNameOrder: number;
  sites: IPropertyFieldSite[];
}

export default class ListSearchWebPart extends BaseClientSideWebPart<IListSearchWebPartProps> {

  public render(): void {
    let renderElement = null;

    let isEditMode: boolean = this.displayMode === DisplayMode.Edit;
    if (!this.isConfig()) {
      const placeholder: React.ReactElement<any> = React.createElement(
        Placeholder,
        {
          iconName: 'Edit',
          iconText: 'Configure List Search webpart properties',
          description: 'You need to complete the configuration of the webpart',
          buttonLabel: 'Configure',
          onConfigure: () => this.context.propertyPane.open(),
          hideButton: !isEditMode,
        }
      );
      renderElement = placeholder;
    }
    else {
      const element: React.ReactElement<IListSearchProps> = React.createElement(
        ListSearch,
        {
          collectionData: this.properties.collectionData,
          ShowListName: this.properties.ShowListName,
          ListNameTitle: this.properties.ListNameTitle,
          ListNameOrder: this.properties.ListNameOrder,
          ShowSite: this.properties.ShowSiteTitle,
          SiteNameTitle: this.properties.SiteNameTitle,
          SiteNameOrder: this.properties.SiteNameOrder,
          Context: this.context
        }
      );
      renderElement = element;
    }


    ReactDom.render(renderElement, this.domElement);
  }

  private isConfig(): boolean {
    return this.properties.sites && this.properties.collectionData && this.properties.collectionData.length > 0;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    switch (propertyPath) {
      case "ShowListName":
        {
          if (!newValue) {
            this.properties.ListNameTitle = '';
          }
          break;
        }
      case "ShowSite":
        {
          if (!newValue) {
            this.properties.SiteNameTitle = '';
          }
          break;
        }
    }
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    try {
      let numSites: number = 0;
      let emptyProperty = PropertyPaneLabel('emptyLabel', {
        text: ""
      });

      let ListNameTitlePropertyPane = this.properties.ShowListName ? PropertyPaneTextField('ListNameTitle', {
        label: strings.ListFieldLabel,
        disabled: !this.properties.ShowListName
      }) : emptyProperty;

      let ListNameOrderPropertyPane = this.properties.ShowListName ? PropertyFieldNumber("ListNameOrder", {
        key: "ListNameOrder",
        label: "List column order",
        minValue: 0,
        description: "Order of List Name column",
        value: this.properties.ListNameOrder || null,
        disabled: !this.properties.ShowListName
      }) : emptyProperty;

      let SiteNameTitlePropertyPane = this.properties.ShowSiteTitle ? PropertyPaneTextField('SiteNameTitle', {
        label: "Site column title",
        disabled: !this.properties.ShowSiteTitle
      }) : emptyProperty;

      let SiteNameOrderPropertyPane = this.properties.ShowSiteTitle ? PropertyFieldNumber("SiteNameOrder", {
        key: "SiteNameOrder",
        label: "Site column Order",
        description: "Order of site title column",
        value: this.properties.SiteNameOrder || null,
        disabled: !this.properties.ShowSiteTitle
      }) : emptyProperty;

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
                  PropertyFieldSitePicker('sites', {
                    label: 'Select sites',
                    initialSites: this.properties.sites || [],
                    context: this.context,
                    multiSelect: true,
                    onPropertyChange: this.onPropertyPaneFieldChanged,
                    properties: this.properties,
                    key: 'sitesFieldId',
                  }),
                  PropertyFieldCollectionData("collectionData", {
                    key: "collectionData",
                    label: "Collection data",
                    panelHeader: "Collection data panel header",
                    manageBtnLabel: "Manage collection data",
                    value: this.properties.collectionData,
                    fields: [
                      {
                        id: "SiteCollectionSource",
                        title: "Site Collection",
                        type: CustomCollectionFieldType.dropdown,
                        options: this.properties.sites && this.properties.sites.map(site => {
                          return {
                            key: site.url,
                            text: site.url
                          }
                        }),
                        required: true
                      },
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
                      },
                      {
                        id: "Order",
                        title: "Order",
                        type: CustomCollectionFieldType.number,
                        required: true
                      }
                    ],
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                  }),
                  PropertyPaneToggle('ShowListName', {
                    label: strings.ListFieldLabel,
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowListName
                  }),
                  ListNameTitlePropertyPane
                  ,
                  ListNameOrderPropertyPane
                  ,
                  PropertyPaneToggle('ShowSiteTitle', {
                    label: "Show site information",
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowSiteTitle
                  }),
                  SiteNameTitlePropertyPane
                  ,
                  SiteNameOrderPropertyPane
                ]
              }
            ]
          }
        ]
      };
    }
    catch (error) {
      console.error(error);
    }
  }
}
