import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IListSearchProps } from './components/IListSearchProps';
import ListSearch from './components/ListSearch';
import * as strings from 'ListSearchWebPartStrings';
import { IListFieldData, IListData } from './model/IListConfigProps';
import { PropertyFieldSitePicker, IPropertyFieldSite, } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';
import { EmptyPropertyPane } from './custompropertyPane/EmptyPropertyPane';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

export interface IListSearchWebPartProps {
  ListName: string;
  collectionData: Array<IListFieldData>;
  ListscollectionData: Array<IListData>;
  ShowListName: boolean;
  ListNameTitle: string;
  ListNameOrder: number;
  ListNameSearcheable: boolean;
  ShowSiteTitle: boolean;
  SiteNameTitle: string;
  SiteNameOrder: number;
  SiteNamePropertyToShow: string;
  sites: IPropertyFieldSite[];
  GeneralFilter: boolean;
  GeneralFilterPlaceHolderText: string;
  IndividualColumnFilter: boolean;
  ShowClearAllFilters: boolean;
  ClearAllFiltersBtnColor: string;
  ClearAllFiltersBtnText: string;
  SiteNameSearcheable: boolean;
  ShowItemCount: boolean;
  ItemCountText: string;
  ItemLimit: number;
  ShowPagination: boolean;
  ItemsInPage: number;

}

export default class ListSearchWebPart extends BaseClientSideWebPart<IListSearchWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

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
      let sercheableFields = this.properties.collectionData.filter(fieldData => { if (fieldData.Searcheable) return fieldData.TargetField })

      if (this.properties.ListNameSearcheable) {
        const listNameData: IListFieldData = { ListSourceField: "", Order: 0, Searcheable: true, SiteCollectionSource: "", SourceField: this.properties.ListNameTitle, TargetField: this.properties.ListNameTitle };
        sercheableFields.push(listNameData);
      }

      if (this.properties.SiteNameSearcheable) {
        const SiteNameData: IListFieldData = { ListSourceField: "", Order: 0, Searcheable: true, SiteCollectionSource: "", SourceField: this.properties.SiteNameTitle, TargetField: this.properties.SiteNameTitle };
        sercheableFields.push(SiteNameData)
      }

      const element: React.ReactElement<IListSearchProps> = React.createElement(
        ListSearch,
        {
          Sites: this.properties.sites,
          collectionData: this.properties.collectionData,
          ListscollectionData: this.properties.ListscollectionData,
          ShowListName: this.properties.ShowListName,
          ListNameTitle: this.properties.ListNameTitle,
          ListNameOrder: this.properties.ListNameOrder,
          ShowSite: this.properties.ShowSiteTitle,
          SiteNameTitle: this.properties.SiteNameTitle,
          SiteNameOrder: this.properties.SiteNameOrder,
          SiteNamePropertyToShow: this.properties.SiteNamePropertyToShow,
          SiteNameSearcheable: this.properties.SiteNameSearcheable,
          Context: this.context,
          GeneralFilter: this.properties.GeneralFilter,
          GeneralFilterPlaceHolderText: this.properties.GeneralFilterPlaceHolderText,
          ShowClearAllFilters: this.properties.ShowClearAllFilters,
          ClearAllFiltersBtnColor: this.properties.ClearAllFiltersBtnColor,
          ClearAllFiltersBtnText: this.properties.ClearAllFiltersBtnText,
          GeneralSearcheableFields: sercheableFields,
          IndividualColumnFilter: this.properties.IndividualColumnFilter,
          ShowItemCount: this.properties.ShowItemCount,
          ItemCountText: this.properties.ItemCountText,
          ItemLimit: this.properties.ItemLimit,
          ShowPagination: this.properties.ShowPagination,
          ItemsInPage: this.properties.ItemsInPage,
          themeVariant: this._themeVariant,
        }
      );
      renderElement = element;
    }


    ReactDom.render(renderElement, this.domElement);
  }

  private getSite

  private isConfig(): boolean {
    return this.properties.sites && this.properties.collectionData && this.properties.collectionData.length > 0 &&
      this.properties.ListscollectionData && this.properties.ListscollectionData.length > 0;
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
      let SiteTitleOptions: IPropertyPaneDropdownOption[] = []
      SiteTitleOptions.push({ key: "id", text: "Id" });
      SiteTitleOptions.push({ key: "title", text: "Title" });
      SiteTitleOptions.push({ key: "url", text: "Url" });

      let emptyProperty = new EmptyPropertyPane();

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

      let ListNameSearcheable = this.properties.ShowListName ? PropertyPaneToggle('ListNameSearcheable', {
        label: "List title searcheable in general filter",
      }) : emptyProperty;

      let SiteNameTitlePropertyPane = this.properties.ShowSiteTitle ? PropertyPaneTextField('SiteNameTitle', {
        label: "Site column title",
        disabled: !this.properties.ShowSiteTitle
      }) : emptyProperty;

      let SiteNamePropertyToShowPropertyPane = this.properties.ShowSiteTitle ? PropertyPaneDropdown('SiteNamePropertyToShow', {
        label: "Property to show",
        disabled: !this.properties.ShowSiteTitle,
        options: SiteTitleOptions
      }) : emptyProperty;

      let SiteNameOrderPropertyPane = this.properties.ShowSiteTitle ? PropertyFieldNumber("SiteNameOrder", {
        key: "SiteNameOrder",
        label: "Site column Order",
        description: "Order of site title column",
        value: this.properties.SiteNameOrder || null,
        disabled: !this.properties.ShowSiteTitle
      }) : emptyProperty;

      let SiteNameSearcheable = this.properties.ShowSiteTitle ? PropertyPaneToggle('SiteNameSearcheable', {
        label: "Site title searcheable in general filter",
      }) : emptyProperty;

      let GeneralFilterPlaceHolder = this.properties.GeneralFilter ? PropertyPaneTextField('GeneralFilterPlaceHolderText', {
        label: "General filter placeholder",
      }) : emptyProperty;

      let ClearAlFiltersBtnText = this.properties.ShowClearAllFilters ? PropertyPaneTextField('ClearAllFiltersBtnText', {
        label: "Clear all filters text",
      }) : emptyProperty;

      let clearAllFiltersBtnColorOptions: IPropertyPaneDropdownOption[] = []
      clearAllFiltersBtnColorOptions.push({ key: "white", text: "White" });
      clearAllFiltersBtnColorOptions.push({ key: "theme", text: "Theme" });
      let ClearAlFiltersBtnColor = this.properties.ShowClearAllFilters ? PropertyPaneDropdown('ClearAllFiltersBtnColor', {
        label: "Clear all filters button color",
        options: clearAllFiltersBtnColorOptions
      }) : emptyProperty;

      let ItemCountTextField = this.properties.ShowItemCount ? PropertyPaneTextField('ItemCountText', {
        label: "Item count text",
        placeholder: "Use {itemCount} to insert items count number"
      }) : emptyProperty;

      let ItemsInPage = this.properties.ShowPagination ? PropertyFieldNumber("ItemsInPage", {
        key: "ItemsInPage",
        label: "Item elements in page",
        value: this.properties.ItemsInPage || null,
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
                  PropertyFieldCollectionData("ListscollectionData", {
                    key: "ListscollectionData",
                    label: "Lists data",
                    panelHeader: "Collection list data panel header",
                    manageBtnLabel: "Manage lists data",
                    value: this.properties.ListscollectionData,
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
                        required: true,
                      },
                      {
                        id: "ListSourceField",
                        title: "List",
                        type: CustomCollectionFieldType.string,
                        required: true
                      },
                      {
                        id: "ListView",
                        title: "List view name",
                        type: CustomCollectionFieldType.string,

                      },
                      {
                        id: "Query",
                        title: "Custom CAML query - empty all elements",
                        type: CustomCollectionFieldType.string,
                      }
                    ],
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                  }),
                  PropertyPaneToggle('ShowItemCount', {
                    label: "Show item count",
                  }),
                  ItemCountTextField,
                  PropertyFieldNumber("ItemLimit", {
                    key: "ItemLimit",
                    label: "Item limit to show",
                    description: "If 0 all items are render",
                    value: this.properties.ItemLimit || null,
                  }),
                  PropertyPaneToggle('ShowPagination', {
                    label: "Show pagination",
                  }),
                  ItemsInPage
                ]
              }
            ]
          },
          {
            header: {
              description: "Field Properties"
            },
            groups: [
              {
                groupName: "Field Properties",
                groupFields: [
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
                        options: this.properties.ListscollectionData && this.properties.ListscollectionData.map(site => {
                          return {
                            key: site.SiteCollectionSource,
                            text: site.SiteCollectionSource
                          }
                        }),
                        required: true
                      },
                      {
                        id: "ListSourceField",
                        title: "List",
                        type: CustomCollectionFieldType.dropdown,
                        options: this.properties.ListscollectionData && this.properties.ListscollectionData.map(site => {
                          return {
                            key: site.ListSourceField,
                            text: site.ListSourceField
                          }
                        }),
                        required: true
                      },
                      {
                        id: "SourceField",
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
                      ,
                      {
                        id: "Searcheable",
                        title: "Searcheable in general filter",
                        type: CustomCollectionFieldType.boolean,
                        defaultValue: true
                      }
                    ],
                    disabled: !this.properties.sites || this.properties.sites.length == 0,

                  })]
              },
              {
                groupName: "Additional Properties",
                groupFields: [
                  PropertyPaneToggle('ShowListName', {
                    label: strings.ListFieldLabel,
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowListName
                  }),
                  ListNameTitlePropertyPane,
                  ListNameOrderPropertyPane,
                  ListNameSearcheable
                  ,
                  PropertyPaneToggle('ShowSiteTitle', {
                    label: "Show site information",
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowSiteTitle
                  }),
                  SiteNamePropertyToShowPropertyPane,
                  SiteNameTitlePropertyPane,
                  SiteNameOrderPropertyPane,
                  SiteNameSearcheable
                ]
              }
            ]
          },
          {
            header: {
              description: "Filter Options"
            },
            groups: [
              {
                groupName: "Field Properties",
                groupFields: [
                  PropertyPaneToggle('GeneralFilter', {
                    label: "General Filter",
                    checked: this.properties.GeneralFilter
                  }),
                  GeneralFilterPlaceHolder
                  ,
                  PropertyPaneToggle('IndividualColumnFilter', {
                    label: "Indovidual column filter",
                    checked: this.properties.IndividualColumnFilter
                  }),
                  PropertyPaneToggle('ShowClearAllFilters', {
                    label: "Show button clear all filters",
                    checked: this.properties.ShowClearAllFilters
                  }),
                  ClearAlFiltersBtnColor,
                  ClearAlFiltersBtnText
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
