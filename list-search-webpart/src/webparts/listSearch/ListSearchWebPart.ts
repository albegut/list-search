import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
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
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import CustomCollectionDataField from './custompropertyPane/CustomCollectionDataField';
import ListService from './services/ListService';


export interface IListSearchWebPartProps {
  ListName: string;
  fieldCollectionData: Array<IListFieldData>;
  listsCollectionData: Array<IListData>;
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
  IndividualFilterPosition: string[];
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
  private sitesLists: {} = {};
  private ListsFields: {} = {};

  constructor(props) {
    super();
    this.saveSiteCollectionLists = this.saveSiteCollectionLists.bind(this);
    this.saveSiteCollectionListsFields = this.saveSiteCollectionListsFields.bind(this);
    this.setNewListFieds = this.setNewListFieds.bind(this);

  }

  protected async onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }

  async onPropertyPaneConfigurationStart() {
    await this.loadCollectionData();
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  private async loadCollectionData() {
    let sitesListsInfo: Promise<any> = this.loadSitesLists();
    let listsFieldsInfo: Promise<any> = this.loadListsFields();
    await Promise.all([sitesListsInfo, listsFieldsInfo]);
  }

  private async loadSitesLists() {
    let listsDataPromises: Promise<any>[] = [];
    let sites: string[] = [];
    this.properties.sites.map((item, index, array) => {
      if (array.indexOf(item) == index) {
        let service: ListService = new ListService(item.url);
        listsDataPromises.push(service.getSiteListsTitle());
        sites.push(item.url);
      }
    });
    let listData = await Promise.all(listsDataPromises);

    listData.map((lists, index) => {
      this.saveSiteCollectionLists(sites[index], lists.map(listInfo => { return listInfo.Title }));
    })
  }

  private async loadListsFields() {
    if (this.properties.listsCollectionData && this.properties.listsCollectionData.length > 0) {
      let siteStructure = {}
      this.properties.listsCollectionData.map(option => {
        if (!siteStructure[option.SiteCollectionSource]) {
          siteStructure[option.SiteCollectionSource] = [];
        }
        siteStructure[option.SiteCollectionSource].push(option.ListSourceField);
      });

      let listsDataPromises: Promise<any>[] = [];
      let lists: string[] = [];
      let sites: string[] = [];

      Object.keys(siteStructure).map(site => {
        let service: ListService = new ListService(site);
        siteStructure[site].map(list => {
          listsDataPromises.push(service.getListFieldsTitle(list));
          lists.push(list);
          sites.push(site);
        })
      })

      let listData = await Promise.all(listsDataPromises);

      listData.map((fields, index) => {
        this.saveSiteCollectionListsFields(sites[index], lists[index], fields.map(fieldInfo => { return fieldInfo.Title }));
      })
    }
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
      let sercheableFields = this.properties.fieldCollectionData.filter(fieldData => { if (fieldData.Searcheable) return fieldData.TargetField; });

      if (this.properties.ListNameSearcheable) {
        const listNameData: IListFieldData = { ListSourceField: "", Order: 0, Searcheable: true, SiteCollectionSource: "", SourceField: this.properties.ListNameTitle, TargetField: this.properties.ListNameTitle };
        sercheableFields.push(listNameData);
      }

      if (this.properties.SiteNameSearcheable) {
        const SiteNameData: IListFieldData = { ListSourceField: "", Order: 0, Searcheable: true, SiteCollectionSource: "", SourceField: this.properties.SiteNameTitle, TargetField: this.properties.SiteNameTitle };
        sercheableFields.push(SiteNameData);
      }

      const element: React.ReactElement<IListSearchProps> = React.createElement(
        ListSearch,
        {
          Sites: this.properties.sites,
          fieldsCollectionData: this.properties.fieldCollectionData,
          listsCollectionData: this.properties.listsCollectionData,
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
          IndividualFilterPosition: this.properties.IndividualFilterPosition,
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

  private isConfig(): boolean {
    return this.properties.sites && this.properties.sites.length > 0 && this.properties.fieldCollectionData && this.properties.fieldCollectionData.length > 0 &&
      this.properties.listsCollectionData && this.properties.listsCollectionData.length > 0;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any, sitesLists?: {}, saveSitesInfoCallback?: any) {
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
      case "sites":
        {
          if (sitesLists) {
            if (newValue && oldValue && newValue.length > 0 && oldValue.length < newValue.length) {
              await newValue.map(async site => {
                if (oldValue.indexOf(site) < 0) {
                  let service: ListService = new ListService(site.url);
                  let lists = await service.getSiteListsTitle();
                  saveSitesInfoCallback(site.url, lists.map(listInfo => { return listInfo.Title }));
                }
              });
            }
          }
          break;
        }
    }
  }

  private getDistinctSiteCollectionSourceOptions(): IDropdownOption[] {
    let options: IDropdownOption[] = [];
    let siteOptions = this.properties.listsCollectionData.map(option => option.SiteCollectionSource);
    siteOptions.map((item, index, array) => {
      if (array.indexOf(item) == index) {
        options.push({
          key: item,
          text: item
        });
      }
    });

    return options;
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    try {
      let SiteTitleOptions: IPropertyPaneDropdownOption[] = [];
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

      let ListNameSearcheablePropertyPane = this.properties.ShowListName ? PropertyPaneToggle('ListNameSearcheable', {
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

      let SiteNameSearcheablePropertyPane = this.properties.ShowSiteTitle ? PropertyPaneToggle('SiteNameSearcheable', {
        label: "Site title searcheable in general filter",
      }) : emptyProperty;

      let GeneralFilterPlaceHolderPropertyPane = this.properties.GeneralFilter ? PropertyPaneTextField('GeneralFilterPlaceHolderText', {
        label: "General filter placeholder",
      }) : emptyProperty;

      let IndividualFilterPositionPropertyPane = this.properties.IndividualColumnFilter ? PropertyFieldMultiSelect('IndividualFilterPosition', {
        key: 'multiSelect',
        label: "Multi select field",
        options: [
          {
            key: "header",
            text: "Header"
          },
          {
            key: "footer",
            text: "Footer"
          },
        ],
        selectedKeys: this.properties.IndividualFilterPosition
      }) : emptyProperty;

      let ClearAlFiltersBtnTextPropertyPane = this.properties.ShowClearAllFilters ? PropertyPaneTextField('ClearAllFiltersBtnText', {
        label: "Clear all filters text",
      }) : emptyProperty;

      let clearAllFiltersBtnColorOptions: IPropertyPaneDropdownOption[] = [];
      clearAllFiltersBtnColorOptions.push({ key: "white", text: "White" });
      clearAllFiltersBtnColorOptions.push({ key: "theme", text: "Theme" });
      let ClearAlFiltersBtnColorPropertyPane = this.properties.ShowClearAllFilters ? PropertyPaneDropdown('ClearAllFiltersBtnColor', {
        label: "Clear all filters button color",
        options: clearAllFiltersBtnColorOptions
      }) : emptyProperty;

      let ItemCountTextFieldPropertyPane = this.properties.ShowItemCount ? PropertyPaneTextField('ItemCountText', {
        label: "Item count text",
        placeholder: "Use {itemCount} to insert items count number"
      }) : emptyProperty;

      let ItemsInPagePropertyPane = this.properties.ShowPagination ? PropertyFieldNumber("ItemsInPage", {
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
                    onPropertyChange: (propertyPath, oldValue, newValue) => this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue, this.sitesLists, this.saveSiteCollectionLists),
                    properties: this.properties,
                    key: 'sitesFieldId',
                  }),
                  PropertyFieldCollectionData("listsCollectionData", {
                    key: "listsCollectionData",
                    label: "Lists data",
                    panelHeader: "Collection list data panel header",
                    manageBtnLabel: "Manage lists data",
                    value: this.properties.listsCollectionData,
                    fields: [
                      {
                        id: "SiteCollectionSource",
                        title: "Site Collection",
                        type: CustomCollectionFieldType.dropdown,
                        options: this.properties.sites && this.properties.sites.map(site => {
                          return {
                            key: site.url,
                            text: site.url
                          };
                        }),
                        required: true,
                      },
                      {
                        id: "ListSourceField",
                        title: "List",
                        type: CustomCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                          if (item.SiteCollectionSource) {
                            return (
                              CustomCollectionDataField.getListPickerBySite(this.sitesLists[item.SiteCollectionSource], field, item, onUpdate, this.setNewListFieds)
                            );
                          }
                        }
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
                  ItemCountTextFieldPropertyPane,
                  PropertyFieldNumber("ItemLimit", {
                    key: "ItemLimit",
                    label: "Item limit to show",
                    description: "If 0 all items are render",
                    value: this.properties.ItemLimit || null,
                  }),
                  PropertyPaneToggle('ShowPagination', {
                    label: "Show pagination",
                  }),
                  ItemsInPagePropertyPane
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
                  PropertyFieldCollectionData("fieldCollectionData", {
                    key: "fieldCollectionData",
                    label: "Collection data",
                    panelHeader: "Collection data panel header",
                    manageBtnLabel: "Manage collection data",
                    value: this.properties.fieldCollectionData,
                    fields: [
                      {
                        id: "SiteCollectionSource",
                        title: "Site Collection",
                        type: CustomCollectionFieldType.dropdown,
                        options: this.getDistinctSiteCollectionSourceOptions(),
                        required: true
                      },
                      {
                        id: "ListSourceField",
                        title: "List",
                        type: CustomCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                          return (
                            CustomCollectionDataField.getListPickerBySiteOptions(this.properties.listsCollectionData, field, item, onUpdate)
                          );
                        }
                      },
                      {
                        id: "SourceField",
                        title: "List Field",
                        type: CustomCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                          if (item.SiteCollectionSource && item.ListSourceField) {
                            return (
                              CustomCollectionDataField.getFieldPickerByList(this.ListsFields[item.SiteCollectionSource][item.ListSourceField], field, item, onUpdate)
                            );
                          }
                        }
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
                      },
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
                  ListNameSearcheablePropertyPane
                  ,
                  PropertyPaneToggle('ShowSiteTitle', {
                    label: "Show site information",
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowSiteTitle
                  }),
                  SiteNamePropertyToShowPropertyPane,
                  SiteNameTitlePropertyPane,
                  SiteNameOrderPropertyPane,
                  SiteNameSearcheablePropertyPane
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
                  GeneralFilterPlaceHolderPropertyPane
                  ,
                  PropertyPaneToggle('IndividualColumnFilter', {
                    label: "Indovidual column filter",
                    checked: this.properties.IndividualColumnFilter
                  }),
                  IndividualFilterPositionPropertyPane,
                  PropertyPaneToggle('ShowClearAllFilters', {
                    label: "Show button clear all filters",
                    checked: this.properties.ShowClearAllFilters
                  }),
                  ClearAlFiltersBtnColorPropertyPane,
                  ClearAlFiltersBtnTextPropertyPane
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

  private saveSiteCollectionLists(site: string, Lists: string[]) {
    this.sitesLists[site] = Lists;
  }

  private saveSiteCollectionListsFields(site: string, list: string, fields: string[]) {
    if (this.ListsFields[site] == undefined) {
      this.ListsFields[site] = {};
    }
    this.ListsFields[site][list] = fields;
  }

  private async setNewListFieds(row: IListData, fieldId: string, optionKey: string, updateFunction: any, errorFunction: any) {
    updateFunction(fieldId, optionKey);
    if (this.ListsFields[row.SiteCollectionSource] == undefined) {
      this.ListsFields[row.SiteCollectionSource] = {};
    }
    let service: ListService = new ListService(row.SiteCollectionSource);
    let fields = await service.getListFieldsTitle(optionKey);
    this.ListsFields[row.SiteCollectionSource][optionKey] = fields.map(fieldInfo => { return fieldInfo.Title });
  }
}
