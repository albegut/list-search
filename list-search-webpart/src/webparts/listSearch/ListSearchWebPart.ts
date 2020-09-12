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
import { IListFieldData, IListData, IDisplayFieldData } from './model/IListConfigProps';
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
import { IListField } from './model/IListField';


export interface IListSearchWebPartProps {
  ListName: string;
  displayFieldsCollectionData: Array<IDisplayFieldData>
  fieldCollectionData: Array<IListFieldData>;
  listsCollectionData: Array<IListData>;
  ShowListName: boolean;
  ListNameTitle: string;
  ShowSiteTitle: boolean;
  SiteNameTitle: string;
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
  UseLocalStorage: boolean;
  minutesToCache: number;
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
          listsDataPromises.push(service.getListFields(list));
          lists.push(list);
          sites.push(site);
        })
      })

      let listData = await Promise.all(listsDataPromises);

      listData.map((fields, index) => {
        this.saveSiteCollectionListsFields(sites[index], lists[index], fields);
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
      let sercheableFields = this.properties.displayFieldsCollectionData.filter(fieldData => { if (fieldData.Searcheable) return fieldData.ColumnTitle; });

      if (this.properties.ShowListName) {
        this.properties.ListNameTitle = this.properties.displayFieldsCollectionData.filter(field => field.IsListTitle)[0].ColumnTitle;
      }

      if (this.properties.ShowSiteTitle) {
        this.properties.SiteNameTitle = this.properties.displayFieldsCollectionData.filter(field => field.IsSiteTitle)[0].ColumnTitle;
      }

      const element: React.ReactElement<IListSearchProps> = React.createElement(
        ListSearch,
        {
          Sites: this.properties.sites,
          displayFieldsCollectionData: this.properties.displayFieldsCollectionData,
          fieldsCollectionData: this.properties.fieldCollectionData,
          listsCollectionData: this.properties.listsCollectionData,
          ShowListName: this.properties.ShowListName,
          ListNameTitle: this.properties.ListNameTitle,
          ShowSite: this.properties.ShowSiteTitle,
          SiteNameTitle: this.properties.SiteNameTitle,
          SiteNamePropertyToShow: this.properties.SiteNamePropertyToShow,
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
          UseLocalStorage: this.properties.UseLocalStorage,
          minutesToCache: this.properties.minutesToCache,
        }
      );
      renderElement = element;
    }


    ReactDom.render(renderElement, this.domElement);
  }

  private isConfig(): boolean {
    return this.properties.sites && this.properties.sites.length > 0 && this.properties.fieldCollectionData && this.properties.fieldCollectionData.length > 0 &&
      this.properties.listsCollectionData && this.properties.listsCollectionData.length > 0 && this.properties.displayFieldsCollectionData && this.properties.displayFieldsCollectionData.length > 0;
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
            this.properties.displayFieldsCollectionData = this.properties.displayFieldsCollectionData.filter(field => !field.IsListTitle);
          }
          else {
            if (!this.properties.displayFieldsCollectionData.some(field => field.IsListTitle)) {
              this.properties.displayFieldsCollectionData.push({ ColumnTitle: "ListName", IsListTitle: true, IsSiteTitle: false, Order: 1, Searcheable: true })
            }
          }
          break;
        }
      case "ShowSiteTitle":
        {
          if (!newValue) {
            this.properties.SiteNameTitle = '';
            this.properties.displayFieldsCollectionData = this.properties.displayFieldsCollectionData.filter(field => !field.IsSiteTitle);
          }
          else {
            if (!this.properties.displayFieldsCollectionData.some(field => field.IsSiteTitle)) {
              this.properties.displayFieldsCollectionData.push({ ColumnTitle: "Site", IsListTitle: false, IsSiteTitle: true, Order: 1, Searcheable: true })
            }
          }
          break;
        }
      case "displayFieldsCollectionData":
        {
          if (newValue && newValue.length > 0) {
            this.properties.ShowSiteTitle = newValue.some(field => field.IsSiteTitle);
            this.properties.ShowListName = newValue.some(field => field.IsListTitle);
          }
          break;
        }
      case "sites":
        {

          if (newValue && oldValue) {
            if (newValue.length > 0 && oldValue.length < newValue.length) {
              await newValue.map(async site => {
                if (oldValue.indexOf(site) < 0) {
                  let service: ListService = new ListService(site.url);
                  let lists = await service.getSiteListsTitle();
                  saveSitesInfoCallback(site.url, lists.map(listInfo => { return listInfo.Title }));
                }
              });
            }
            else {
              let difference = oldValue.filter(x => newValue.indexOf(x) === -1);

              difference.map(site => {
                this.properties.listsCollectionData = this.properties.listsCollectionData.filter(item => item.SiteCollectionSource != site.url);
                this.properties.fieldCollectionData = this.properties.fieldCollectionData.filter(item => item.SiteCollectionSource != site.url);
              });
            }
          }
        }
        break;
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

      let SiteNamePropertyToShowPropertyPane = this.properties.ShowSiteTitle ? PropertyPaneDropdown('SiteNamePropertyToShow', {
        label: strings.GeneralFieldsPropertiesSiteProperty,
        disabled: !this.properties.ShowSiteTitle,
        options: SiteTitleOptions
      }) : emptyProperty;

      let GeneralFilterPlaceHolderPropertyPane = this.properties.GeneralFilter ? PropertyPaneTextField('GeneralFilterPlaceHolderText', {
        label: strings.FilterPropertiesGeneralFilterPlaceHolder,
      }) : emptyProperty;

      let IndividualFilterPositionPropertyPane = this.properties.IndividualColumnFilter ? PropertyFieldMultiSelect('IndividualFilterPosition', {
        key: 'multiSelect',
        label: strings.FilterPropertiesIndividualFilterPostion,
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
        label: strings.FilterPropertiesClearAllBtnText,
      }) : emptyProperty;

      let clearAllFiltersBtnColorOptions: IPropertyPaneDropdownOption[] = [];
      clearAllFiltersBtnColorOptions.push({ key: "white", text: "White" });
      clearAllFiltersBtnColorOptions.push({ key: "theme", text: "Theme" });
      let ClearAlFiltersBtnColorPropertyPane = this.properties.ShowClearAllFilters ? PropertyPaneDropdown('ClearAllFiltersBtnColor', {
        label: strings.FilterPropertiesClearAllBtnColor,
        options: clearAllFiltersBtnColorOptions
      }) : emptyProperty;

      let ItemCountTextFieldPropertyPane = this.properties.ShowItemCount ? PropertyPaneTextField('ItemCountText', {
        label: strings.GeneralPropertiesItemCountText,
        placeholder: strings.GeneralPropertiesItemCountPlaceholder
      }) : emptyProperty;

      let ItemsInPagePropertyPane = this.properties.ShowPagination ? PropertyFieldNumber("ItemsInPage", {
        key: "ItemsInPage",
        label: strings.GeneralPropertiesItemPerPage,
        value: this.properties.ItemsInPage || null,
      }) : emptyProperty;

      let cacheeTimePropertyPane = this.properties.UseLocalStorage ? PropertyFieldNumber("minutesToCache", {
        key: "minutesToCache",
        label: "SS",
        value: this.properties.minutesToCache || null,
      }) : emptyProperty;

      return {
        pages: [
          {
            header: {
              description: strings.PropertyPaneDescription
            },
            displayGroupsAsAccordion: true,
            groups: [
              {
                groupName: strings.SourceSelectorGroup,
                isCollapsed: true,
                groupFields: [
                  PropertyFieldSitePicker('sites', {
                    label: strings.SitesSelector,
                    initialSites: this.properties.sites || [],
                    context: this.context,
                    multiSelect: true,
                    onPropertyChange: (propertyPath, oldValue, newValue) => this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue, this.sitesLists, this.saveSiteCollectionLists),
                    properties: this.properties,
                    key: 'sitesFieldId',
                  }),
                  PropertyFieldCollectionData("listsCollectionData", {
                    key: "listsCollectionData",
                    label: strings.ListSelector,
                    panelHeader: strings.ListSelectorPanelHeader,
                    manageBtnLabel: strings.ListSelectorLabel,
                    value: this.properties.listsCollectionData,
                    fields: [
                      {
                        id: "SiteCollectionSource",
                        title: strings.CollectionDataSiteCollectionTitle,
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
                        title: strings.CollectionDataListTitle,
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
                        title: strings.CollectionDataListViewNameTitle,
                        type: CustomCollectionFieldType.string,

                      },
                      {
                        id: "Query",
                        title: strings.CollectionDataListCamlQueryTitle,
                        placeholder: strings.CollectionDataListCamlQueryPlaceHolder,
                        type: CustomCollectionFieldType.string,
                      }
                    ],
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                  })
                ]
              },
              {
                groupName: strings.GeneralPropertiesGroup,
                isCollapsed: true,
                groupFields: [
                  PropertyPaneToggle('ShowItemCount', {
                    label: strings.GeneralPropertiesShowItemCount,
                  }),
                  ItemCountTextFieldPropertyPane,
                  PropertyFieldNumber("ItemLimit", {
                    key: "ItemLimit",
                    label: strings.GeneralPropertiesRowLimitLabel,
                    description: strings.GeneralPropertiesRowLimitDescription,
                    value: this.properties.ItemLimit || null,
                  }),
                  PropertyPaneToggle('ShowPagination', {
                    label: strings.GeneralPropertiesShowPagination,
                  }),
                  ItemsInPagePropertyPane
                ]
              }

            ]
          },
          {
            header: {
              description: strings.FieldPropertiesGroup
            },
            displayGroupsAsAccordion: true,
            groups: [
              {
                groupName: strings.DisplayFieldsPropertiesGroup,
                isCollapsed: true,
                groupFields: [
                  PropertyPaneToggle('ShowListName', {
                    label: strings.GeneralFieldsPropertiesShowListName,
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowListName,
                  }),
                  PropertyPaneToggle('ShowSiteTitle', {
                    label: strings.GeneralFieldsPropertiesShowSiteInformation,
                    disabled: !this.properties.sites || this.properties.sites.length == 0,
                    checked: !!this.properties.sites && this.properties.sites.length > 0 && this.properties.ShowSiteTitle
                  }),
                  SiteNamePropertyToShowPropertyPane,
                  PropertyFieldCollectionData("displayFieldsCollectionData", {
                    key: "displayFieldsCollectionData",
                    label: "AA",
                    panelHeader: "BB",
                    manageBtnLabel: "CC",
                    value: this.properties.displayFieldsCollectionData,
                    fields: [
                      {
                        id: "ColumnTitle",
                        title: "Column Title",
                        type: CustomCollectionFieldType.string,
                        required: true,
                      },
                      {
                        id: "Order",
                        title: strings.CollectionDataFieldsOrder,
                        type: CustomCollectionFieldType.number,
                        required: true
                      },
                      {
                        id: "ColumnWidth",
                        title: "Column Width",
                        type: CustomCollectionFieldType.number,
                      },
                      {
                        id: "IsSiteTitle",
                        title: "IsSiteTitleColumn",
                        type: CustomCollectionFieldType.boolean,
                        disableEdit: true,
                      },
                      {
                        id: "IsListTitle",
                        title: "IsListTitleColumn",
                        type: CustomCollectionFieldType.boolean,
                        disableEdit: true,
                      },
                      {
                        id: "Searcheable",
                        title: strings.CollectionDataFieldsSearchable,
                        type: CustomCollectionFieldType.boolean,
                        defaultValue: true
                      }
                    ],
                  })
                ]
              },
              {
                groupName: strings.CollectionDataFieldsProperties,
                isCollapsed: true,
                groupFields: [
                  PropertyFieldCollectionData("fieldCollectionData", {
                    key: "fieldCollectionData",
                    label: strings.CollectionDataFieldsToRetreive,
                    panelHeader: strings.CollectionDataFieldsHeader,
                    manageBtnLabel: strings.CollectionDataFieldsSelectBtn,
                    value: this.properties.fieldCollectionData,
                    fields: [
                      {
                        id: "SiteCollectionSource",
                        title: strings.CollectionDataFieldsSiteCollection,
                        type: CustomCollectionFieldType.dropdown,
                        options: this.getDistinctSiteCollectionSourceOptions(),
                        required: true
                      },
                      {
                        id: "ListSourceField",
                        title: strings.CollectionDataFieldsList,
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
                        title: strings.CollectionDataFieldsListField,
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
                        title: strings.CollectionDataFieldsTargetField,
                        type: CustomCollectionFieldType.dropdown,
                        options: this.properties.displayFieldsCollectionData && this.properties.displayFieldsCollectionData.filter(field => !field.IsListTitle && !field.IsSiteTitle).map(field => {
                            return {
                            key: field.ColumnTitle,
                            text: field.ColumnTitle
                          };

                        }),
                        required: true
                      }
                    ],
                    disabled: !this.properties.sites || this.properties.sites.length == 0 || !this.properties.displayFieldsCollectionData || this.properties.displayFieldsCollectionData.length == 0,

                  })]
              }
            ]
          },
          {
            header: {
              description: strings.FilterPropertiesGroup
            },
            displayGroupsAsAccordion: true,
            groups: [
              {
                groupName: strings.FilterPropertiesGroupName,
                isCollapsed: true,
                groupFields: [
                  PropertyPaneToggle('GeneralFilter', {
                    label: strings.FilterPropertiesGeneralFilter,
                    checked: this.properties.GeneralFilter
                  }),
                  GeneralFilterPlaceHolderPropertyPane
                  ,
                  PropertyPaneToggle('IndividualColumnFilter', {
                    label: strings.FilterPropertiesIndividualFilter,
                    checked: this.properties.IndividualColumnFilter
                  }),
                  IndividualFilterPositionPropertyPane,
                  PropertyPaneToggle('ShowClearAllFilters', {
                    label: strings.FilterPropertiesClearAllBtn,
                    checked: this.properties.ShowClearAllFilters
                  }),
                  ClearAlFiltersBtnColorPropertyPane,
                  ClearAlFiltersBtnTextPropertyPane
                ],
              },
              {
                groupName: "DD",
                isCollapsed: true,
                groupFields: [
                  PropertyPaneToggle('UseLocalStorage', {
                    label: "RR",
                    checked: this.properties.UseLocalStorage
                  }),
                  cacheeTimePropertyPane
                ],
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

  private saveSiteCollectionListsFields(site: string, list: string, fields: IListField[]) {
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
    let fields: IListField[] = await service.getListFields(optionKey);
    this.ListsFields[row.SiteCollectionSource][optionKey] = fields;
  }
}
