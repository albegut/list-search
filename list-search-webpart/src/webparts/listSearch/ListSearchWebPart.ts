import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption, PropertyPaneToggle, IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IListSearchProps } from './components/IListSearchProps';
import ListSearch from './components/ListSearch';
import * as strings from 'ListSearchWebPartStrings';
import { IListFieldData, IListData, IDisplayFieldData, ICompleteModalData, IRedirectData, ICustomOption } from './model/IListConfigProps';
import { IPropertyFieldSite, PropertyFieldSitePicker, } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
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
import { IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { IDynamicDataCallables } from '@microsoft/sp-dynamic-data';
import { IDynamicItem } from './model/IDynamicItem';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';



export interface IListSearchWebPartProps {
  ListName: string;
  displayFieldsCollectionData: Array<IDisplayFieldData>;
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
  onClickSelectedOption: string;
  clickEnabled: boolean;
  clickIsSimpleModal: boolean;
  clickIsCompleteModal: boolean;
  clickIsRedirect: boolean;
  clickIsDynamicData: boolean;
  completeModalFields: Array<ICompleteModalData>;
  redirectData: Array<IRedirectData>;
  onRedirectIdQuery: string;
  onClickNumberOfClicksOption: string;
}

export default class ListSearchWebPart extends BaseClientSideWebPart<IListSearchWebPartProps> implements IDynamicDataCallables {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private selectedItem: IDynamicItem;
  private sitesLists: {} = {};
  private ListsFields: {} = {};

  constructor(props) {
    super();
    this.saveSiteCollectionLists = this.saveSiteCollectionLists.bind(this);
    this.saveSiteCollectionListsFields = this.saveSiteCollectionListsFields.bind(this);
    this.setNewListFieds = this.setNewListFieds.bind(this);
    this.handleSourceSiteChange = this.handleSourceSiteChange.bind(this);

  }

  protected async onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    this.context.dynamicDataSourceManager.initializeSource(this);
    this.selectedItem = { webUrl: '', listName: '', itemId: -1 };
    return super.onInit();
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'selectedItem',
        title: 'Selected Item'
      }
    ];
  }

  public getPropertyValue(propertyId: string): IDynamicItem {
    switch (propertyId) {
      case 'selectedItem':
        return this.selectedItem;
    }

    throw new Error('Unsupported property id');
  }

  private onSelectedItem = (selectedItem: IDynamicItem): void => {
    this.selectedItem = selectedItem;
    // notify that the value has changed
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedItem');
  }

  protected async onPropertyPaneConfigurationStart() {
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
      this.saveSiteCollectionLists(sites[index], lists.map(listInfo => { return listInfo.Title; }));
    });
  }

  private async loadListsFields() {
    if (this.properties.listsCollectionData && this.properties.listsCollectionData.length > 0) {
      let siteStructure = {};
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
        });
      });

      let listData = await Promise.all(listsDataPromises);

      listData.map((fields, index) => {
        this.saveSiteCollectionListsFields(sites[index], lists[index], fields);
      });
    }
  }

  private OrderfieldsCollectionData() {
    if (this.properties.listsCollectionData && this.properties.fieldCollectionData) {
      this.properties.fieldCollectionData = this.properties.fieldCollectionData.map(element => {
        let founded = this.properties.listsCollectionData.find(source => source.ListSourceField === element.ListSourceField && source.SiteCollectionSource === element.SiteCollectionSource);
        return { Order: founded.sortIdx, ...element };
      });
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
      if (this.properties.clickEnabled) {
        this.setSelectedOnClickOption(this.properties.onClickSelectedOption);
      }
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
          listsCollectionData: this.properties.listsCollectionData.sort(),
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
          clickEnabled: this.properties.clickEnabled,
          clickIsSimpleModal: this.properties.clickIsSimpleModal,
          clickIsCompleteModal: this.properties.clickIsCompleteModal,
          clickIsRedirect: this.properties.clickIsRedirect,
          clickIsDynamicData: this.properties.clickIsDynamicData,
          completeModalFields: this.properties.completeModalFields,
          redirectData: this.properties.redirectData,
          onRedirectIdQuery: this.properties.onRedirectIdQuery,
          onSelectedItem: this.onSelectedItem.bind(this),
          oneClickOption: this.properties.onClickNumberOfClicksOption == "oneClick",
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
    return Version.parse(this.context.manifest.version);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any, sitesLists?: {}, saveSitesInfoCallback?: any) {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    switch (propertyPath) {
      case "listsCollectionData":
        {
          this.properties.completeModalFields = this.properties.completeModalFields && this.properties.completeModalFields.filter(modalField => {
            return newValue.filter(option => option.SiteCollectionSource === modalField.SiteCollectionSource && option.ListSourceField === modalField.ListSourceField).length > 0;
          });

          this.properties.redirectData = this.properties.redirectData && this.properties.redirectData.filter(modalField => {
            return newValue.filter(option => option.SiteCollectionSource === modalField.SiteCollectionSource && option.ListSourceField === modalField.ListSourceField).length > 0;
          });

          this.properties.fieldCollectionData = this.properties.fieldCollectionData && this.properties.fieldCollectionData.filter(modalField => {
            return newValue.filter(option => option.SiteCollectionSource === modalField.SiteCollectionSource && option.ListSourceField === modalField.ListSourceField).length > 0;
          });

          this.OrderfieldsCollectionData();

          break;
        }
      case "ShowListName":
        {
          if (!newValue) {
            this.properties.ListNameTitle = '';
            this.properties.displayFieldsCollectionData = this.properties.displayFieldsCollectionData.filter(field => !field.IsListTitle);
          }
          else {
            if (!this.properties.displayFieldsCollectionData.some(field => field.IsListTitle)) {
              this.properties.displayFieldsCollectionData.push({ ColumnTitle: "ListName", IsListTitle: true, IsSiteTitle: false, Searcheable: true });
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
              this.properties.displayFieldsCollectionData.push({ ColumnTitle: "Site", IsListTitle: false, IsSiteTitle: true, Searcheable: true });
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
                  saveSitesInfoCallback(site.url, lists.map(listInfo => { return listInfo.Title; }));
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
          break;
        }
      case "onClickSelectedOption":
        {
          this.properties.clickIsSimpleModal = false;
          this.properties.clickIsCompleteModal = false;
          this.properties.clickIsRedirect = false;
          this.properties.clickIsDynamicData = false;
          this.setSelectedOnClickOption(newValue);

          break;
        }
      case "clickEnabled":
        {
          if (newValue) {
            this.properties.onClickSelectedOption = "simpleModal";
            this.properties.clickIsSimpleModal = true;
            this.properties.clickIsCompleteModal = false;
            this.properties.clickIsRedirect = false;
            this.properties.clickIsDynamicData = false;
          }
          break;
        }
      case "fieldCollectionData":
        {
          if (newValue) {
            this.OrderfieldsCollectionData();
          }
          break;
        }
    }
  }

  private setSelectedOnClickOption(newValue: string) {
    switch (newValue) {
      case "simpleModal":
        {
          this.properties.clickIsSimpleModal = true;
          this.properties.redirectData = undefined;
          this.properties.completeModalFields = undefined;
          break;
        }
      case "completeModal":
        {
          this.properties.clickIsCompleteModal = true;
          this.properties.redirectData = undefined;
          break;
        }
      case "redirect":
        {
          this.properties.clickIsRedirect = true;
          this.properties.completeModalFields = undefined;
          break;
        }
      case "dynamicData":
        {
          this.properties.clickIsDynamicData = true;
          this.properties.redirectData = undefined;
          this.properties.completeModalFields = undefined;
          break;
        }
    }
  }

  private getDistinctSiteCollectionSourceOptions(): IDropdownOption[] {
    let options: IDropdownOption[] = [];
    let siteOptions = this.properties.listsCollectionData && this.properties.listsCollectionData.map(option => option.SiteCollectionSource);
    if (siteOptions) {
      siteOptions.map((item, index, array) => {
        if (array.indexOf(item) == index) {
          options.push({
            key: item,
            text: item
          });
        }
      });
    }

    return options;
  }

  private getCustomsOptions(): Array<ICustomOption> {
    return [{ Key: "SiteUrl", Option: "Site information" }, { Key: "ListName", Option: "List Name" }];
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

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
      onGetErrorMessage: (value: number) => {
        if (!value) {
          return "If pagination are enabled, items per page are required";
        }
        if (value < 0) {
          return 'Only positive values are allowed';
        }
        return '';
      }
    }) : emptyProperty;

    let cacheTimePropertyPane = this.properties.UseLocalStorage ? PropertyFieldNumber("minutesToCache", {
      key: "minutesToCache",
      label: strings.MinutesToCacheData,
      value: this.properties.minutesToCache || null,
    }) : emptyProperty;

    let onclickEventOptionPropertyPane = this.properties.clickEnabled ? PropertyPaneDropdown('onClickSelectedOption', {
      label: strings.OnClickOptionsToSelect,
      selectedKey: this.properties.onClickSelectedOption || "simpleModal",
      options: [
        {
          key: "simpleModal",
          text: strings.OnClickSimpleModalText
        },
        {
          key: "completeModal",
          text: strings.OnClickCompleteModalText
        },
        {
          key: "redirect",
          text: strings.OnClickRedirectText
        },
        {
          key: "dynamicData",
          text: strings.OnClickDynamicText
        }
      ],
    }) : emptyProperty;

    let onClickNumberOfClicksOptionPropertyPane = this.properties.clickEnabled ? PropertyPaneDropdown('onClickNumberOfClicksOption', {
      label: strings.OnClickNumberOfClickOptionsToSelect,
      selectedKey: this.properties.onClickNumberOfClicksOption || "twoClicks",
      options: [
        {
          key: "oneClick",
          text: strings.OneClickTriggerText
        },
        {
          key: "twoClicks",
          text: strings.TwoClickTriggerText
        }
      ],
    }) : emptyProperty;

    let onClickCompleteModalPropertyPane = this.properties.clickEnabled && this.properties.clickIsCompleteModal ? PropertyFieldCollectionData("completeModalFields", {
      key: "completeModalFields",
      label: strings.CompleteModalFieldSelector,
      panelHeader: strings.CompleteModalHeaderSelector,
      manageBtnLabel: strings.CompleteModalButton,
      enableSorting: true,
      value: this.properties.completeModalFields,
      fields: [
        {
          id: "SiteCollectionSource",
          title: strings.CompleteModalFieldsSiteCollection,
          type: CustomCollectionFieldType.dropdown,
          options: this.getDistinctSiteCollectionSourceOptions(),
          required: true
        },
        {
          id: "ListSourceField",
          title: strings.CompleteModalFieldsList,
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
          title: strings.CompleteModalFieldsListField,
          type: CustomCollectionFieldType.custom,
          required: true,
          onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
            if (item.SiteCollectionSource && item.ListSourceField) {
              return (
                CustomCollectionDataField.getFieldPickerByList(this.ListsFields[item.SiteCollectionSource][item.ListSourceField], field, item, onUpdate, this.getCustomsOptions())
              );
            }
          }
        },
        {
          id: "TargetField",
          title: strings.CompleteModalFieldsTargetField,
          type: CustomCollectionFieldType.string,
          required: true
        }
      ]
    }) : emptyProperty;

    let onclickRedirectPropertyPane = this.properties.clickEnabled && this.properties.clickIsRedirect ? PropertyFieldCollectionData("redirectData", {
      key: "redirectData",
      label: strings.redirectDataFieldSelector,
      panelHeader: strings.redirectDataHeaderSelector,
      manageBtnLabel: strings.redirectDataButton,
      value: this.properties.redirectData,
      fields: [
        {
          id: "SiteCollectionSource",
          title: strings.redirectDataFieldsSiteCollection,
          type: CustomCollectionFieldType.dropdown,
          options: this.getDistinctSiteCollectionSourceOptions(),
          required: true
        },
        {
          id: "ListSourceField",
          title: strings.redirectDataFieldsList,
          type: CustomCollectionFieldType.custom,
          required: true,
          onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
            return (
              CustomCollectionDataField.getListPickerBySiteOptions(this.properties.listsCollectionData, field, item, onUpdate)
            );
          }
        },
        {
          id: "Url",
          title: strings.redirectDataUrl,
          type: CustomCollectionFieldType.string,
          required: true
        }
      ]
    }) : emptyProperty;

    let onClickRedirectIdQueryParamProperyPane = this.properties.clickEnabled && this.properties.clickIsRedirect ? PropertyPaneTextField('onRedirectIdQuery', {
      label: strings.OnclickRedirectIdText,
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
                  enableSorting: true,
                  value: this.properties.listsCollectionData,
                  fields: [
                    {
                      id: "SiteCollectionSource",
                      title: strings.CollectionDataSiteCollectionTitle,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        let aa = this.properties.sites.map(site => { return site.url });
                        return (
                          CustomCollectionDataField.getPickerByStringOptions(aa, field, item, onUpdate, this.handleSourceSiteChange)
                        );
                      },
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
                            CustomCollectionDataField.getPickerByStringOptions(this.sitesLists[item.SiteCollectionSource], field, item, onUpdate, this.setNewListFieds)
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
                ItemsInPagePropertyPane,
                PropertyPaneToggle('clickEnabled', {
                  label: strings.OnClickEvent,
                }),
                onClickNumberOfClicksOptionPropertyPane,
                onclickEventOptionPropertyPane,
                onClickCompleteModalPropertyPane,
                onclickRedirectPropertyPane,
                onClickRedirectIdQueryParamProperyPane
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
                  enableSorting: true,
                  key: "displayFieldsCollectionData",
                  label: strings.CollectionDataFieldTitle,
                  panelHeader: strings.CollectionDataFieldHeader,
                  manageBtnLabel: strings.CollectionDataFieldsButton,
                  value: this.properties.displayFieldsCollectionData,
                  fields: [
                    {

                      id: "ColumnTitle",
                      title: "Column Title",
                      type: CustomCollectionFieldType.string,
                      required: true,
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
              groupName: strings.StoragePropertiesGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('UseLocalStorage', {
                  label: strings.UseLocalStorage,
                  checked: this.properties.UseLocalStorage
                }),
                cacheTimePropertyPane
              ],
            }
          ]
        },
        {
          header: {
            description: strings.InformationPropertiesGroupName
          },
          groups: [
            {
              groupName: strings.AboutPropertiesGroupName,
              groupFields: [
                PropertyPaneWebPartInformation({
                  description: `Version:  ${this.dataVersion}`,
                  key: 'webPartInfoId'
                })
              ],
            }
          ]
        }
      ]
    };

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

  private handleSourceSiteChange(row: IListData, fieldId: string, optionKey: string, updateFunction: any, errorFunction: any) {
    updateFunction(fieldId, optionKey);
    if (row) {
      let savedValue = this.properties.listsCollectionData.find(element => element.uniqueId === row.uniqueId);
      if (savedValue && savedValue.SiteCollectionSource != optionKey) {
        row.ListSourceField = "";
      }
    }
  }
}
