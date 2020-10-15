import * as React from 'react';
import styles from '../ListSearchWebPart.module.scss';
import * as strings from 'ListSearchWebPartStrings';
import ListService from '../services/ListService';
import { IListSearchState, IColumnFilter } from './IListSearchState';
import { IListSearchProps } from './IListSearchProps';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import {
  DetailsList,
  IColumn,
  IDetailsFooterProps,
  IDetailsRowBaseProps,
  DetailsRow,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { IListSearchListQuery } from '../model/ListSearchQuery';
import {
  getTheme,
  IconButton,
  MessageBar,
  MessageBarType
} from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import Pagination from "react-js-pagination";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { Icon, ITheme } from 'office-ui-fabric-react';
import SessionStorage from '../services/SessionStorageService';
import { ISessionStorageElement } from '../model/ISessiongStorageElement';
import { Shimmer } from 'office-ui-fabric-react/lib/Shimmer';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { Log } from '@microsoft/sp-core-library';
import { IBaseFieldData } from '../model/IListConfigProps';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { SharePointType } from '../model/ISharePointFieldTypes';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Facepile, OverflowButtonType, IFacepilePersona } from 'office-ui-fabric-react/lib/Facepile';
import StringUtils from '../services/Utils';
import { Image, IImageProps, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Link } from 'office-ui-fabric-react';




const LOG_SOURCE = "IListdSearchWebPart";
const filterIcon: IIconProps = { iconName: 'Filter' };

export default class IListdSearchWebPart extends React.Component<IListSearchProps, IListSearchState> {
  private columns: IColumn[] = [];
  private keymapQuerys: {} = {};
  constructor(props: IListSearchProps, state: IListSearchState) {
    super(props);
    this.state = {
      activePage: 1,
      items: null,
      filterItems: null,
      isLoading: true,
      errorMsg: "",
      columnFilters: [],
      generalFilter: "",
      isModalHidden: true,
      isModalLoading: false,
      selectedItem: null,
      completeModalItemData: null
    };
    this.GetJSXElementByType = this.GetJSXElementByType.bind(this);
    this._renderItemColumn = this._renderItemColumn.bind(this);

  }

  public componentDidUpdate(prevProps: Readonly<IListSearchProps>, prevState: Readonly<IListSearchState>, snapshot?: any): void {
    if (prevProps != this.props) {
      this.columns = [];
      this.setState({ items: null, filterItems: null, isLoading: true });
      this.getData();
    }
  }

  public componentDidMount() {
    this.getData();
  }

  private SetError(error: Error, methodName: string) {
    Log.warn(LOG_SOURCE, `${methodName} set an error`, this.props.Context.serviceScope);
    Log.error(LOG_SOURCE, error, this.props.Context.serviceScope);
    this.setState({
      errorMsg: `Error ${error.message}`,
      isLoading: false,
    });
  }

  private async getData() {
    try {
      let result: any[] = [];
      if (!this.props.UseLocalStorage) {
        result = await this.readListsItems();
      }
      else {
        let session: SessionStorage = new SessionStorage();
        let cacheData: ISessionStorageElement = session.getSotareElementByKey("sharepointData");
        if (cacheData.hasExpired || !cacheData.elements) {
          result = await this.readListsItems();
          session.setSotareElementByKey("sharepointData", result, this.props.minutesToCache);
        }
        else {
          this.AddColumnsToDisplay();
          result = cacheData.elements;
        }
      }

      if (this.props.ItemLimit) {
        result = result.slice(0, this.props.ItemLimit);
      }

      this.setState({ items: result, filterItems: result, isLoading: false });
    } catch (error) {
      this.SetError(error, "getData");
    }
  }

  private async readListsItems(): Promise<Array<any>> {
    this.generateKeymap();
    let itemPromise: Array<Promise<Array<any>>> = [];

    //TODO check if it is necessary to generate the columns
    this.AddColumnsToDisplay();

    Object.keys(this.keymapQuerys).map(site => {
      let listService: ListService = new ListService(site);
      let siteProperties = this.props.Sites.filter(siteInformation => siteInformation.url === site);
      Object.keys(this.keymapQuerys[site]).map(listQuery => {
        itemPromise.push(listService.getListItems(this.keymapQuerys[site][listQuery], this.props.ListNameTitle, this.props.SiteNameTitle, siteProperties[0][this.props.SiteNamePropertyToShow], this.props.ItemLimit));
      });
    });

    let items = await Promise.all(itemPromise);
    let result = [];
    items.map(partialResult => {
      result.push(...partialResult);
    });

    return result;
  }

  private AddColumnsToDisplay(): void {
    this.props.displayFieldsCollectionData.sort().map(column => {
      this.columns.push({ key: column.ColumnTitle, name: column.ColumnTitle, fieldName: column.ColumnTitle, minWidth: 100, maxWidth: column.ColumnWidth || undefined, isResizable: true, data: column.SPFieldType });
    });
  }

  private generateKeymap() {
    this.keymapQuerys = {};
    this.props.fieldsCollectionData.map(item => {
      if (this.keymapQuerys[item.SiteCollectionSource] != undefined) {
        if (this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] != undefined) {
          if (this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField].fields.filter(field => field.originalField === item.SourceField).length == 0) {
            this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField].fields.push({ originalField: item.SourceField, newField: item.TargetField, fieldType: item.SPFieldType });
          }
        }
        else {
          let listQueryInfo = this.props.listsCollectionData.filter(list => list.SiteCollectionSource == item.SiteCollectionSource && list.ListSourceField == item.ListSourceField);

          let newQueryListItem: IListSearchListQuery = { list: item.ListSourceField, fields: [{ originalField: item.SourceField, newField: item.TargetField, fieldType: item.SPFieldType }], camlQuery: listQueryInfo.length > 0 && listQueryInfo[0].Query, viewName: listQueryInfo.length > 0 && listQueryInfo[0].ListView };
          this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] = newQueryListItem;
        }
      }
      else {
        let listQueryInfo = this.props.listsCollectionData.filter(list => list.SiteCollectionSource == item.SiteCollectionSource && list.ListSourceField == item.ListSourceField);

        let newQueryListItem: IListSearchListQuery = { list: item.ListSourceField, fields: [{ originalField: item.SourceField, newField: item.TargetField, fieldType: item.SPFieldType }], camlQuery: listQueryInfo.length > 0 && listQueryInfo[0].Query, viewName: listQueryInfo.length > 0 && listQueryInfo[0].ListView };
        this.keymapQuerys[item.SiteCollectionSource] = {};
        this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] = newQueryListItem;
      }
    });
  }

  public filterColumnListItems(propertyName: string, propertyValue: string) {
    try {
      let isNewFilter: boolean = true;
      let clearFilter: boolean = false;
      let isMoreRestricted: boolean = false;
      let newFitlers: IColumnFilter[] = this.state.columnFilters.filter(filter => {
        if (filter.columnName === propertyName) {
          isMoreRestricted = filter.filterToApply.length < propertyValue.length;
          filter.filterToApply = propertyValue;
          isNewFilter = false;

        }
        if (filter.filterToApply && filter.filterToApply.length > 0) { //Remove empty filters
          return filter;
        }
        else {
          clearFilter = true;
        }
      });

      if (isNewFilter) newFitlers.push({ columnName: propertyName, filterToApply: propertyValue });

      let itemsToRefine = (clearFilter || this.state.generalFilter) ? this.filterListItemsByGeneralFilter(this.state.generalFilter, true, false)
        : (isMoreRestricted ? this.state.filterItems : this.state.items);

      this.filterListItemsByColumnsFilter(itemsToRefine, newFitlers, false);
    }
    catch (error) {
      this.SetError(error, "filterColumnListItems")
    }
  }

  public filterListItemsByColumnsFilter(itemsToRefine: any[], newFilters: IColumnFilter[], isFromClearGeneralFilter: boolean) {
    if (this.props.IndividualColumnFilter) {
      let newItems: Array<any> = [];
      itemsToRefine.map(item => {
        let itemFounded: boolean = true;
        newFilters.map(filter => {
          if (item[filter.columnName] == undefined || item[filter.columnName] == "" || item[filter.columnName].toString().toLowerCase().indexOf(filter.filterToApply.toLowerCase()) < 0) {
            itemFounded = false;
          }
        });
        if (itemFounded) newItems.push(item);
      });

      this.setState({ filterItems: newItems, columnFilters: newFilters, generalFilter: isFromClearGeneralFilter ? "" : this.state.generalFilter });
    }
    else {
      this.setState({ filterItems: itemsToRefine, generalFilter: isFromClearGeneralFilter ? "" : this.state.generalFilter });
    }
  }

  public filterListItemsByGeneralFilter(valueToFilter: string, isClearFilter: boolean, reloadComponents: boolean) {
    if (valueToFilter && valueToFilter.length > 0) {
      let filterItems: Array<any> = [];
      let itemsToFilter = (isClearFilter || valueToFilter.length < this.state.generalFilter.length) ? this.state.items : this.state.filterItems;
      itemsToFilter.map(item => {
        this.props.GeneralSearcheableFields.map(field => {
          if (filterItems.indexOf(item) < 0) {
            if (item[field.ColumnTitle] && item[field.ColumnTitle].toString().toLowerCase().indexOf(valueToFilter.toLowerCase()) > -1) {
              filterItems.push(item);
              return item;
            }
          }
        });

      });
      if (reloadComponents) {
        this.setState({ filterItems, generalFilter: valueToFilter });
      }
      else {
        return filterItems;
      }
    }
    else {
      if (reloadComponents) {
        this.clearGeneralFilter();
      }
      else {
        return this.state.items;
      }
    }
  }

  public clearGeneralFilter() {
    try {
      this.filterListItemsByColumnsFilter(this.state.items, this.state.columnFilters, true);
    }
    catch (error) {
      this.SetError(error, "clearGeneralFilter");
    }
  }


  private _onRenderDetails(detailsFooterProps: IDetailsFooterProps): JSX.Element {
    let _renderDetailsFooterItemColumn: IDetailsRowBaseProps['onRenderItemColumn'] = (item, index, column) => {
      let filter = this.state.columnFilters.filter(colFilter => colFilter.columnName == column.name);
      if (this.props.IndividualColumnFilter) {
        return (
          <SearchBox placeholder={column.name} iconProps={filterIcon} value={filter && filter.length > 0 ? filter[0].filterToApply : ""}
            underlined={true} onChange={(ev, value) => this.filterColumnListItems(column.name, value)} onClear={(ev) => this.filterColumnListItems(column.name, "")} />
        );
      }
      else {
        return undefined;
      }
    };
    return (
      <DetailsRow
        {...detailsFooterProps}
        item={{}}
        itemIndex={-1}
        onRenderItemColumn={_renderDetailsFooterItemColumn}
      />
    );
  }

  private handlePageChange(pageNumber) {
    this.setState({ activePage: pageNumber });
  }

  private _clearAllFilters() {
    this.setState({ columnFilters: [], filterItems: this.state.items, generalFilter: "" });
  }

  private _checkIndividualFilter(position: string): boolean {
    return this.props.IndividualColumnFilter && this.props.IndividualFilterPosition && this.props.IndividualFilterPosition.indexOf(position) > -1;
  }

  private _getItems(): Array<any> {
    let result = [];
    if (this.state.filterItems) {
      if (this.props.ShowPagination) {
        let start = this.props.ItemsInPage * (this.state.activePage - 1);
        result = this.state.filterItems.slice(start, start + this.props.ItemsInPage);
      }
      else {
        result = this.state.filterItems;
      }

    }

    return result;
  }

  private _onItemInvoked = (item: any) => {
    this.props.clickIsCompleteModal && this.GetCompleteItemData(item);
    this.setState({ isModalHidden: false, selectedItem: item, isModalLoading: this.props.clickIsCompleteModal });
  }

  private _closeModalGlosarioModal = (): void => {
    this.setState({ isModalHidden: true, selectedItem: null });
  }

  private async GetCompleteItemData(item: any) {
    let listService: ListService = new ListService(item.SiteUrl);
    let completeItem = await listService.getListItemById(item.ListName, item.Id);
    if (completeItem) {
      completeItem.SiteUrl = item.SiteUrl;
      completeItem.ListName = item.ListName;
    }
    this.setState({ completeModalItemData: completeItem, isModalLoading: false });
  }

  public GetOnClickAction() {
    try {
      if (this.props.clickIsSimpleModal || this.props.clickIsCompleteModal) {
        return this.GetModal();
      }
      else {
        if (this.props.clickIsRedirect) {
          let config = this.props.redirectData.filter(f => f.SiteCollectionSource == this.state.selectedItem.SiteUrl && f.ListSourceField == this.state.selectedItem.ListName);
          if (config && config.length > 0) {
            if (this.props.onRedirectIdQuery) {
              var url = new URL(config[0].Url);
              url.searchParams.append(this.props.onRedirectIdQuery, this.state.selectedItem.Id);
              window.location.replace(url.toString());
            }
            else {
              window.location.replace(`${config[0].Url}`);
            }
          }
        }
        else {
          this.props.onSelectedItem({
            webUrl: this.state.selectedItem.SiteUrl,
            listName: this.state.selectedItem.ListName,
            itemId: this.state.selectedItem.Id
          });
        }
      }
    } catch (error) {
      this.SetError(error, "GetOnClickAction")
    }
  }

  public GetModal = () => {
    try {
      const cancelIcon: IIconProps = { iconName: 'Cancel' };
      const theme = getTheme();
      const iconButtonStyles = {
        root: {
          color: theme.palette.neutralPrimary,
          marginLeft: 'auto',
          marginTop: '4px',
          marginRight: '2px',
          float: 'right'
        },
        rootHovered: {
          color: theme.palette.neutralDark,
        },
      };
      const modal: JSX.Element =
        <Modal
          isOpen={!this.state.isModalHidden}
          onDismiss={this._closeModalGlosarioModal}
          isBlocking={false}
          containerClassName={styles.containerModal}
        >
          <div className={styles.headerModal}>
            {this.state.selectedItem &&
              <IconButton
                styles={iconButtonStyles}
                iconProps={cancelIcon}
                onClick={this._closeModalGlosarioModal}
              />}
          </div>
          <div className={styles.bodyModal}>
            {this.getBodyModal()}
          </div>
        </Modal>
      return modal;
    }
    catch (error) {
      this.SetError(error, "GetModal");
    }
  }

  private getBodyModal() {
    let body: JSX.Element;
    if (this.props.clickIsSimpleModal) {
      body = <>
        {
          this.props.fieldsCollectionData.filter(f => f.SiteCollectionSource == this.state.selectedItem.SiteUrl &&
            f.ListSourceField === this.state.selectedItem.ListName).map(val => {
              return <>
                <div className={styles.propertyModal}>
                  {val.TargetField}
                </div>
                {this.GetRenderByFieldType(this.state.selectedItem, val)}
              </>;
            })
        }
      </>;
    }
    else {
      body = <>
        {this.props.completeModalFields.filter(field => field.SiteCollectionSource == this.state.selectedItem.SiteUrl &&
          field.ListSourceField == this.state.selectedItem.ListName).map(val => {
            return <>
              <div className={styles.propertyModal}>
                {val.TargetField}
              </div>
              <div>
                {this.state.isModalLoading ? <Shimmer /> : this.state.completeModalItemData[val.SourceField]}
              </div>
            </>;
          })
        }
      </>;
    }
    return body;
  }

  private getOnRowClickRender(detailrow: any, defaultRender: any): JSX.Element {
    return this.props.clickEnabled ?
      this.props.oneClickOption ?
        <div onClick={() => this._onItemInvoked(detailrow.item)}>
          {defaultRender({ ...detailrow, styles: { root: { cursor: 'pointer' } } })}
        </div>
        :
        <>
          {defaultRender({ ...detailrow, styles: { root: { cursor: 'pointer' } } })}
        </>
      :
      <>
        {defaultRender({ ...detailrow })}
      </>
  }

  private _renderItemColumn(item: any, index: number, column: IColumn): JSX.Element {
    const fieldContent = item[column.fieldName];
    let result = this.GetJSXElementByType(fieldContent, column.data);
    /*switch (column.data) {
      case SharePointType.User:
        {
          result = <Persona
            {...{
              imageUrl: fieldContent.Email ? `/_layouts/15/userphoto.aspx?UserName=${fieldContent.Email}` : undefined,
              imageInitials: StringUtils.GetUserInitials(fieldContent.Name),
              text: fieldContent.Name
            }}
            size={PersonaSize.size32}
            hidePersonaDetails={false}
          />
          break;
        }
      case SharePointType.UserMulti:
        {
          let personas: IFacepilePersona[] = fieldContent.map(user => {
            let email = StringUtils.GetUserEmail(user.Name);
            return { imageUrl: email ? `/_layouts/15/userphoto.aspx?UserName=${email}` : undefined, personaName: user.Title, imageInitials: StringUtils.GetUserInitials(user.Title), };
          });
          const overflowButtonProps = {
            ariaLabel: 'More users',
          };
          result = <Facepile
            personaSize={PersonaSize.size32}
            personas={personas}
            maxDisplayablePersonas={3}
            overflowButtonType={OverflowButtonType.descriptive}
            overflowButtonProps={overflowButtonProps}
          />
          break;
        }
      case SharePointType.Taxonomy:
        result = <span>{fieldContent.Term}</span>;
        break;
      case SharePointType.TaxonomyMulti:
        {
          if (fieldContent) {
            let labels = fieldContent.map(tax => { return tax.Label });
            result = <span>{labels.map((val, index) => {
              if (index + 1 == labels.lenght) {
                return <span>{val}</span>;
              }
              else {
                return <span>{val}<br></br></span>;
              }
            })}
            </span>
          }
          else {
            result = <span> </span>
          }
          break;
        }
      case SharePointType.Lookup:
        if (fieldContent && fieldContent.Title) {
          result = <Link href="#">{fieldContent.Title}</Link>;
        }
        else {
          result = <span> </span>
        }
        break;
      case SharePointType.LookupMulti:
        if (fieldContent) {
          let values = fieldContent.map(value => { return value.Title });
          result = <span>{values.map((val, index) => {
            if (index + 1 == values.lenght) {
              return <Link href="#">{val}</Link>;
            }
            else {
              return <span><Link href="#">{val}</Link><br></br></span>;
            }
          })}
          </span>
        }
        else {
          result = <span> </span>
        }
        break;
      case SharePointType.Url:
        result = <Link href={fieldContent.Url}>{fieldContent.Description}</Link>;
        break;
      case SharePointType.Image:
        {
          if (fieldContent && fieldContent.Url) {
            const imageProps: IImageProps = {
              src: fieldContent.Url,
              imageFit: ImageFit.contain,
              width: 100,
              height: 100,
            };

            result = <Image
              {...imageProps}
              src={fieldContent.Url}
              alt={fieldContent.Description}
            />
          }
          else {
            result = <span></span>
          }
          break;
        }
      case SharePointType.Boolean:
        {
          result = <span>{fieldContent.toString()}</span>;
          break;
        }
      case SharePointType.NoteFullHtml:
        {
          result = <span dangerouslySetInnerHTML={{ __html: fieldContent }}></span>;
          break;
        }
      default:
        result = <span>{fieldContent}</span>;
        break;
    }*/

    return result;
  }
  private GetRenderByFieldType(item: any, config: IBaseFieldData): JSX.Element {
    let result = this.GetJSXElementByType(item[config.TargetField], config.SPFieldType);

    switch (config.SPFieldType) {
      case SharePointType.Text:
      case SharePointType.Note:
        result = <>
          {item[config.TargetField]}
        </>
        break;
      case SharePointType.Boolean:
        result = <Toggle checked={item[config.TargetField]} />;
        break;
      case SharePointType.Choice:
        const options: IDropdownOption[] = [
          { key: item[config.TargetField], text: item[config.TargetField] },
        ];
        <Dropdown
          label="Disabled example with defaultSelectedKey"
          options={options}
          disabled={true}
        />
        break;
      default:
        result = <>
          {item[config.TargetField]}
        </>;
        break;
    }

    return result;
  }

  private GetJSXElementByType(value: any, fieldType: SharePointType): JSX.Element {
    let result;
    switch (fieldType) {
      case SharePointType.User:
        {
          result = <Persona
            {...{
              imageUrl: value.Email ? `/_layouts/15/userphoto.aspx?UserName=${value.Email}` : undefined,
              imageInitials: StringUtils.GetUserInitials(value.Name),
              text: value.Name
            }}
            size={PersonaSize.size32}
            hidePersonaDetails={false}
          />
          break;
        }
      case SharePointType.UserMulti:
        {
          let personas: IFacepilePersona[] = value.map(user => {
            let email = StringUtils.GetUserEmail(user.Name);
            return { imageUrl: email ? `/_layouts/15/userphoto.aspx?UserName=${email}` : undefined, personaName: user.Title, imageInitials: StringUtils.GetUserInitials(user.Title), };
          });
          const overflowButtonProps = {
            ariaLabel: 'More users',
          };
          result = <Facepile
            personaSize={PersonaSize.size32}
            personas={personas}
            maxDisplayablePersonas={3}
            overflowButtonType={OverflowButtonType.descriptive}
            overflowButtonProps={overflowButtonProps}
          />
          break;
        }
      case SharePointType.Taxonomy:
        result = <span>{value.Term}</span>;
        break;
      case SharePointType.TaxonomyMulti:
        {
          if (value) {
            let labels = value.map(tax => { return tax.Label });
            result = <span>{labels.map((val, index) => {
              if (index + 1 == labels.lenght) {
                return <span>{val}</span>;
              }
              else {
                return <span>{val}<br></br></span>;
              }
            })}
            </span>
          }
          else {
            result = <span> </span>
          }
          break;
        }
      case SharePointType.Lookup:
        if (value && value.Title) {
          result = <Link href="#">{value.Title}</Link>;
        }
        else {
          result = <span> </span>
        }
        break;
      case SharePointType.LookupMulti:
        if (value) {
          let values = value.map(value => { return value.Title });
          result = <span>{values.map((val, index) => {
            if (index + 1 == values.lenght) {
              return <Link href="#">{val}</Link>;
            }
            else {
              return <span><Link href="#">{val}</Link><br></br></span>;
            }
          })}
          </span>
        }
        else {
          result = <span> </span>
        }
        break;
      case SharePointType.Url:
        result = <Link href={value.Url}>{value.Description}</Link>;
        break;
      case SharePointType.Image:
        {
          if (value && value.Url) {
            const imageProps: IImageProps = {
              src: value.Url,
              imageFit: ImageFit.contain,
              width: 100,
              height: 100,
            };

            result = <Image
              {...imageProps}
              src={value.Url}
              alt={value.Description}
            />
          }
          else {
            result = <span></span>
          }
          break;
        }
      case SharePointType.Boolean:
        {
          result = <span>{value.toString()}</span>;
          break;
        }
      case SharePointType.NoteFullHtml:
        {
          result = <span dangerouslySetInnerHTML={{ __html: value }}></span>;
          break;
        }
      default:
        result = <span>{value}</span>;
        break;
    }

    return result;
  }


  public render(): React.ReactElement<IListSearchProps> {
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;
    let clearAllButton = this.props.ClearAllFiltersBtnColor == "white" ? <DefaultButton text={this.props.ClearAllFiltersBtnText} className={styles.btn} onClick={(ev) => this._clearAllFilters()} /> :
      <PrimaryButton text={this.props.ClearAllFiltersBtnText} className={styles.btn} onClick={(ev) => this._clearAllFilters()} />;
    return (
      <div className={styles.listSearch} style={{ backgroundColor: semanticColors.bodyBackground }}>
        <div className={styles.row}>
          <div className={styles.column}>
            {this.state.isLoading ?
              <Spinner label={strings.ListSearchLoading} size={SpinnerSize.large} style={{ backgroundColor: semanticColors.bodyBackground }} /> :
              this.state.errorMsg ?
                <MessageBar
                  messageBarType={MessageBarType.error}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >{this.state.errorMsg}
                </MessageBar> :
                <React.Fragment>
                  {this.props.clickEnabled && !this.state.isModalHidden && this.state.selectedItem && this.GetOnClickAction()}
                  <div className={styles.rowTopInformation}>
                    {this.props.GeneralFilter && <div className={this.props.ShowClearAllFilters ? styles.ColGeneralFilterWithBtn : styles.ColGeneralFilterOnly}><SearchBox value={this.state.generalFilter} placeholder={this.props.GeneralFilterPlaceHolderText} onClear={() => this.clearGeneralFilter()} onChange={(ev, newValue) => this.filterListItemsByGeneralFilter(newValue, false, true)} /></div>}
                    <div className={styles.ColClearAll}>
                      {this.props.ShowClearAllFilters && clearAllButton}
                    </div>
                  </div>
                  <div className={styles.rowData}>
                    <div className={styles.colData}>
                      {this.props.ShowItemCount && <div className={styles.template_resultCount}>{this.props.ItemCountText.replace("{itemCount}", `${this.state.filterItems.length}`)}</div>}
                      <DetailsList items={this._getItems()} columns={this.columns}
                        onRenderDetailsFooter={this._checkIndividualFilter("footer") ? (detailsFooterProps) => this._onRenderDetails(detailsFooterProps) : undefined}
                        onRenderDetailsHeader={this._checkIndividualFilter("header") ? (detailsHeaderProps) => this._onRenderDetails(detailsHeaderProps) : undefined}
                        className={styles.searchListData}
                        selectionMode={SelectionMode.none}
                        onItemInvoked={this.props.clickEnabled && !this.props.oneClickOption ? this._onItemInvoked : null}
                        onRenderRow={(props, defaultRender) => this.getOnRowClickRender(props, defaultRender)}
                        onRenderItemColumn={this._renderItemColumn}
                      />
                      {this.props.ShowPagination &&
                        <div className={styles.paginationContainer}>
                          <div className={styles.paginationContainer__paginationContainer}>
                            <div className={`${styles.paginationContainer__paginationContainer__pagination}`}>
                              <div className={styles.standard}>
                                <Pagination
                                  activePage={this.state.activePage}
                                  firstPageText={<Icon theme={this.props.themeVariant as ITheme} iconName='DoubleChevronLeft' />}
                                  lastPageText={<Icon theme={this.props.themeVariant as ITheme} iconName='DoubleChevronRight' />}
                                  prevPageText={<Icon theme={this.props.themeVariant as ITheme} iconName='ChevronLeft' />}
                                  nextPageText={<Icon theme={this.props.themeVariant as ITheme} iconName='ChevronRight' />}
                                  activeLinkClass={styles.active}
                                  itemsCountPerPage={this.props.ItemsInPage}
                                  totalItemsCount={this.state.filterItems ? this.state.filterItems.length : 0}
                                  pageRangeDisplayed={5}
                                  onChange={this.handlePageChange.bind(this)}
                                />
                              </div>
                            </div>
                          </div>
                        </div>
                      }
                    </div>
                  </div>
                </React.Fragment>}
          </div>
        </div>
      </div >);
  }
}
