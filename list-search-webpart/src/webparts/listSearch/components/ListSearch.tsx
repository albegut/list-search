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
} from 'office-ui-fabric-react/lib/DetailsList';
import { IListSearchListQuery } from '../model/ListSearchQuery';
import {
  MessageBar,
  MessageBarType
} from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import Pagination from "react-js-pagination";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';




const filterIcon: IIconProps = { iconName: 'Filter' };

export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {
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
    };

  }

  public componentDidUpdate(prevProps: Readonly<IListSearchProps>, prevState: Readonly<IListSearchState>, snapshot?: any): void {
    if (prevProps != this.props) {
      this.columns = [];
      this.setState({ items: null, filterItems: null, isLoading: true });
      this.readItems();
    }
  }

  public componentDidMount() {
    this.readItems();
  }

  private addColumnIfNotExists(columnDisplayName: string, orderColumn: number): void {
    if (this.columns.filter(column => column.key == columnDisplayName).length == 0) {
      this.columns.push({ key: columnDisplayName, name: columnDisplayName, fieldName: columnDisplayName, minWidth: 100, maxWidth: 200, isResizable: true, data: orderColumn });
    }
  }

  private async readItems() {
    this.generateKeymap();
    let itemPromise: Array<Promise<Array<any>>> = [];
    try {
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

      this.setState({
        items: result,
        filterItems: result,
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        errorMsg: `readItemsError ${error.message}`,
        isLoading: false,
      });
    }
  }

  private generateKeymap() {
    this.props.fieldsCollectionData.map(item => {
      if (this.keymapQuerys[item.SiteCollectionSource] != undefined) {
        if (this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] != undefined) {
          if (this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField].fields.filter(field => field.originalField === item.SourceField).length == 0) {
            this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField].fields.push({ originalField: item.SourceField, newField: item.TargetField });
          }
        }
        else {
          let listQueryInfo = this.props.listsCollectionData.filter(list => list.SiteCollectionSource == item.SiteCollectionSource && list.ListSourceField == item.ListSourceField);

          let newQueryListItem: IListSearchListQuery = { list: item.ListSourceField, fields: [{ originalField: item.SourceField, newField: item.TargetField }], camlQuery: listQueryInfo.length > 0 && listQueryInfo[0].Query, viewName: listQueryInfo.length > 0 && listQueryInfo[0].ListView };
          this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] = newQueryListItem;
        }
      }
      else {
        let listQueryInfo = this.props.listsCollectionData.filter(list => list.SiteCollectionSource == item.SiteCollectionSource && list.ListSourceField == item.ListSourceField);

        let newQueryListItem: IListSearchListQuery = { list: item.ListSourceField, fields: [{ originalField: item.SourceField, newField: item.TargetField }], camlQuery: listQueryInfo.length > 0 && listQueryInfo[0].Query, viewName: listQueryInfo.length > 0 && listQueryInfo[0].ListView };
        this.keymapQuerys[item.SiteCollectionSource] = {};
        this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] = newQueryListItem;
      }
      this.addColumnIfNotExists(item.TargetField, item.Order);
    });

    if (this.props.ShowListName) {
      this.addColumnIfNotExists(this.props.ListNameTitle, this.props.ListNameOrder || Number.MAX_VALUE);
    }

    if (this.props.ShowSite) {
      this.addColumnIfNotExists(this.props.SiteNameTitle, this.props.SiteNameOrder || Number.MAX_VALUE);
    }
  }

  public filterColumnListItems(propertyName: string, propertyValue: string) {
    let isNewFilter: boolean = true;
    let clearFilter: boolean = false;
    let isMoreRestricted: boolean = false;
    let newFitlers: IColumnFilter[] = this.state.columnFilters.map(filter => {
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
      this.setState({ generalFilter: isFromClearGeneralFilter ? "" : this.state.generalFilter });
    }
  }

  public filterListItemsByGeneralFilter(valueToFilter: string, isClearFilter: boolean, reloadComponents: boolean) {
    if (valueToFilter && valueToFilter.length > 0) {
      let filterItems: Array<any> = [];
      let itemsToFilter = (isClearFilter || valueToFilter.length < this.state.generalFilter.length) ? this.state.items : this.state.filterItems;
      itemsToFilter.map(item => {
        this.props.GeneralSearcheableFields.map(field => {
          if (filterItems.indexOf(item) < 0) {
            if (item[field.TargetField] && item[field.TargetField].toString().toLowerCase().indexOf(valueToFilter.toLowerCase()) > -1) {
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
    this.filterListItemsByColumnsFilter(this.state.items, this.state.columnFilters, true);
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
                  <div className={styles.rowTopInformation}>
                    {this.props.GeneralFilter && <div className={styles.ColGeneralFilter}><SearchBox value={this.state.generalFilter} placeholder={this.props.GeneralFilterPlaceHolderText} onClear={() => this.clearGeneralFilter()} onChange={(ev, newValue) => this.filterListItemsByGeneralFilter(newValue, false, true)} /></div>}
                    <div className={styles.ColClearAll}>
                      {this.props.ShowClearAllFilters && clearAllButton}
                    </div>
                  </div>
                  <div className={styles.rowData}>
                    <div className={styles.colData}>
                      {this.props.ShowItemCount && this.props.ItemCountText.replace("{itemCount}", this.state.filterItems.length.toString())}
                      <DetailsList items={this.state.filterItems || []} columns={this.columns.sort((prev, next) => prev.data - next.data)}
                        onRenderDetailsFooter={this._checkIndividualFilter("footer") ? (detailsFooterProps) => this._onRenderDetails(detailsFooterProps) : undefined}
                        onRenderDetailsHeader={this._checkIndividualFilter("header") ? (detailsHeaderProps) => this._onRenderDetails(detailsHeaderProps) : undefined}
                        className={styles.searchListData} />
                      {this.props.ShowPagination &&
                        <Pagination
                          activePage={this.state.activePage}
                          itemsCountPerPage={this.props.ItemsInPage}
                          totalItemsCount={this.state.items ? this.state.items.length : 0}
                          pageRangeDisplayed={5}
                          onChange={this.handlePageChange.bind(this)}
                        />
                      }
                    </div>
                  </div>
                </React.Fragment>}
          </div>
        </div>
      </div >);
  }
}
