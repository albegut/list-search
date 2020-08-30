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
  IDetailsRowCheckStyles,
  DetailsRowCheck,
  DetailsRow,
  SelectionMode,
} from 'office-ui-fabric-react/lib/DetailsList';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IListSearchListQuery } from '../model/ListSearchQuery';
import {
  MessageBar,
  MessageBarType
} from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import Pagination from "react-js-pagination";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';




const filterIcon: IIconProps = { iconName: 'Filter' };

export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {
  columns: IColumn[] = [];
  keymapQuerys: {} = {};
  constructor(props: IListSearchProps, state: IListSearchState) {
    super(props);
    this.state = {
      activePage: 1,
      items: null,
      filterItems: null,
      isLoading: true,
      errorMsg: "",
      columnFilters: [],
      generalFilter: ""
    };

  }

  componentWillReceiveProps() {
    console.log("new props");
    console.log(this.props.collectionData)
    this.columns = [];
    this.setState({ items: null, isLoading: true })
    this.readItems();
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
        let siteProperties = this.props.Sites.filter(siteInformation => { if (siteInformation.url === site) return siteInformation })
        Object.keys(this.keymapQuerys[site]).map(listQuery => {
          itemPromise.push(listService.getListItems(this.keymapQuerys[site][listQuery], this.props.ListNameTitle, this.props.SiteNameTitle, siteProperties[0][this.props.SiteNamePropertyToShow]));
        })
      })

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
        errorMsg: "readItemsError",
        isLoading: false,
      });
    }
  }

  private generateKeymap() {
    //BUG WITH NEW FILTERS
    this.props.collectionData.map(item => {
      if (this.keymapQuerys[item.SiteCollectionSource] != undefined) {
        if (this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] != undefined) {
          if (this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField].fields.filter(field => { if (field.originalField === item.SourceField) return field }).length == 0) {
            this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField].fields.push({ originalField: item.SourceField, newField: item.TargetField });
          }
        }
        else {
          let listQueryInfo = this.props.ListscollectionData.filter(list => { if (list.SiteCollectionSource == item.SiteCollectionSource && list.ListSourceField == item.ListSourceField) return list });

          let newQueryListItem: IListSearchListQuery = { list: item.ListSourceField, fields: [{ originalField: item.SourceField, newField: item.TargetField }], camlQuery: listQueryInfo.length > 0 && listQueryInfo[0].Query, viewName: listQueryInfo.length > 0 && listQueryInfo[0].ListView };
          this.keymapQuerys[item.SiteCollectionSource][item.ListSourceField] = newQueryListItem;
        }
      }
      else {
        let listQueryInfo = this.props.ListscollectionData.filter(list => { if (list.SiteCollectionSource == item.SiteCollectionSource && list.ListSourceField == item.ListSourceField) return list });

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

    let columnFiltersToApply = this.state.columnFilters.length > 0 ? this.state.columnFilters : [{ columnName: propertyName, filterToApply: propertyValue }];
    let isNewFilter: boolean = true;
    let clearFilter: boolean = false;
    let newFitlers: IColumnFilter[] = columnFiltersToApply.filter(filter => {
      if (filter.columnName === propertyName) {
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

    if (isNewFilter) newFitlers.push({ columnName: propertyName, filterToApply: propertyValue })

    let itemsToRefine = clearFilter || this.state.generalFilter ? this.filterListItemsByGeneralFilter(this.state.generalFilter, true, false) : this.state.filterItems;

    this.filterListItemsByColumnsFilter(itemsToRefine, newFitlers, false);
  }

  public filterListItemsByColumnsFilter(itemsToRefine: any[], newFilters: IColumnFilter[], isFromClearGeneralFilter: boolean) {
    if (this.props.IndividualColumnFilter) {
      let newItems: Array<any> = [];
      itemsToRefine.filter(item => {
        let itemFounded: boolean = true;
        newFilters.map(filter => {
          if (item[filter.columnName] == undefined || item[filter.columnName] == "" || item[filter.columnName].toString().toLowerCase().indexOf(filter.filterToApply.toLowerCase()) < 0) {
            itemFounded = false;
          }
        })
        if (itemFounded) newItems.push(item);
      })

      this.setState({ filterItems: newItems, columnFilters: newFilters, generalFilter: isFromClearGeneralFilter ? "" : this.state.generalFilter });
    }
    else {
      this.setState({ generalFilter: isFromClearGeneralFilter ? "" : this.state.generalFilter });
    }
  }

  public filterListItemsByGeneralFilter(valueToFilter: string, isClearFilter: boolean, reloadComponents: boolean) {
    if (valueToFilter && valueToFilter.length > 0) {
      let filterItems: Array<any> = []
      let itemsToFilter = isClearFilter ? this.state.items : this.state.filterItems;
      itemsToFilter.filter(item => {
        this.props.GeneralSearcheableFields.map(field => {
          if (filterItems.indexOf(item) < 0) {
            if (item[field.TargetField] && item[field.TargetField].toString().toLowerCase().indexOf(valueToFilter.toLowerCase()) > -1) {
              filterItems.push(item);
              return item;
            }
          }
        })

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


  private _onRenderDetailsFooter(detailsFooterProps: IDetailsFooterProps): JSX.Element {
    if (this.props.IndividualColumnFilter) {
      let _renderDetailsFooterItemColumn: IDetailsRowBaseProps['onRenderItemColumn'] = (item, index, column) => {
        if (column) {
          return (
            <SearchBox placeholder={column.name} iconProps={filterIcon}
              underlined={true} onChange={(ev, value) => this.filterColumnListItems(column.name, value)} onClear={(ev) => this.filterColumnListItems(column.name, "")} />
          );
        }
        return undefined;
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
    else {
      return <React.Fragment />
    }
  }

  handlePageChange(pageNumber) {
    this.setState({ activePage: pageNumber });
  }


  public render(): React.ReactElement<IListSearchProps> {

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    return (
      <div className={styles.listSearch} style={{
        backgroundColor: semanticColors.bodyBackground
      }}>
        <div className={styles.row}>
          {this.state.isLoading ?
            <Spinner label="Cargando..." size={SpinnerSize.large} style={{ backgroundColor: semanticColors.bodyBackground }} /> :
            this.state.errorMsg ?
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.errorMsg}
              </MessageBar> :
              <React.Fragment>
                {this.props.GeneralFilter && <SearchBox placeholder={this.props.GeneralFilterPlaceHolderText} onClear={() => this.clearGeneralFilter()} onChange={(ev, newValue) => this.filterListItemsByGeneralFilter(newValue, false, true)} />}
                <div>{this.props.ShowItemCount && this.props.ItemCountText.replace("{itemCount}", this.state.filterItems.length.toString())}</div>
                <DetailsList items={this.state.filterItems || []} columns={this.columns.sort((prev, next) => prev.data - next.data)}
                  onRenderDetailsFooter={(detailsFooterProps) => this._onRenderDetailsFooter(detailsFooterProps)} />
                {this.props.ShowPagination &&
                  <Pagination
                    activePage={this.state.activePage}
                    itemsCountPerPage={this.props.ItemsInPage}
                    totalItemsCount={this.state.items ? this.state.items.length : 0}
                    pageRangeDisplayed={5}
                    onChange={this.handlePageChange.bind(this)}
                  />
                }
              </React.Fragment>}
        </div>
      </div >);
  }
}
