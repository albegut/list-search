import * as React from 'react';
import styles from '../ListSearchWebPart.module.scss';
import * as strings from 'ListSearchWebPartStrings';
import ListService from '../services/ListService';
import { IListSearchState } from './IListSearchState';
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
import { SearchBox, ISearchBoxStyles } from 'office-ui-fabric-react/lib/SearchBox';
import Pagination from "react-js-pagination";
import { IReadonlyTheme } from '@microsoft/sp-component-base';



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
    let filterItems: Array<any> = this.state.items.filter(item => {
      return item[propertyName] && item[propertyName].toString().toLowerCase().indexOf(propertyValue.toLowerCase()) > -1
    });
    this.setState({ filterItems });
  }

  public filterListItems(valueToFilter: string) {
    if (valueToFilter && valueToFilter.length > 0) {
      let filterItems: Array<any> = []
      this.state.items.filter(item => {
        this.props.GeneralSearcheableFields.map(field => {
          if (filterItems.indexOf(item) < 0) {
            if (item[field.TargetField] && item[field.TargetField].toString().toLowerCase().indexOf(valueToFilter.toLowerCase()) > -1) {
              filterItems.push(item);
              return item;
            }
          }
        })

      });
      this.setState({ filterItems });
    }
    else {
      this.clearGeneralFilter();
    }
  }

  public clearGeneralFilter() {
    this.setState({ filterItems: this.state.items });
  }


  private _onRenderDetailsFooter(detailsFooterProps: IDetailsFooterProps): JSX.Element {
    if (this.props.IndividualColumnFilter) {
      let _renderDetailsFooterItemColumn: IDetailsRowBaseProps['onRenderItemColumn'] = (item, index, column) => {
        if (column) {
          return (
            <TextField onChange={(ev, value) => this.filterColumnListItems(column.name, value)} />
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
      <div className={styles.listSearch} style={{ backgroundColor: semanticColors.bodyBackground }}>
        <div className={styles.container}>
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
                  {this.props.GeneralFilter && <SearchBox placeholder={this.props.GeneralFilterPlaceHolderText} onClear={() => this.clearGeneralFilter()} onChange={(ev, newValue) => this.filterListItems(newValue)} />}
                  <div >{this.props.ShowItemCount && this.props.ItemCountText.replace("{itemCount}", this.state.filterItems.length.toString())}</div>
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
        </div>
      </div>);
  }
}
