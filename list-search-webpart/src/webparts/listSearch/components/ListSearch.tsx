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

export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {
  columns: IColumn[] = [];
  keymapQuerys: {} = {};
  constructor(props: IListSearchProps, state: IListSearchState) {
    super(props);
    this.state = {
      items: null,
      filterItems: null,
      isLoading: true,
      errorMsg: "",
    };

  }

  public filterListItems(propertyName: string, propertyValue: string) {
    let filterItems: Array<any> = this.state.items.filter(item => {
      return item[propertyName] && item[propertyName].toString() === propertyValue
    });
    this.setState({ filterItems });
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
        Object.keys(this.keymapQuerys[site]).map(listQuery => {
          itemPromise.push(listService.getListItems(this.keymapQuerys[site][listQuery], this.props.ListNameTitle, this.props.SiteNameTitle));
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
    this.props.collectionData.map(item => {
      if (this.keymapQuerys[item.SiteCollectionSource] != undefined) {
        if (this.keymapQuerys[item.SiteCollectionSource][item.ListSoruceField] != undefined) {
          this.keymapQuerys[item.SiteCollectionSource][item.ListSoruceField].fields.push({ originalField: item.SoruceField, newField: item.TargetField });
        }
        else {
          let newQueryListItem: IListSearchListQuery = { list: item.ListSoruceField, fields: [{ originalField: item.SoruceField, newField: item.TargetField }] };
          this.keymapQuerys[item.SiteCollectionSource][item.ListSoruceField] = newQueryListItem;
        }
      }
      else {
        let newQueryListItem: IListSearchListQuery = { list: item.ListSoruceField, fields: [{ originalField: item.SoruceField, newField: item.TargetField }] };
        this.keymapQuerys[item.SiteCollectionSource] = {};
        this.keymapQuerys[item.SiteCollectionSource][item.ListSoruceField] = newQueryListItem;
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



  private _onRenderDetailsFooter(detailsFooterProps: IDetailsFooterProps): JSX.Element {
    let _renderDetailsFooterItemColumn: IDetailsRowBaseProps['onRenderItemColumn'] = (item, index, column) => {
      if (column) {
        return (
          <TextField onChange={(ev, value) => this.filterListItems(column.name, value)}  />
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


  public render(): React.ReactElement<IListSearchProps> {
    console.log(this.state.filterItems)
    return (
      <div className={styles.listSearch}>
        <div className={styles.container}>
          <div className={styles.row}>
            {this.state.isLoading ?
              <Spinner label="Cargando..." size={SpinnerSize.large} /> :
              this.state.errorMsg ?
                <MessageBar
                  messageBarType={MessageBarType.error}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >{this.state.errorMsg}
                </MessageBar> :
                <React.Fragment><p>The data bellow has {this.state.filterItems ? this.state.filterItems.length : 0} items</p>
                  <DetailsList items={this.state.filterItems || []} columns={this.columns.sort((prev, next) => prev.data - next.data)}
                    onRenderDetailsFooter={(detailsFooterProps) =>this._onRenderDetailsFooter(detailsFooterProps)} />
                </React.Fragment>}
          </div>
        </div>
      </div>);
  }
}
