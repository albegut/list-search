import * as React from 'react';

import styles from '../ListSearchWebPart.module.scss';
import * as strings from 'ListSearchWebPartStrings';

import IListService from '../services/IListService'
import ListService from '../services/ListService';

import { IListSearchState } from './IListSearchState';
import { IListSearchProps } from './IListSearchProps';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import {
  DetailsList,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { IListSearchListQuery } from '../model/ListSearchQuery';




export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {
  columns: IColumn[] = [];

  constructor(props: IListSearchProps, state: IListSearchState) {
    super(props);
    this.state = {
      items: null,
      isLoading: true,
      errorMsg: "",
    };

  }

  public componentDidMount() {
    this.readItems();
  }

  private addColumnIfNotExists(columnDisplayName): void {
    if (this.columns.filter(column => column.key == columnDisplayName).length == 0) {
      this.columns.push({ key: columnDisplayName, name: columnDisplayName, fieldName: columnDisplayName, minWidth: 100, maxWidth: 200, isResizable: true });
    }
  }

  private async readItems() {
    //
    //let keymapQuerys = {};

    let keymapQuerys = {};
    this.props.collectionData.map(item => {
      if (keymapQuerys[item.SiteCollectionSource] != undefined) {
        if (keymapQuerys[item.SiteCollectionSource][item.ListSoruceField] != undefined) {
          keymapQuerys[item.SiteCollectionSource][item.ListSoruceField].viewFields.push({originalField: item.SoruceField, newField: item.TargetField});
        }
        else {
          let newQueryListItem: IListSearchListQuery = {list: item.ListSoruceField, fields:[{originalField: item.SoruceField, newField: item.TargetField}]};
          keymapQuerys[item.SiteCollectionSource][item.ListSoruceField] = newQueryListItem;
        }
      }
      else {
        let newQueryListItem: IListSearchListQuery = {list: item.ListSoruceField, fields:[{originalField: item.SoruceField, newField: item.TargetField}]};
        keymapQuerys[item.SiteCollectionSource] = {};
        keymapQuerys[item.SiteCollectionSource][item.ListSoruceField] = newQueryListItem;
      }
      this.addColumnIfNotExists(item.TargetField);
    });

    if (this.props.ShowListName) {
      this.addColumnIfNotExists(this.props.ListNameTitle);
    }

    if (this.props.ShowSite) {
      this.addColumnIfNotExists(this.props.SiteNameTitle);
    }

    let itemPromise: Array<Promise<Array<any>>> = [];
    try {
      Object.keys(keymapQuerys).map(site => {
        let listService: ListService = new ListService(site);
        Object.keys(keymapQuerys[site]).map(listQuery => {
          itemPromise.push(listService.getListItems(keymapQuerys[site][listQuery], this.props.ListNameTitle, this.props.SiteNameTitle));
        })
      })

      let items = await Promise.all(itemPromise);
      let result = [];
      items.map(partialResult => {
        result.push(...partialResult);
      });

      this.setState({
        items: result,
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        errorMsg: "readItemsError",
        isLoading: false,
      });
    }
  }

  public render(): React.ReactElement<IListSearchProps> {

    return (
      <div className={styles.listSearch}>
        <div className={styles.container}>
          <div className={styles.row}>
            {this.state.isLoading ? <Spinner label="Cargando..." size={SpinnerSize.large} /> :
              <React.Fragment><p>The data bellow has {this.state.items ? this.state.items.length : 0} items</p>
                <DetailsList items={this.state.items} columns={this.columns} />
              </React.Fragment>}
          </div>
        </div>
      </div>);
  }
}
