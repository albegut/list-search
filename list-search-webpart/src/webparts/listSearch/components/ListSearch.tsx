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




export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {
  private listService: IListService;
  columnsNames: string[] = [];
  columns: IColumn[] = [];

  constructor(props: IListSearchProps, state: IListSearchState) {
    super(props);
    this.listService = new ListService(this.props.Context);
    this.state = {
      items: null,
      isLoading: true,
      errorMsg: "",
    };

  }

  //TODO revisar este metodo esta deprecado
  public componentWillReceiveProps(nextProps: IListSearchProps) {
    if (nextProps.ListName != this.props.ListName) {
      this.readItems();
    }
  }

  public componentDidMount() {
    this.readItems();
  }

  private addColumnIfNotExists(columnName: string, columnDisplayName): void {
    if (this.columnsNames.indexOf(columnName) < 0) {
      this.columns.push({ key: columnName, name: columnDisplayName, fieldName: columnName, minWidth: 100, maxWidth: 200, isResizable: true });
    }
  }

  private async readItems() {
    //let keymap: { listName: string, viewFields: Array<string> };
    let keymapQuerys = {};
    this.props.collectionData.map(item => {
      if (keymapQuerys[item.ListSoruceField] != undefined) {
        keymapQuerys[item.ListSoruceField].push(item.SoruceField);
        this.addColumnIfNotExists(item.SoruceField, item.TargetField);
      }
      else {
        keymapQuerys[item.ListSoruceField] = [item.SoruceField];
        this.addColumnIfNotExists(item.SoruceField, item.TargetField);
        if(this.props.ShowListName)
        {
          this.addColumnIfNotExists("CustomListFieldName", this.props.ListNameTitle);
        }
      }
    });

    let itemPromise: Array<Promise<Array<any>>> = [];
    try {
      Object.keys(keymapQuerys).map(listQuery => {
        itemPromise.push(this.listService.getListItems(listQuery, keymapQuerys[listQuery], "ID", true));
      })

      let items = await Promise.all(itemPromise);
      let result = [];
      items.map(partialResult => {
        result.push(...partialResult);
      });

      console.log(result);
      console.log(this.columns);
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
