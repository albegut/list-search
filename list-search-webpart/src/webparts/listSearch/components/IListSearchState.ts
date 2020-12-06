import { IColumn } from 'office-ui-fabric-react';
import { SharePointType } from '../model/ISharePointFieldTypes';
export interface IListSearchState {
  isLoading: boolean;
  errorMsg: string;
  errorHeader: string;
  items: Array<any>;
  filterItems: Array<any>;
  generalFilter: string;
  columnFilters: IColumnFilter[];
  activePage: number;
  isModalHidden: boolean;
  isModalLoading: boolean;
  selectedItem: any;
  completeModalItemData: any;
  groupedItems: IGroupedItems[];
  columns: IColumn[];
}

export default interface IGroupedItems {
  GroupName: string;
  Items: any[];
}


export interface IColumnFilter {
  columnName: string;
  filterToApply: string;
  columnType: SharePointType;
}
