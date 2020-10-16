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
}

export interface IColumnFilter {
  columnName: string;
  filterToApply: string;
}
