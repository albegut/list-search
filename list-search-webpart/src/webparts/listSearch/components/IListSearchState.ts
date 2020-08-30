export interface IListSearchState {
  isLoading: boolean;
  errorMsg: string;
  items: Array<any>;
  filterItems: Array<any>;
  generalFilter: string;
  columnFilters: IColumnFilter[];
  activePage: number;
}

export interface IColumnFilter {
  columnName: string;
  filterToApply: string;
}
