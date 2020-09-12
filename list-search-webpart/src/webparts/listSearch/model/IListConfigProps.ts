export interface IListFieldData {
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
}

export interface IListData {
  SiteCollectionSource: string;
  ListSourceField: string;
  ListView: string;
  Query: string;
}

export interface IDisplayFieldData {
  IsSiteTitle: boolean;
  IsListTitle: boolean;
  ColumnTitle: string;
  ColumnWidth?: number;
  Order: number;
  Searcheable: boolean;
}
