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
  uniqueId: string;
  sortIdx: number;
}

export interface IDisplayFieldData {
  IsSiteTitle: boolean;
  IsListTitle: boolean;
  ColumnTitle: string;
  ColumnWidth?: number;
  Searcheable: boolean;
}

export interface ICompleteModalData {
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
}

export interface IRedirectData {
  SiteCollectionSource: string;
  ListSourceField: string;
  Url: string;
}
