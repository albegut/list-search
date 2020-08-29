export interface IListFieldData{
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
  Order: number;
  Searcheable: boolean;
}

export interface IListData{
  SiteCollectionSource: string;
  ListSourceField: string;
  ListView: string;
  Query: string;
}
