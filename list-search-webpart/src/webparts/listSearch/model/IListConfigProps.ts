import { SharePointType } from "./ISharePointFieldTypes";

export interface IBaseFieldData {
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
  SPFieldType: SharePointType;
}


export interface IMappingFieldData extends IBaseFieldData {
  Order: number;
  sortIdx: number;
}

export interface IListData {
  SiteCollectionSource: string;
  ListSourceField: string;
  ListView: string;
  Query: string;
  uniqueId: string;
  sortIdx: number;
}

export interface IDetailListFieldData {
  IsSiteTitle: boolean;
  IsListTitle: boolean;
  ColumnTitle: string;
  ColumnWidth?: number;
  Searcheable: boolean;
}

export interface ICompleteModalData extends IBaseFieldData {

}

export interface IRedirectData {
  SiteCollectionSource: string;
  ListSourceField: string;
  Url: string;
}

export interface ICustomOption {
  Key: string;
  Option: string;
  CustomData: string;
}
