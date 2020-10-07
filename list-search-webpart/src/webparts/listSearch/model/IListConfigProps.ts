export interface IBaseFieldData{
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
  FieldType: string;
}


export interface IListFieldData extends IBaseFieldData{
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

export interface IDisplayFieldData {
  IsSiteTitle: boolean;
  IsListTitle: boolean;
  ColumnTitle: string;
  ColumnWidth?: number;
  Searcheable: boolean;
}

export interface ICompleteModalData extends IBaseFieldData{

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

export type  SharePointFieldTypes = {
  Text:'Text',
  Note:'Note',
  Choice:'Choice',
  Integer:'Integer',
  Number:'Number',
  Money:'Money',
  DateTime:'DateTime',
  Lookup:'Lookup',
  LookupMulti:'LookupMulti',
  Boolean:'Boolean',
  User:'Note',
  UserMulti:'Note',
  Url:'Note',
  Calculated:'Note',
  Image:'Note',
  Taxonomy:'Note',
  Computed:'Note',
  Attachments:'Note',
  Counter:'Note',
  ContentTypeId:'Note',
  Guid:'Note',
}

/*
export enum SharePointFieldTypes {
  Text = 0,
  Note,
  Choice,
  Integer,
  Number,
  Money,
  DateTime,
  Lookup,
  LookupMulti,
  Boolean,
  User,
  UserMulti,
  Url,
  Calculated,
  Image,
  Taxonomy,
  Computed,
  Attachments,
  Counter,
  ContentTypeId,
  Guid,
}
*/
