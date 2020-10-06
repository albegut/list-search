export interface IListFieldData {
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
  Order: number;
  sortIdx: number;
  FieldType: SharePointFieldTypes;
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
  FieldType: SharePointFieldTypes;
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
