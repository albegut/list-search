import { SharePointType } from "./ISharePointFieldTypes";

export interface IBaseFieldData {
  SiteCollectionSource: string;
  ListSourceField: string;
  SourceField: string;
  TargetField: string;
  SPFieldType: SharePointType;
}


export interface IMappingFieldData extends IBaseFieldData {
  uniqueId: string;
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

export class IDetailListFieldData {

  constructor(IsSiteTitle: boolean, IsListTitle: boolean, IsFileIcon: boolean, ColumnTitle: string, MinColumnWidth: number, MaxColumnWidth: number, Searcheable: boolean) {
    this.IsSiteTitle = IsSiteTitle;
    this.IsListTitle = IsListTitle;
    this.IsFileIcon = IsFileIcon;
    this.ColumnTitle = ColumnTitle;
    this.MinColumnWidth = MinColumnWidth;
    this.MaxColumnWidth = MaxColumnWidth;
    this.Searcheable = Searcheable;
  }

  public static CreateListColumn(MinColumnWidth: number, MaxColumnWidth: number, Searcheable: boolean): IDetailListFieldData {
    return new IDetailListFieldData(false, true, false, "ListName", MinColumnWidth, MaxColumnWidth, Searcheable);
  }

  public static CreateSiteColumn(MinColumnWidth: number, MaxColumnWidth: number, Searcheable: boolean): IDetailListFieldData {
    return new IDetailListFieldData(true, false, false, "Site", MinColumnWidth, MaxColumnWidth, Searcheable);
  }

  public static CreateFileColumn(): IDetailListFieldData {
    return new IDetailListFieldData(false, false, true, "FileIcon", 30, 30, false);
  }

  public static IsGeneralColumn(object: IDetailListFieldData): boolean {
    return object.IsSiteTitle != false && object.IsListTitle != false && object.IsFileIcon != false; //undefined values are also general columns
  }

  public IsSiteTitle: boolean;
  public IsListTitle: boolean;
  public IsFileIcon: boolean;
  public ColumnTitle: string;
  public MinColumnWidth?: number;
  public MaxColumnWidth?: number;
  public Searcheable: boolean;


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
