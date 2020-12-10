import { SiteList } from "./IListConfigProps";
import { SharePointType } from "./ISharePointFieldTypes";

export interface IListSearchListQuery {
  list: SiteList;
  camlQuery?: string;
  viewName?: string;
  fields: Array<{ originalField: string, newField: string, fieldType: SharePointType }>;
}
