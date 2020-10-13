import { SharePointType } from "./ISharePointFieldTypes";

export interface IListSearchListQuery {
  list: string;
  camlQuery?: string;
  viewName?: string;
  fields: Array<{ originalField: string, newField: string, fieldType: SharePointType }>;
}
