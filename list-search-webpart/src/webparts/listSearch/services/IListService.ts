import { IListSearchListQuery } from "../model/ListSearchQuery";
import { IListField } from "../model/IListField";

export default interface IListService {
  getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string, sitePropertyValue: string, rowLimit: number): Promise<Array<any>>;
  getSiteListsTitle(): Promise<Array<any>>;
  getListFields(listTitle: string): Promise<Array<IListField>>;
}
