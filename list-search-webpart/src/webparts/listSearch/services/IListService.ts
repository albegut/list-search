import { IListSearchListQuery } from "../model/ListSearchQuery";

export default interface IListService {
  getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string, sitePropertyValue: string, rowLimit: number): Promise<Array<any>>;
  getSiteListsTitle(): Promise<Array<any>>;
  getListFieldsTitle(listTitle: string): Promise<Array<any>>;
}
