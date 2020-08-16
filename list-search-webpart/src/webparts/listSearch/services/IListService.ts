import { IListSearchListQuery } from "../model/ListSearchQuery";

export default interface IListService {
  getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string): Promise<Array<any>>
}
