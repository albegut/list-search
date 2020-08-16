import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import IListService from "./IListService";
import { IListSearchListQuery } from "../model/ListSearchQuery";

export default class ListService implements IListService {

  private _siteUrl: string;

  constructor(siteUrl: string) {
    this._siteUrl = siteUrl;
    sp.setup({
      sp: {
        baseUrl: siteUrl
      },
    });
  }

  public async getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string): Promise<Array<any>> {
    try {
      let viewFields: string[] = listQueryOptions.fields.map(field => { return field.originalField });
      let items = await sp.web.lists.getByTitle(listQueryOptions.list).items.select(viewFields.join(',')).get();
      let mappedItems = items.map(i => {
        listQueryOptions.fields.map(field => {
          i[field.newField] = i[field.originalField];
          delete i[field.originalField];
        });
        if (listPropertyName) {
          i[listPropertyName] = listQueryOptions.list
        }
        if (sitePropertyName) {
          i[sitePropertyName] = this._siteUrl;
        }
        return i;
      })
      return mappedItems;
    } catch (error) {
      return Promise.reject(error);
    }
  }
}
