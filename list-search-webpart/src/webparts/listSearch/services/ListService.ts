import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import IListService from "./IListService";
import { IListSearchListQuery } from "../model/ListSearchQuery";
import { ICamlQuery } from "@pnp/sp/lists";
import { IView } from "@pnp/sp/views";

export default class ListService implements IListService {

  constructor(siteUrl: string) {
    sp.setup({
      sp: {
        baseUrl: siteUrl
      },
    });
  }

  public async getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string, sitePropertyValue: string): Promise<Array<any>> {
    try {
      let items: any = undefined;
      if (listQueryOptions.camlQuery) {
        items = await this.getListItemsByCamlQuery(listQueryOptions.list, listQueryOptions.camlQuery);
      }
      else {
        if (listQueryOptions.viewName) {
          let viewInfo: any = await sp.web.lists.getByTitle(listQueryOptions.list).views.getByTitle(listQueryOptions.viewName).get();
          items = await this.getListItemsByCamlQuery(listQueryOptions.list, viewInfo.ViewQuery.toString())
        }
        else {
          let viewFields: string[] = listQueryOptions.fields.map(field => { return field.originalField });
          items = await sp.web.lists.getByTitle(listQueryOptions.list).items.select(viewFields.join(',')).get();
        }

      }
      let mappedItems = items.map(i => {
        listQueryOptions.fields.map(field => {
          i[field.newField] = i[field.originalField];
          delete i[field.originalField];
        });
        if (listPropertyName) {
          i[listPropertyName] = listQueryOptions.list
        }
        if (sitePropertyName) {
          i[sitePropertyName] = sitePropertyValue;
        }
        return i;
      })
      return mappedItems;
    } catch (error) {
      return Promise.reject(error);
    }
  }

  private async getListItemsByCamlQuery(listName: string, camlQuery: string): Promise<Array<any>> {
    try {
      const caml: ICamlQuery = {
        ViewXml: camlQuery,
      };
      return await sp.web.lists.getByTitle(listName).getItemsByCAMLQuery(caml);
    } catch (error) {
      return Promise.reject(error);
    }
  }
}
