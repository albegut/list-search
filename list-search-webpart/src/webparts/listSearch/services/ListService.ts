import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import IListService from "./IListService";

export default class ListService implements IListService {

    constructor(context: WebPartContext) {
        sp.setup({
            spfxContext: context
        });
    }

    public async getListItems(listName:string, fields: Array<string>, orderBy:string, asc: boolean): Promise<Array<any>> {
        try {
            //return sp.web.lists.getByTitle(listName).items.select(fields.join(',')).expand('Cobertura').orderBy("Orden", true).get();
            return sp.web.lists.getByTitle(listName).items.select(fields.join(',')).orderBy(orderBy, asc).get();
        } catch (error) {
            return Promise.reject(error);
        }
    }
}