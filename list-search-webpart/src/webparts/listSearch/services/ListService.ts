import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import IListService from "./IListService";
import { IListSearchListQuery } from "../model/ListSearchQuery";
import { ICamlQuery } from "@pnp/sp/lists";
import { ICamlQueryXml } from '../model/ICamlQueryXml';
import XMLParser from 'react-xml-parser';

export default class ListService implements IListService {

  constructor(siteUrl: string) {
    sp.setup({
      sp: {
        baseUrl: siteUrl
      },
    });
  }

  public async getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string, sitePropertyValue: string, rowLimit: number): Promise<Array<any>> {
    try {
      let viewFields: string[] = listQueryOptions.fields.map(field => { return field.originalField; });
      let items: any = undefined;
      if (listQueryOptions.camlQuery) {
        let query = this.getCamlQueryWithViewFieldsAndRowLimit(listQueryOptions.camlQuery, viewFields, rowLimit);
        items = await this.getListItemsByCamlQuery(listQueryOptions.list, query);
      }
      else {
        if (listQueryOptions.viewName) {
          let viewInfo: any = await sp.web.lists.getByTitle(listQueryOptions.list).views.getByTitle(listQueryOptions.viewName).select("ViewQuery").get();
          let query = this.getCamlQueryWithViewFieldsAndRowLimit(`<View><Query>${viewInfo.ViewQuery}</Query></View>`, viewFields, rowLimit);
          items = await this.getListItemsByCamlQuery(listQueryOptions.list, query);
        }
        else {

          if (rowLimit) {
            items = await sp.web.lists.getByTitle(listQueryOptions.list).items.top(rowLimit).select(viewFields.join(',')).get();
          }
          else {
            items = await sp.web.lists.getByTitle(listQueryOptions.list).items.select(viewFields.join(',')).get();
          }
        }

      }
      let mappedItems = items.map(i => {
        listQueryOptions.fields.map(field => {
          i[field.newField] = i[field.originalField];
          delete i[field.originalField];
        });
        if (listPropertyName) {
          i[listPropertyName] = listQueryOptions.list;
        }
        if (sitePropertyName) {
          i[sitePropertyName] = sitePropertyValue;
        }
        return i;
      });
      return mappedItems;
    } catch (error) {
      return Promise.reject(error);
    }
  }

  public async getSiteListsTitle(): Promise<Array<any>> {
    try {
      return sp.web.lists.filter('Hidden eq false').select('Title').get();
    } catch (error) {
      return Promise.reject(error);
    }
  }

  public async getListFieldsTitle(listTitle: string): Promise<Array<any>> {
    try {
      return sp.web.lists.getByTitle(listTitle).fields.select('Title').get();
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

  private getCamlQueryWithViewFieldsAndRowLimit(camlQuery: string, viewFields: Array<string>, rowLimit: number): string {
    try {
      let XmlParser = new XMLParser();
      let xml: ICamlQueryXml = XmlParser.parseFromString(camlQuery);

      let rowLimitXml: ICamlQueryXml = { name: "RowLimit", value: rowLimit ? rowLimit.toString() : "0", attributes: undefined, children: [] };

      let viewFieldsChildren: ICamlQueryXml[] = viewFields.map(viewField => {
        return { name: "FieldRef", attributes: { Name: viewField }, value: "", children: [] };
      });
      let viewFieldsXml: ICamlQueryXml = { name: "ViewFields", value: "", children: viewFieldsChildren, attributes: undefined };

      let queryXml: ICamlQueryXml;
      let hasPrevRowLimit: boolean = false;
      xml.children.map(child => {
        if (child.name == "Query") {
          queryXml = child;
        }

        if (child.name == "RowLimit") { //If the user set a camlquery with row limit or the view has row limit, it is not override
          rowLimitXml = child;
        }
      });

      if (queryXml) {
        xml.children = [viewFieldsXml, rowLimitXml, queryXml];
      }

      return XmlParser.toString(xml);
    } catch (error) {
      console.error(`getCamlQueryWithViewFieldsAndRowLimit -> ${error.message}`);
      return `getCamlQueryWithViewFieldsAndRowLimit -> ${error.message}`;
    }

  }
}
