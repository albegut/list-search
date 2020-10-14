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
import { IWeb, Web } from '@pnp/sp/webs';
import { IListField } from '../model/IListField';
import { SharePointType } from '../model/ISharePointFieldTypes';


export interface QueryHelperEntity {
  viewFields: string[];
  expandFields: string[];
}

export default class ListService implements IListService {
  private web: IWeb;
  private baseUrl: string;

  constructor(siteUrl: string) {
    sp.setup({
      sp: {
        headers: {
          Accept: 'application/json;odata=nometadata'
        },
        baseUrl: siteUrl
      },
    });
    this.web = Web(siteUrl);
    this.baseUrl = siteUrl;
  }

  private GetViewFieldsWithId(listQueryOptions: IListSearchListQuery): QueryHelperEntity {
    let result: QueryHelperEntity = { expandFields: [], viewFields: [] };
    listQueryOptions.fields.map(field => {
      switch (field.fieldType) {
        case SharePointType.User:
          result.viewFields.push(`${field.originalField}/EMail`);
          result.viewFields.push(`${field.originalField}/Name`);
          result.expandFields.push(`${field.originalField}`);
          break;
        case SharePointType.UserMulti:
          result.viewFields.push(`${field.originalField}/Title`);
          result.viewFields.push(`${field.originalField}/Name`);
          result.expandFields.push(`${field.originalField}`);
          break;
        case SharePointType.Lookup:
        case SharePointType.LookupMulti:
          {
            result.viewFields.push(`${field.originalField}Id`);
            break;
          }
        default:
          {
            result.viewFields.push(field.originalField);
            break;
          }
      }
    });

    return result;
  }

  private GetItemValue(item: any, field: any): any {

    switch (field.fieldType) {
      case SharePointType.User:
      case SharePointType.UserMulti:
        item[field.newField] = item[field.originalField];
        delete item[field.originalField];
        break;
      case SharePointType.Lookup:
      case SharePointType.LookupMulti:
        {
          item[field.newField] = item[`${field.originalField}Id`];
          delete item[`${field.originalField}Id`];
          break;
        }
      default:
        {
          item[field.newField] = item[field.originalField];
          delete item[field.originalField];
          break;
        }
    }


    return item;
  }

  public async getListItems(listQueryOptions: IListSearchListQuery, listPropertyName: string, sitePropertyName: string, sitePropertyValue: string, rowLimit: number): Promise<Array<any>> {
    try {
      let queryConfig: QueryHelperEntity = this.GetViewFieldsWithId(listQueryOptions);
      queryConfig.viewFields.push("Id");
      let items: any = undefined;
      if (listQueryOptions.camlQuery) {
        let query = this.getCamlQueryWithViewFieldsAndRowLimit(listQueryOptions.camlQuery, queryConfig, rowLimit);
        items = await this.getListItemsByCamlQuery(listQueryOptions.list, query);
      }
      else {
        if (listQueryOptions.viewName) {
          let viewInfo: any = await this.web.lists.getByTitle(listQueryOptions.list).views.getByTitle(listQueryOptions.viewName).select("ViewQuery").get();
          let query = this.getCamlQueryWithViewFieldsAndRowLimit(`<View><Query>${viewInfo.ViewQuery}</Query></View>`, queryConfig, rowLimit);
          items = await this.getListItemsByCamlQuery(listQueryOptions.list, query);
        }
        else {

          if (rowLimit) {
            if (queryConfig.expandFields && queryConfig.expandFields.length > 0) {
              items = await this.web.lists.getByTitle(listQueryOptions.list).items.select(queryConfig.viewFields.join(',')).expand(queryConfig.expandFields.join(',')).get();
            }
            else {
              items = await this.web.lists.getByTitle(listQueryOptions.list).items.top(rowLimit).select(queryConfig.viewFields.join(',')).get();
            }
          }
          else {
            if (queryConfig.expandFields && queryConfig.expandFields.length > 0) {
              items = await this.web.lists.getByTitle(listQueryOptions.list).items.select(queryConfig.viewFields.join(',')).expand(queryConfig.expandFields.join(',')).get();
            }
            else {
              items = await this.web.lists.getByTitle(listQueryOptions.list).items.select(queryConfig.viewFields.join(',')).get();
            }
          }
        }

      }
      let mappedItems = items.map(i => {
        listQueryOptions.fields.map(field => {
          i = this.GetItemValue(i, field);
        });
        i["SiteUrl"] = this.baseUrl;
        i["ListName"] = listQueryOptions.list;
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

  public async getListItemById(listName: string, itemId: number): Promise<any> {
    try {
      return this.web.lists.getByTitle(listName).items.getById(itemId).get();
    } catch (error) {
      return Promise.reject(error);
    }
  }

  public async getSiteListsTitle(): Promise<Array<any>> {
    try {
      return this.web.lists.filter('Hidden eq false').select('Title').get();
    } catch (error) {
      return Promise.reject(error);
    }
  }

  public async getListFields(listTitle: string): Promise<Array<IListField>> {
    try {
      return this.web.lists.getByTitle(listTitle).fields.filter(`substringof('${encodeURIComponent("OData__")}',EntityPropertyName) eq false`).select('EntityPropertyName,Title,InternalName,TypeAsString').get();
    } catch (error) {
      return Promise.reject(error);
    }
  }

  private async getListItemsByCamlQuery(listName: string, camlQuery: string): Promise<Array<any>> {
    try {
      const caml: ICamlQuery = {
        ViewXml: camlQuery,
      };
      return await this.web.lists.getByTitle(listName).getItemsByCAMLQuery(caml);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  private getCamlQueryWithViewFieldsAndRowLimit(camlQuery: string, queryConfig: QueryHelperEntity, rowLimit: number): string {
    try {
      let XmlParser = new XMLParser();
      let xml: ICamlQueryXml = XmlParser.parseFromString(camlQuery);

      let rowLimitXml: ICamlQueryXml = { name: "RowLimit", value: rowLimit ? rowLimit.toString() : "0", attributes: undefined, children: [] };

      let viewFieldsChildren: ICamlQueryXml[] = queryConfig.viewFields.map(viewField => {
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
      return `getCamlQueryWithViewFieldsAndRowLimit -> ${error.message}`;
    }

  }
}
