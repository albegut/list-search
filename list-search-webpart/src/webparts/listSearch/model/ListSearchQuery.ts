export interface IListSearchListQuery {
  list: string;
  camlQuery?: string;
  viewName?: string;
  fields: Array<{ originalField: string, newField: string }>;
}
