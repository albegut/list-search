export interface IListSearchListQuery {
  list: string;
  fields: Array<{ originalField: string, newField: string }>;
}
