export interface ICamlQueryXml {
  name: string;
  attributes: IViewField | undefined;
  value: string;
  children: Array<ICamlQueryXml>;
}

export interface IViewField {
  Name: string;
}

