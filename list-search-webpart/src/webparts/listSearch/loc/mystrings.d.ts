declare interface IListSearchWebPartStrings {
  PropertyPaneDescription: string;
  GeneralPropertiesGroup: string;
  SourceSelectorGroup: string;
  SitesSelector:string;
  ListSelector:string;
  ListSelectorLabel:string;
  ListSelectorPanelHeader:string;
  ListFieldLabel: string;

}

declare module 'ListSearchWebPartStrings' {
  const strings: IListSearchWebPartStrings;
  export = strings;
}
