import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListConfigProps } from "../model/IListConfigProps";
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

export interface IListSearchProps {
  Context: WebPartContext;
  collectionData: Array<IListConfigProps>;
  ShowListName: boolean;
  ListNameTitle: string;
  ListNameOrder: number;
  ShowSite: boolean;
  SiteNameTitle: string;
  SiteNameOrder: number;
}
