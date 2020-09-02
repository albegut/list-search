import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListFieldData, IListData } from "../model/IListConfigProps";
import { IPropertyFieldSite, } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { IReadonlyTheme } from '@microsoft/sp-component-base';


export interface IListSearchProps {
  Context: WebPartContext;
  collectionData: Array<IListFieldData>;
  ListscollectionData : Array<IListData>;
  ShowListName: boolean;
  ListNameTitle: string;
  ListNameOrder: number;
  ShowSite: boolean;
  SiteNameTitle: string;
  SiteNameOrder: number;
  SiteNameSearcheable: boolean;
  SiteNamePropertyToShow: string;
  GeneralFilter: boolean;
  GeneralFilterPlaceHolderText: string;
  GeneralSearcheableFields: Array<IListFieldData>;
  IndividualColumnFilter: boolean;
  ShowClearAllFilters: boolean;
  ClearAllFiltersBtnColor: string;
  ClearAllFiltersBtnText: string;
  Sites: IPropertyFieldSite[];
  ShowItemCount: boolean;
  ItemCountText: string;
  ItemLimit: number;
  ShowPagination: boolean;
  ItemsInPage: number;
  themeVariant: IReadonlyTheme | undefined;
}
