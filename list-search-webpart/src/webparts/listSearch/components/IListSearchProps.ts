import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListFieldData, IListData, IDisplayFieldData } from "../model/IListConfigProps";
import { IPropertyFieldSite, } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { IReadonlyTheme } from '@microsoft/sp-component-base';


export interface IListSearchProps {
  Context: WebPartContext;
  displayFieldsCollectionData: Array<IDisplayFieldData>
  fieldsCollectionData: Array<IListFieldData>;
  listsCollectionData: Array<IListData>;
  ShowListName: boolean;
  ListNameTitle: string;
  ShowSite: boolean;
  SiteNameTitle: string;
  SiteNamePropertyToShow: string;
  GeneralFilter: boolean;
  GeneralFilterPlaceHolderText: string;
  GeneralSearcheableFields: Array<IDisplayFieldData>;
  IndividualColumnFilter: boolean;
  IndividualFilterPosition: string[];
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
  UseLocalStorage: boolean;
  minutesToCache: number;
}
