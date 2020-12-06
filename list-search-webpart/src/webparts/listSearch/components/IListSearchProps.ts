import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IMappingFieldData, IListData, IDetailListFieldData, ICompleteModalData, IRedirectData } from "../model/IListConfigProps";
import { IPropertyFieldSite, } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SharePointType } from "../model/ISharePointFieldTypes";


export interface IListSearchProps {
  Context: WebPartContext;
  detailListFieldsCollectionData: Array<IDetailListFieldData>;
  mappingFieldsCollectionData: Array<IMappingFieldData>;
  listsCollectionData: Array<IListData>;
  ShowListName: boolean;
  ShowFileIcon: boolean;
  ListNameTitle: string;
  ShowSite: boolean;
  SiteNameTitle: string;
  SiteNamePropertyToShow: string;
  GeneralFilter: boolean;
  GeneralFilterPlaceHolderText: string;
  GeneralSearcheableFields: Array<IDetailListFieldData>;
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
  clickEnabled: boolean;
  clickIsSimpleModal: boolean;
  clickIsCompleteModal: boolean;
  clickIsRedirect: boolean;
  clickIsDynamicData: boolean;
  completeModalFields: Array<ICompleteModalData>;
  redirectData: Array<IRedirectData>;
  onRedirectIdQuery: string;
  onSelectedItem: any;
  oneClickOption: boolean;
  groupByField: string;
  groupByFieldType: SharePointType;
  AnyCamlQuery: boolean;
}
