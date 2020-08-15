import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListConfigProps } from "../model/IListConfigProps";

export interface IListSearchProps{
    ListName:string;
    Context: WebPartContext;
    collectionData: Array<IListConfigProps>;
    ShowListName : boolean;
    ListNameTitle: string;
}
