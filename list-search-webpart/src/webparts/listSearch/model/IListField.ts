
import { SharePointFieldTypes } from './IListConfigProps'

export interface IListField {
  Title: string;
  InternalName: string;
  TypeAsString: SharePointFieldTypes;
}
