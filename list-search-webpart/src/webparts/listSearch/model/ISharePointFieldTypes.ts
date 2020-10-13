export class SharePointFieldTypes {

  public static GetSPFieldTypeByString(fieldTypeAsString: string): SharePointType {
    let result: SharePointType = SharePointType.Text;
    switch (fieldTypeAsString) {
      case 'Text': {
        result = SharePointType.Text;
        break;
      }
      case 'Note': {
        result = SharePointType.Note;
        break;
      }
      case 'Choice': {
        result = SharePointType.Choice;
        break;
      }
      case 'Integer': {
        result = SharePointType.Integer;
        break;
      }
      case 'Number': {
        result = SharePointType.Number;
        break;
      }
      case 'Money': {
        result = SharePointType.Money;
        break;
      }
      case 'DateTime': {
        result = SharePointType.DateTime;
        break;
      }
      case 'Lookup': {
        result = SharePointType.Lookup;
        break;
      }
      case 'LookupMulti': {
        result = SharePointType.LookupMulti;
        break;
      }
      case 'Boolean': {
        result = SharePointType.Boolean;
        break;
      }
      case 'User': {
        result = SharePointType.User;
        break;
      }
      case 'UserMulti': {
        result = SharePointType.UserMulti;
        break;
      }
      case 'Url': {
        result = SharePointType.Url;
        break;
      }
      case 'Computed': {
        result = SharePointType.Computed;
        break;
      }
      case 'Image': {
        result = SharePointType.Image;
        break;
      }
      case 'Taxonomy': {
        result = SharePointType.Taxonomy;
        break;
      }
      case 'Attachments': {
        result = SharePointType.Attachments;
        break;
      }
      case 'Counter': {
        result = SharePointType.Counter;
        break;
      }
      case 'ContentTypeId': {
        result = SharePointType.ContentTypeId;
        break;
      }
      case 'Guid': {
        result = SharePointType.Guid;
        break;
      }
    }

    return result;
  }

  public static GetSharePointTypesAsArray(): Array<string> {
    return Object.keys(SharePointType);
  }

}

export enum SharePointType {
  Text = "Text",
  Note = "Note",
  Choice = "Choice",
  Integer = "Integer",
  Number = "Number",
  Money = "Money",
  DateTime = "DateTime",
  Lookup = "Lookup",
  LookupMulti = "LookupMulti",
  Boolean = "Boolean",
  User = "User",
  UserMulti = "UserMulti",
  Url = "Url",
  Image = "Image",
  Taxonomy = "Taxonomy",
  TaxonomyMulti = "TaxonomyMulti",
  Computed = "Computed",
  Attachments = "Attachments",
  Counter = "Counter",
  ContentTypeId = "ContentTypeId",
  Guid = "Guid",
}
