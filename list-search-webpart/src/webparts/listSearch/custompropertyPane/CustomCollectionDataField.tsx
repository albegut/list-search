import * as React from 'react';
import { Dropdown } from 'office-ui-fabric-react/lib/components/Dropdown';
import { ICustomCollectionField } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IListFieldData, IListData, ICustomOption } from '../model/IListConfigProps';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { IListField } from '../model/IListField';



export default class CustomCollectionDataField {
  private static getCustomCollectionDropDown(options: IPropertyPaneDropdownOption[], field: ICustomCollectionField, row: any, updateFunction: any, errorFunction?: any, customOnchangeFunction?: any): JSX.Element {
    return (<Dropdown placeholder={field.placeholder || field.title}
      options={options}
      selectedKey={row[field.id] || null}
      required={field.required}
      onChange={(evt, option, index) => customOnchangeFunction ? customOnchangeFunction(row, field.id, option.key, updateFunction, errorFunction) : updateFunction(field.id, option.key)}
      onRenderOption={field.onRenderOption}
      className="PropertyFieldCollectionData__panel__dropdown-field" />);
  }

  public static getListPickerBySiteOptions(possibleOptions: Array<IListData>, field: ICustomCollectionField, row: IListFieldData, updateFunction: any): JSX.Element {
    let currentOptions = [];
    possibleOptions.filter(option => {
      if (row.SiteCollectionSource && option.SiteCollectionSource == row.SiteCollectionSource) {
        currentOptions.push({
          key: option.ListSourceField,
          text: option.ListSourceField
        });
      }
    });
    return this.getCustomCollectionDropDown(currentOptions.sort(), field, row, updateFunction);
  }

  public static getPickerByStringOptions(possibleOptions: Array<string>, field: ICustomCollectionField, row: IListData, updateFunction: any, customOnChange: any): JSX.Element {
    let options = [];
    if (possibleOptions) {
      options = possibleOptions.map(option => { return { key: option, text: option }; });
    }
    return this.getCustomCollectionDropDown(options.sort(), field, row, updateFunction, null, customOnChange);
  }

  public static getFieldPickerByList(possibleOptions: Array<IListField>, field: ICustomCollectionField, row: IListData, updateFunction: any, customOptions?: Array<ICustomOption>): JSX.Element {
    let options = [];
    if (possibleOptions) {
      options = possibleOptions.map(option => { return { key: option.InternalName, text: option.Title }; });
    }
    if (customOptions) {
      customOptions.map(option => {
        options.push({
          key: option.Key,
          text: option.Option
        });
      });
    }
    return this.getCustomCollectionDropDown(options.sort(), field, row, updateFunction);
  }
}
