import * as React from 'react';
import { Dropdown } from 'office-ui-fabric-react/lib/components/Dropdown';
import { ICustomCollectionField } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { TextField } from 'office-ui-fabric-react/lib/components/TextField';
import { IMappingFieldData, IListData, ICustomOption } from '../model/IListConfigProps';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { IListField } from '../model/IListField';
import styles from '../ListSearchWebPart.module.scss';



export default class CustomCollectionDataField {
  private static getCustomCollectionDropDown(options: IPropertyPaneDropdownOption[], field: ICustomCollectionField, row: any, updateFunction: any, errorFunction?: any, customOnchangeFunction?: any): JSX.Element {
    return (<Dropdown placeholder={field.placeholder || field.title}
      options={options.sort((a, b) => { return a.text.localeCompare(b.text); })}
      selectedKey={row[field.id] || null}
      required={field.required}
      onChange={(evt, option, index) => customOnchangeFunction ? customOnchangeFunction(row, field.id, option, updateFunction, errorFunction) : updateFunction(field.id, option.key)}
      onRenderOption={field.onRenderOption}
      className="PropertyFieldCollectionData__panel__dropdown-field" />);
  }

  public static getListPickerBySiteOptions(possibleOptions: Array<IListData>, field: ICustomCollectionField, row: IMappingFieldData, updateFunction: any): JSX.Element {
    let currentOptions = [];
    possibleOptions.filter(option => {
      if (row.SiteCollectionSource && option.SiteCollectionSource == row.SiteCollectionSource) {
        currentOptions.push({
          key: option.ListSourceField,
          text: option.ListSourceField
        });
      }
    });
    return this.getCustomCollectionDropDown(currentOptions, field, row, updateFunction);
  }

  public static getPickerByStringOptions(possibleOptions: Array<string>, field: ICustomCollectionField, row: IListData, updateFunction: any, customOnChange: any): JSX.Element {
    let options = [];
    if (possibleOptions) {
      options = possibleOptions.map(option => { return { key: option, text: option }; });
    }
    return this.getCustomCollectionDropDown(options, field, row, updateFunction, null, customOnChange);
  }

  public static getFieldPickerByList(possibleOptions: Array<IListField>, field: ICustomCollectionField, row: IListData, updateFunction: any, customOnchangeFunction?: any, customOptions?: Array<ICustomOption>): JSX.Element {
    let options = [];
    if (possibleOptions) {
      options = possibleOptions.map(option => { return { key: option.InternalName, text: option.Title, FieldType: option.TypeAsString }; });
    }
    if (customOptions) {
      customOptions.map(option => {
        options.push({
          key: option.Key,
          text: option.Option,
          FieldType: option.CustomData
        });
      });
    }
    return this.getCustomCollectionDropDown(options, field, row, updateFunction, null, customOnchangeFunction);
  }

  public static getDisableTextField(field: ICustomCollectionField, item: any, updateFunction: any): JSX.Element {
    return <TextField placeholder={field.placeholder || field.title}
      className={styles.collectionDataField}
      value={item[field.id] ? item[field.id] : ""}
      required={field.required}
      disabled={true}
      onChange={(value) => updateFunction(field.id, value)}
      deferredValidationTime={field.deferredValidationTime || field.deferredValidationTime >= 0 ? field.deferredValidationTime : 200}
      inputClassName="PropertyFieldCollectionData__panel__string-field" />;
  }
}
