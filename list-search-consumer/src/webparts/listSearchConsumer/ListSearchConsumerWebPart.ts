import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  DynamicDataSharedDepth,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField
} from '@microsoft/sp-property-pane';
import {
  BaseClientSideWebPart,
  IWebPartPropertiesMetadata,
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListSearchConsumerWebPart.module.scss';
import * as strings from 'ListSearchConsumerWebPartStrings';
import { DynamicProperty } from '@microsoft/sp-component-base';


export interface IListSearchConsumerWebPartProps {
  description: string;
  webUrl:DynamicProperty<string>;
  listName: DynamicProperty<string>;
  itemId: DynamicProperty<number>;
}

export default class ListSearchConsumerWebPart extends BaseClientSideWebPart<IListSearchConsumerWebPartProps> {

  public render(): void {
    const webUrl: string | undefined = this.properties.webUrl.tryGetValue();
    const listName: string | undefined = this.properties.listName.tryGetValue();
    const itemId: number | undefined = this.properties.itemId.tryGetValue();
    this.domElement.innerHTML = `
      <div class="${styles.listSearchConsumer}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">List search consumer webpart</span>
              <p class="${styles.description}">WebUrl: ${webUrl}</p>
              <p class="${styles.description}">List Name: ${listName}</p>
              <p class="${styles.description}">ItemId: ${itemId}</p>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Select web Url',
                  fields: [
                    PropertyPaneDynamicField('webUrl', {
                      label: 'Web Url'
                    }),
                    PropertyPaneDynamicField('listName', {
                      label: 'List Name'
                    }),
                    PropertyPaneDynamicField('itemId', {
                      label: 'Item Id'
                    })
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
