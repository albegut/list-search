import * as React from 'react';
import styles from './ListSearchConsumerWebPart.module.scss';
import { IListSearchConsumerProps } from './IListSearchConsumerProps';

export default class ListSearchConsumer extends React.Component<IListSearchConsumerProps, {}> {
  public render(): React.ReactElement<IListSearchConsumerProps> {
    return (
      <div className={styles.listSearchConsumerWebPart}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>List search consumer webpart</span>
            </div>
            <div className={styles.column}>
              <div className={styles.description}>WebUrl:
                <p>{this.props.webUrl}</p>
              </div>
            </div>
            <div className={styles.column}>
              <div className={styles.description}>ListId:
                <p>${this.props.listId}</p>
              </div>
            </div>
            <div className={styles.column}>
              <div className={styles.description}>ItemId:
                <p>${this.props.itemId}</p>
              </div>
            </div>
          </div>
        </div>
      </div >
    );
  }
}
