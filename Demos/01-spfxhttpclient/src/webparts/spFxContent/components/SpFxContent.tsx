import * as React from 'react';
import styles from './SpFxContent.module.scss';
import { ISpFxContentProps } from './ISpFxContentProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpFxContent extends React.Component<ISpFxContentProps, {}> {
  public render(): React.ReactElement<ISpFxContentProps> {
    return (
      <div className={ styles.spFxContent }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p className={ styles.title }>SharePoint Content!</p>
              <a href="#" className={ styles.button } onClick={ this.onGetListItemsClicked }>
                <span className={ styles.label }>Get Counties</span>
              </a>
            </div>
          </div>
          
          <div className={ styles.row }>
            <ul className={ styles.list }>
              { this.props.spListItems &&
                this.props.spListItems.map((list) =>
                  <li key={list.Id} className={styles.item}>
                    <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
                  </li>
                )
              }
            </ul>
          </div>

        </div>
      </div>
    );
  }

  private onGetListItemsClicked = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();

    this.props.onGetListItems();
  }
}
