// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as React from 'react';
import styles from './SpFxHttpClientDemo.module.scss';
import { ISpFxHttpClientDemoProps } from './ISpFxHttpClientDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpFxHttpClientDemo extends React.Component<ISpFxHttpClientDemoProps, {}> {
  public render(): React.ReactElement<ISpFxHttpClientDemoProps> {
    return (
      <div className={styles.spFxHttpClientDemo}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p className={styles.title}>SharePoint Content!</p>
              <a href="#" className={styles.button} onClick={this.onGetListItemsClicked}>
                <span className={styles.label}>Get Counties</span>
              </a>
              <a href="#" className={ styles.button } onClick={ this.onAddListItemClicked }>
                <span className={ styles.label }>Add List Item</span>
              </a>
              <a href="#" className={ styles.button } onClick={ this.onUpdateListItemClicked }>
                <span className={ styles.label }>Update List Item</span>
              </a>
              <a href="#" className={ styles.button } onClick={ this.onDeleteListItemClicked }>
                <span className={ styles.label }>Delete List Item</span>
              </a>              
            </div>
          </div>

          <div className={styles.row}>
            <ul className={styles.list}>
              {this.props.spListItems &&
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

  private onAddListItemClicked = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onAddListItem();
  }
  
  private onUpdateListItemClicked = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onUpdateListItem();
  }
  
  private onDeleteListItemClicked = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onDeleteListItem();
  }  
}
