// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as React from 'react';
import styles from './SpFxHttpClientDemo.module.scss';
import { ISpFxHttpClientDemoProps } from './ISpFxHttpClientDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpFxHttpClientDemo extends React.Component<ISpFxHttpClientDemoProps, {}> {
  public render(): React.ReactElement<ISpFxHttpClientDemoProps> {
    const {
      spListItems,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.spFxHttpClientDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
        </div>
        <div className={styles.buttons}>
          <button type="button" onClick={this.onGetListItemsClicked}>Get Countries</button>
          <button type="button" onClick={this.onAddListItemClicked}>Add List Item</button>
          <button type="button" onClick={this.onUpdateListItemClicked}>Update List Item</button>
          <button type="button" onClick={this.onDeleteListItemClicked}>Delete List Item</button>
        </div>
        <div>
          <ul>
            {spListItems && spListItems.map((list) =>
              <li key={list.Id}>
                <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
              </li>
            )
            }
          </ul>
        </div>
      </section>
    );
  }

  private onGetListItemsClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();

    this.props.onGetListItems();
  }

  private onAddListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();

    this.props.onAddListItem();
  }

  private onUpdateListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();

    this.props.onUpdateListItem();
  }

  private onDeleteListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();

    this.props.onDeleteListItem();
  }
}
