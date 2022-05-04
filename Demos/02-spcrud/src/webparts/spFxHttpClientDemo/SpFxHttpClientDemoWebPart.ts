// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpFxHttpClientDemoWebPartStrings';
import SpFxHttpClientDemo from './components/SpFxHttpClientDemo';
import { ISpFxHttpClientDemoProps } from './components/ISpFxHttpClientDemoProps';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ICountryListItem } from '../../models';

export interface ISpFxHttpClientDemoWebPartProps {
  description: string;
}

export default class SpFxHttpClientDemoWebPart extends BaseClientSideWebPart<ISpFxHttpClientDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _countries: ICountryListItem[] = [];

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ISpFxHttpClientDemoProps> = React.createElement(
      SpFxHttpClientDemo,
      {
        spListItems: this._countries,
        onGetListItems: this._onGetListItems,
        onAddListItem: this._onAddListItem,
        onUpdateListItem: this._onUpdateListItem,
        onDeleteListItem: this._onDeleteListItem,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _onGetListItems = (): void => {
    this._getListItems()
      .then(response => {
        this._countries = response;
        this.render();
      });
  }

  private _onAddListItem = (): void => {
    this._addListItem()
      .then(() => {
        this._getListItems()
          .then(response => {
            this._countries = response;
            this.render();
          });
      });
  }
  
  private _onUpdateListItem = (): void => {
    this._updateListItem()
      .then(() => {
        this._getListItems()
          .then(response => {
            this._countries = response;
            this.render();
          });
      });
  }
  
  private _onDeleteListItem = (): void => {
    this._deleteListItem()
      .then(() => {
        this._getListItems()
          .then(response => {
            this._countries = response;
            this.render();
          });
      });
  }

  private _getListItems(): Promise<ICountryListItem[]> {
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')/items?$select=Id,Title`,
      SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value;
      }) as Promise<ICountryListItem[]>;
  }

  private _getItemEntityType(): Promise<string> {
    return this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.ListItemEntityTypeFullName;
      }) as Promise<string>;
  }
  
  private _addListItem(): Promise<SPHttpClientResponse> {
    return this._getItemEntityType()
      .then(spEntityType => {
        const request: any = {};
        request.body = JSON.stringify({
          Title: new Date().toUTCString(),
          '@odata.type': spEntityType
        });
  
        return this.context.spHttpClient.post(
          this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')/items`,
          SPHttpClient.configurations.v1,
          request);
        }
      ) ;
  }

  private _updateListItem(): Promise<SPHttpClientResponse> {
    // get the first item
    return this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')/items?$select=Id,Title&$filter=Title eq 'United States'`,
        SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value[0];
      })
      .then((listItem: ICountryListItem) => {
        // update item
        listItem.Title = 'USA';
        // save it
        const request: any = {};
        request.headers = {
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': (listItem as any)['@odata.etag']
        };
        request.body = JSON.stringify(listItem);
  
        return this.context.spHttpClient.post(
          this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')/items(${listItem.Id})`,
          SPHttpClient.configurations.v1,
          request);
      });
  }

  private _deleteListItem(): Promise<SPHttpClientResponse> {
    // get the last item
    return this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')/items?$select=Id,Title&$orderby=ID desc&$top=1`,
        SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value[0];
      })
      .then((listItem: ICountryListItem) => {
        const request: any = {};
        request.headers = {
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*'
        };
        request.body = JSON.stringify(listItem);
  
        return this.context.spHttpClient.post(
          this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Countries')/items(${listItem.Id})`,
          SPHttpClient.configurations.v1,
          request);
      });
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
