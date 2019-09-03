// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxHttpClientDemoWebPartStrings';
import SpFxHttpClientDemo from './components/SpFxHttpClientDemo';
import { ISpFxHttpClientDemoProps } from './components/ISpFxHttpClientDemoProps';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ICountryListItem } from '../../models';

export interface ISpFxHttpClientDemoWebPartProps {
  description: string;
}

export default class SpFxHttpClientDemoWebPart extends BaseClientSideWebPart<ISpFxHttpClientDemoWebPartProps> {
  private _countries: ICountryListItem[] = [];

  public render(): void {
    const element: React.ReactElement<ISpFxHttpClientDemoProps > = React.createElement(
      SpFxHttpClientDemo,
      {
        spListItems: this._countries,
        onGetListItems: this._onGetListItems
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
