import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxContentWebPartStrings';
import SpFxContent from './components/SpFxContent';
import { ISpFxContentProps } from './components/ISpFxContentProps';

import { SPHttpClient } from '@microsoft/sp-http';
import { ICountryListItem } from '../../models';

export interface ISpFxContentWebPartProps {
  description: string;
}

export default class SpFxContentWebPart extends BaseClientSideWebPart<ISpFxContentWebPartProps> {
  private _countries: ICountryListItem[] = [];

  public render(): void {
    const element: React.ReactElement<ISpFxContentProps > = React.createElement(
      SpFxContent,
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
