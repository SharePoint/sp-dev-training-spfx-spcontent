// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import {
  ButtonClickedCallback,
  ICountryListItem
} from '../../../models';

export interface ISpFxHttpClientDemoProps {
  spListItems: ICountryListItem[];
  onGetListItems?: ButtonClickedCallback;
  onAddListItem?: ButtonClickedCallback;
  onUpdateListItem?: ButtonClickedCallback;
  onDeleteListItem?: ButtonClickedCallback;
}
