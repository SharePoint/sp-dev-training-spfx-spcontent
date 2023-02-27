import {
  ButtonClickedCallback,
  ICountryListItem
} from '../../../models';

export interface ISpFxHttpClientDemoProps {
  spListItems: ICountryListItem[];
  onGetListItems?: ButtonClickedCallback;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  onAddListItem?: ButtonClickedCallback;
  onUpdateListItem?: ButtonClickedCallback;
  onDeleteListItem?: ButtonClickedCallback;
}
