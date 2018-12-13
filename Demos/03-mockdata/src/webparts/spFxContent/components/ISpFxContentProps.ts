import { 
  ButtonClickedCallback,
  ICountryListItem 
} from '../../../models';

export interface ISpFxContentProps {
  spListItems: ICountryListItem[];
  onGetListItems?: ButtonClickedCallback;
  onAddListItem?: ButtonClickedCallback;
  onUpdateListItem?: ButtonClickedCallback;
  onDeleteListItem?: ButtonClickedCallback;
}
