import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { MessageBarType } from 'office-ui-fabric-react';

export interface ISpfxAzureTranslatorState {
  toLanguage: string;
  langarr: IDropdownOption[];
  richtext: string;
  Title: string;
  Description?: any;
  MessageText?: string;
  MessageType?: MessageBarType;
}
