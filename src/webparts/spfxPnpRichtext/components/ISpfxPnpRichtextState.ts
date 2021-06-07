import { MessageBarType } from 'office-ui-fabric-react';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface ISpfxPnpRichtextState {
  ID?: number;
  Title: string;
  Description?: any;
  editorState?: any;
  MessageText?: string;
  MessageType?: MessageBarType;
}
