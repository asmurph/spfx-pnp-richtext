import * as React from 'react';
import styles from './SpfxPnpRichtext.module.scss';
import { ISpfxPnpRichtextProps } from './ISpfxPnpRichtextProps';
import { ISpfxPnpRichtextState } from './ISpfxPnpRichtextState';
import { sp } from "@pnp/sp";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { ISpfxAzureTranslatorState } from './ISpfxAzureTranslatorState';
import { IDropdownOption, Dropdown } from 'office-ui-fabric-react';
import $ from "jquery";
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';

const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const stackTokens = { childrenGap: 50 };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};
const smallcolumnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 180 } },
};

export default class SpfxPnpRichtext extends React.Component<ISpfxPnpRichtextProps, ISpfxAzureTranslatorState> {
  constructor(props: ISpfxPnpRichtextProps, state: ISpfxAzureTranslatorState,) {
    super(props);    
    this.state = ({ toLanguage: '', Title: '', Description: '', richtext:'', langarr: [] });
    this._getSupportedLangualge();
    this._getData();
  }


  public render(): React.ReactElement<ISpfxPnpRichtextProps> {
  
    return (
      <div className={styles.spfxPnpRichtext}>
         <Stack horizontal tokens={stackTokens} styles={stackStyles}>
         <Stack {...columnProps}>
         { this.state.Description }
          </Stack>
          <Stack {...smallcolumnProps}>
            <Dropdown
              placeholder="Select a language"
              label="Select Language"
              options={this.state.langarr}
              onChanged={(value) => { this.setState({ toLanguage: value.key.toString() }); this._translate(); }}
            />
          </Stack>
          <Stack {...columnProps}>
            <label>{this.state.richtext}</label>
          </Stack>
        </Stack>
      </div>
    );
  }
   public async _getData()
   {
    try {
      const richTextItem = await sp.web.lists.getByTitle('ListTest').items.getById(1)
        .select("ID", "Title", "Description")
        .get();   
   

      this.setState({     
        Title: richTextItem.Title,
        Description: richTextItem.Description,    
      });
    }
    catch (error) {
      this.setState({
        MessageText: "Exception reading item",
        MessageType: MessageBarType.error
      });

      return Promise.reject(error);
    }
   }



  private async _getSupportedLangualge() {
    $.get({
      url: 'https://api.cognitive.microsofttranslator.com/languages?api-version=3.0&scope=translation'
    })
      .done((languages: any): void => {
        let droparr: IDropdownOption[] = [];
        let langobjs = languages.translation;
        for (var key in langobjs) {
          if (langobjs.hasOwnProperty(key)) {
            droparr.push({ key: key, text: langobjs[key].name });
          }
        }
        this.setState({ langarr: droparr });
      });
  }
  private async _translate() {
    $.post({
      url: 'https://' + this.props.ServiceName + '.cognitiveservices.azure.com/sts/v1.0/issueToken',
      headers: {
        'Ocp-Apim-Subscription-Key': this.props.AzureSubscriptionKey,
        'Authorization': this.props.ServiceName + '.cognitiveservices.azure.com'
      }
    })
      .done((tocken: any): void => {
        $.post({
          url: 'https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to=' + this.state.toLanguage,
          headers: {
            'Ocp-Apim-Subscription-Key': this.props.AzureSubscriptionKey,
            'Authorization': 'Bearer ' + tocken,
            'Content-Type': 'application/json'
          },
          data: JSON.stringify([{ "Text": this.state.Description }])
        })
          .done((result: any): void => {
            console.log(result);
            this.setState({ Description: result[0].translations[0].text });
          });

      });
  }

}
