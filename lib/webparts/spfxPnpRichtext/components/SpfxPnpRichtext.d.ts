import * as React from 'react';
import { ISpfxPnpRichtextProps } from './ISpfxPnpRichtextProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ISpfxAzureTranslatorState } from './ISpfxAzureTranslatorState';
export default class SpfxPnpRichtext extends React.Component<ISpfxPnpRichtextProps, ISpfxAzureTranslatorState> {
    constructor(props: ISpfxPnpRichtextProps, state: ISpfxAzureTranslatorState);
    render(): React.ReactElement<ISpfxPnpRichtextProps>;
    _getData(): Promise<any>;
    private _getSupportedLangualge;
    private _translate;
}
//# sourceMappingURL=SpfxPnpRichtext.d.ts.map