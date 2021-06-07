import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface ISpfxPnpRichtextWebPartProps {
    AzureSubscriptionKey: string;
    ServiceName: string;
}
export default class SpfxPnpRichtextWebPart extends BaseClientSideWebPart<ISpfxPnpRichtextWebPartProps> {
    protected onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SpfxPnpRichtextWebPart.d.ts.map