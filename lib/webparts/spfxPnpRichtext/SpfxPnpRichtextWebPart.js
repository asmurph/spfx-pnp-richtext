var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import * as strings from 'SpfxPnpRichtextWebPartStrings';
import SpfxPnpRichtext from './components/SpfxPnpRichtext';
var SpfxPnpRichtextWebPart = /** @class */ (function (_super) {
    __extends(SpfxPnpRichtextWebPart, _super);
    function SpfxPnpRichtextWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    SpfxPnpRichtextWebPart.prototype.onInit = function () {
        sp.setup({
            spfxContext: this.context
        });
        return Promise.resolve();
    };
    SpfxPnpRichtextWebPart.prototype.render = function () {
        var element = React.createElement(SpfxPnpRichtext, {
            AzureSubscriptionKey: this.properties.AzureSubscriptionKey,
            ServiceName: this.properties.ServiceName
        });
        ReactDom.render(element, this.domElement);
    };
    SpfxPnpRichtextWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(SpfxPnpRichtextWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    SpfxPnpRichtextWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                PropertyPaneTextField('AzureSubscriptionKey', {
                                    label: "Azure Subscription Key"
                                }),
                                PropertyPaneTextField('ServiceName', {
                                    label: "Service Name"
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return SpfxPnpRichtextWebPart;
}(BaseClientSideWebPart));
export default SpfxPnpRichtextWebPart;
//# sourceMappingURL=SpfxPnpRichtextWebPart.js.map