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
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './SpfxPnpRichtext.module.scss';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Dropdown } from 'office-ui-fabric-react';
import $ from "jquery";
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { MessageBarType } from 'office-ui-fabric-react';
var stackStyles = { root: { width: 650 } };
var stackTokens = { childrenGap: 50 };
var columnProps = {
    tokens: { childrenGap: 15 },
    styles: { root: { width: 300 } },
};
var smallcolumnProps = {
    tokens: { childrenGap: 15 },
    styles: { root: { width: 180 } },
};
var SpfxPnpRichtext = /** @class */ (function (_super) {
    __extends(SpfxPnpRichtext, _super);
    function SpfxPnpRichtext(props, state) {
        var _this = _super.call(this, props) || this;
        _this.state = ({ toLanguage: '', Title: '', Description: '', richtext: '', langarr: [] });
        _this._getSupportedLangualge();
        _this._getData();
        return _this;
    }
    SpfxPnpRichtext.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.spfxPnpRichtext },
            React.createElement(Stack, { horizontal: true, tokens: stackTokens, styles: stackStyles },
                React.createElement(Stack, __assign({}, columnProps), this.state.Description),
                React.createElement(Stack, __assign({}, smallcolumnProps),
                    React.createElement(Dropdown, { placeholder: "Select a language", label: "Select Language", options: this.state.langarr, onChanged: function (value) { _this.setState({ toLanguage: value.key.toString() }); _this._translate(); } })),
                React.createElement(Stack, __assign({}, columnProps),
                    React.createElement("label", null, this.state.richtext)))));
    };
    SpfxPnpRichtext.prototype._getData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var richTextItem, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists.getByTitle('ListTest').items.getById(1)
                                .select("ID", "Title", "Description")
                                .get()];
                    case 1:
                        richTextItem = _a.sent();
                        this.setState({
                            Title: richTextItem.Title,
                            Description: richTextItem.Description,
                        });
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        this.setState({
                            MessageText: "Exception reading item",
                            MessageType: MessageBarType.error
                        });
                        return [2 /*return*/, Promise.reject(error_1)];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    SpfxPnpRichtext.prototype._getSupportedLangualge = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                $.get({
                    url: 'https://api.cognitive.microsofttranslator.com/languages?api-version=3.0&scope=translation'
                })
                    .done(function (languages) {
                    var droparr = [];
                    var langobjs = languages.translation;
                    for (var key in langobjs) {
                        if (langobjs.hasOwnProperty(key)) {
                            droparr.push({ key: key, text: langobjs[key].name });
                        }
                    }
                    _this.setState({ langarr: droparr });
                });
                return [2 /*return*/];
            });
        });
    };
    SpfxPnpRichtext.prototype._translate = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                $.post({
                    url: 'https://' + this.props.ServiceName + '.cognitiveservices.azure.com/sts/v1.0/issueToken',
                    headers: {
                        'Ocp-Apim-Subscription-Key': this.props.AzureSubscriptionKey,
                        'Authorization': this.props.ServiceName + '.cognitiveservices.azure.com'
                    }
                })
                    .done(function (tocken) {
                    $.post({
                        url: 'https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to=' + _this.state.toLanguage,
                        headers: {
                            'Ocp-Apim-Subscription-Key': _this.props.AzureSubscriptionKey,
                            'Authorization': 'Bearer ' + tocken,
                            'Content-Type': 'application/json'
                        },
                        data: JSON.stringify([{ "Text": _this.state.Description }])
                    })
                        .done(function (result) {
                        console.log(result);
                        _this.setState({ Description: result[0].translations[0].text });
                    });
                });
                return [2 /*return*/];
            });
        });
    };
    return SpfxPnpRichtext;
}(React.Component));
export default SpfxPnpRichtext;
//# sourceMappingURL=SpfxPnpRichtext.js.map