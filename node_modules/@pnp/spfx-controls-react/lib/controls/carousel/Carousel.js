"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var Button_1 = require("office-ui-fabric-react/lib/Button");
var icons_1 = require("@uifabric/icons");
icons_1.initializeIcons();
var React = require("react");
var Carousel_module_scss_1 = require("./Carousel.module.scss");
var _1 = require(".");
var lib_1 = require("@uifabric/utilities/lib");
var ICarouselState_1 = require("./ICarouselState");
var Spinner_1 = require("office-ui-fabric-react/lib/Spinner");
var common_1 = require("@pnp/common");
var telemetry = require("../../common/telemetry");
var Carousel = (function (_super) {
    __extends(Carousel, _super);
    function Carousel(props) {
        var _this = _super.call(this, props) || this;
        /**
         * Merges the styles of the components.
         */
        _this.getMergedStyles = function (componentStyles, userStyles) {
            var mergedStyles = userStyles ? lib_1.css(componentStyles, userStyles) : lib_1.css(componentStyles);
            return mergedStyles;
        };
        /**
         * Determines if the carousel button can be clicked.
         */
        _this.isCarouselButtonDisabled = function (nextButton) {
            // false by default
            var isInfinite = _this.props.isInfinite != undefined ? _this.props.isInfinite : false;
            var currentIndex = _this.state.currentIndex;
            var result = false;
            // Use validation from parent control or calcualte it based on the current index
            if (nextButton) {
                result = _this.props.canMoveNext != undefined ?
                    !_this.props.canMoveNext :
                    (currentIndex === _this.props.element.length - 1) && !isInfinite;
            }
            else {
                result = _this.props.canMovePrev != undefined ?
                    !_this.props.canMovePrev :
                    (0 === currentIndex) && !isInfinite;
            }
            return result;
        };
        /**
         * Handles carousel button click.
         */
        _this.onCarouselButtonClicked = function (nextButtonClicked) {
            var currentIndex = _this.state.currentIndex;
            var nextIndex = _this.state.currentIndex;
            var processingState = ICarouselState_1.ProcessingState.processing;
            // Trigger parent control to update provided element
            if (_this.props.triggerPageEvent) {
                // Index validation needs to be done by the parent control specyfing canMove Next|Prev
                nextIndex = nextButtonClicked ? (currentIndex + 1) : (currentIndex - 1);
                // Trigger parent to provide new data
                _this.props.triggerPageEvent(nextIndex);
                processingState = ICarouselState_1.ProcessingState.processing;
            }
            else {
                nextIndex = _this.getNextIndex(nextButtonClicked);
                var canMoveNext = _this.props.canMoveNext != undefined ? _this.props.canMoveNext : true;
                var canMovePrev = _this.props.canMovePrev != undefined ? _this.props.canMovePrev : true;
                if (canMoveNext && nextButtonClicked && _this.props.onMoveNextClicked) {
                    _this.props.onMoveNextClicked(nextIndex);
                }
                else if (canMovePrev && !nextButtonClicked && _this.props.onMovePrevClicked) {
                    _this.props.onMovePrevClicked(nextIndex);
                }
                processingState = ICarouselState_1.ProcessingState.idle;
            }
            _this.setState({
                currentIndex: nextIndex,
                processingState: processingState
            });
        };
        /**
         * Returns next index after carousel button is clicked.
         */
        _this.getNextIndex = function (nextButtonClicked) {
            var currentIndex = _this.state.currentIndex;
            var nextIndex = currentIndex;
            var isInfinite = _this.props.isInfinite !== undefined ? _this.props.isInfinite : false;
            var length = _this.props.element.length;
            // Next Button clicked
            if (nextButtonClicked) {
                // If there is next element
                if (currentIndex < length - 1) {
                    nextIndex = currentIndex + 1;
                }
                else if (isInfinite) {
                    nextIndex = 0;
                }
            }
            else {
                if (currentIndex - 1 >= 0) {
                    // If there is previous element
                    nextIndex = currentIndex - 1;
                }
                else if (isInfinite) {
                    // If there is no previous element but isInfitineLoop -> reset index to the last element
                    nextIndex = length - 1;
                }
            }
            return nextIndex;
        };
        /**
         * Returns current element to be displayed.
         */
        _this.getElementToDisplay = function () {
            var element = _this.props.element;
            var currentIndex = _this.state.currentIndex;
            var result = null;
            // If no element has been provided.
            if (!element) {
                result = null;
            }
            else if (common_1.isArray(element) && currentIndex >= 0 && element.length > currentIndex) {
                result = element[currentIndex];
            }
            else {
                result = element;
            }
            return result;
        };
        var currentIndex = props.startIndex ? props.startIndex : 0;
        telemetry.track('ReactCarousel', {});
        _this.state = {
            currentIndex: currentIndex,
            processingState: ICarouselState_1.ProcessingState.idle
        };
        return _this;
    }
    /**
     * Handles component update lifecycle method.
     * @param prevProps
     */
    Carousel.prototype.componentDidUpdate = function (prevProps) {
        var currProps = this.props;
        var prevPropsElementKey = prevProps.triggerPageEvent && prevProps.element ? prevProps.element.key : null;
        var nextPropsElementKey = currProps.triggerPageEvent && currProps.element ? currProps.element.key : null;
        // Checking if component is in processing state and the key of the current element has been changed
        if (this.state.processingState === ICarouselState_1.ProcessingState.processing && nextPropsElementKey != null && prevPropsElementKey != nextPropsElementKey) {
            this.setState({
                processingState: ICarouselState_1.ProcessingState.idle
            });
        }
    };
    Carousel.prototype.render = function () {
        var _this = this;
        var _a = this.state, currentIndex = _a.currentIndex, processingState = _a.processingState;
        var _b = this.props, containerStyles = _b.containerStyles, contentContainerStyles = _b.contentContainerStyles, containerButtonsStyles = _b.containerButtonsStyles, prevButtonStyles = _b.prevButtonStyles, nextButtonStyles = _b.nextButtonStyles, loadingComponentContainerStyles = _b.loadingComponentContainerStyles;
        var prevButtonIconName = this.props.prevButtonIconName ? this.props.prevButtonIconName : "ChevronLeft";
        var nextButtonIconName = this.props.nextButtonIconName ? this.props.nextButtonIconName : "ChevronRight";
        var processing = processingState === ICarouselState_1.ProcessingState.processing;
        var prevButtonDisabled = processing || this.isCarouselButtonDisabled(false);
        var nextButtonDisabled = processing || this.isCarouselButtonDisabled(true);
        var loadingComponent = this.props.loadingComponent ? this.props.loadingComponent : React.createElement(Spinner_1.Spinner, null);
        var element = this.getElementToDisplay();
        return (React.createElement("div", { className: this.getMergedStyles(Carousel_module_scss_1.default.container, containerStyles) },
            React.createElement("div", { className: this.getMergedStyles(this.getButtonContainerStyles(), containerButtonsStyles), onClick: function () { if (!prevButtonDisabled) {
                    _this.onCarouselButtonClicked(false);
                } } },
                React.createElement(Button_1.IconButton, { className: this.getMergedStyles(this.getButtonStyles(false), prevButtonStyles), iconProps: { iconName: prevButtonIconName }, disabled: prevButtonDisabled, onClick: function () { _this.onCarouselButtonClicked(false); } })),
            React.createElement("div", { className: this.getMergedStyles(Carousel_module_scss_1.default.contentContainer, contentContainerStyles) },
                processing &&
                    React.createElement("div", { className: this.getMergedStyles(Carousel_module_scss_1.default.loadingComponent, loadingComponentContainerStyles) }, loadingComponent),
                !processing && element &&
                    element),
            React.createElement("div", { className: this.getMergedStyles(this.getButtonContainerStyles(), containerButtonsStyles), onClick: function () { if (!nextButtonDisabled) {
                    _this.onCarouselButtonClicked(true);
                } } },
                React.createElement(Button_1.IconButton, { className: this.getMergedStyles(this.getButtonStyles(true), nextButtonStyles), iconProps: { iconName: nextButtonIconName }, disabled: nextButtonDisabled, onClick: function () { _this.onCarouselButtonClicked(true); } }))));
    };
    /**
     * Return merged styles for Button containers.
     */
    Carousel.prototype.getButtonContainerStyles = function () {
        var buttonsDisplayMode = this.props.buttonsDisplay ? this.props.buttonsDisplay : _1.CarouselButtonsDisplay.block;
        var buttonDisplayModeCss = "";
        switch (buttonsDisplayMode) {
            case _1.CarouselButtonsDisplay.block:
                buttonDisplayModeCss = Carousel_module_scss_1.default.blockButtonsContainer;
                break;
            case _1.CarouselButtonsDisplay.buttonsOnly:
                buttonDisplayModeCss = Carousel_module_scss_1.default.buttonsOnlyContainer;
                break;
            case _1.CarouselButtonsDisplay.hidden:
                buttonDisplayModeCss = Carousel_module_scss_1.default.hiddenButtonsContainer;
                break;
            default:
                return "";
        }
        var buttonsLocation = this.props.buttonsLocation ? this.props.buttonsLocation : _1.CarouselButtonsLocation.top;
        var buttonsLocationCss = "";
        switch (buttonsLocation) {
            case _1.CarouselButtonsLocation.top:
                buttonsLocationCss = Carousel_module_scss_1.default.blockButtonsContainer;
                break;
            case _1.CarouselButtonsLocation.center:
                buttonsLocationCss = Carousel_module_scss_1.default.centralButtonsContainer;
                break;
            case _1.CarouselButtonsLocation.bottom:
                buttonsLocationCss = Carousel_module_scss_1.default.bottomButtonsContainer;
                break;
            default:
                return "";
        }
        var result = lib_1.css(buttonDisplayModeCss, buttonsLocationCss);
        return result;
    };
    /**
     * Return merged styles for Buttons.
     * @param nextButton
     */
    Carousel.prototype.getButtonStyles = function (nextButton) {
        var buttonsDisplayMode = this.props.buttonsDisplay ? this.props.buttonsDisplay : _1.CarouselButtonsDisplay.block;
        var result = "";
        if (buttonsDisplayMode === _1.CarouselButtonsDisplay.buttonsOnly) {
            result = nextButton ? Carousel_module_scss_1.default.buttonsOnlyNextButton : Carousel_module_scss_1.default.buttonsOnlyPrevButton;
        }
        return lib_1.css(result);
    };
    return Carousel;
}(React.Component));
exports.Carousel = Carousel;

//# sourceMappingURL=Carousel.js.map
