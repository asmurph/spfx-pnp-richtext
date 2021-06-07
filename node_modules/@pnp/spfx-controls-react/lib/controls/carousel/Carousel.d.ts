/// <reference types="react" />
import * as React from "react";
import { ICarouselProps, ICarouselState } from ".";
export declare class Carousel extends React.Component<ICarouselProps, ICarouselState> {
    constructor(props: ICarouselProps);
    /**
     * Handles component update lifecycle method.
     * @param prevProps
     */
    componentDidUpdate(prevProps: ICarouselProps): void;
    render(): React.ReactElement<ICarouselProps>;
    /**
     * Return merged styles for Button containers.
     */
    private getButtonContainerStyles();
    /**
     * Return merged styles for Buttons.
     * @param nextButton
     */
    private getButtonStyles(nextButton);
    /**
     * Merges the styles of the components.
     */
    private getMergedStyles;
    /**
     * Determines if the carousel button can be clicked.
     */
    private isCarouselButtonDisabled;
    /**
     * Handles carousel button click.
     */
    private onCarouselButtonClicked;
    /**
     * Returns next index after carousel button is clicked.
     */
    private getNextIndex;
    /**
     * Returns current element to be displayed.
     */
    private getElementToDisplay;
}
