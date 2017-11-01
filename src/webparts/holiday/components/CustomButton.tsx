import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from '../components/Holiday.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, IButtonProps, Spinner, SpinnerType, SpinnerSize } from 'office-ui-fabric-react';
import { CustomButtonProps } from './CustomButtonProps';
import { CustomButtonState } from './CustomButtonState';

export class CustomButton extends React.Component<CustomButtonProps, CustomButtonState> {

    public render(): React.ReactElement<CustomButtonProps> {
        return (
            <DefaultButton
                onClick={this.props.onClick}
                disabled={this.props.disabled}
            >
                {this.props.value}
                {(this.props.loadSpinner) ? < Spinner type={SpinnerType.normal} size={SpinnerSize.xSmall} /> : null}
            </DefaultButton>
        );
    }
}