import * as React from 'react';
import * as ReactDom from 'react-dom';

import styles from '../components/Holiday.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, Label } from 'office-ui-fabric-react';
import { CustomTextFieldProps } from './CustomTextFieldProps';

export class CustomTextField extends React.Component<CustomTextFieldProps, void> {

    public render(): React.ReactElement<CustomTextFieldProps> {
        return (
            <div>
                <div className={`ms-Grid-row  ${styles.row}`}>
                    <Label required={this.props.isRequared} disabled={false}>{this.props.label}</Label>
                </div>
                <div className={`ms-Grid-row  ${styles.row}`}>
                    <TextField
                        multiline={this.props.isMultiline}
                        rows={4}
                        resizable={false}
                        onChanged={this.props.onChanged}
                        disabled={this.props.disabled}
                        value={this.props.value}
                        onGetErrorMessage={this.props.onErrorMsg}
                    />
                </div>
            </div>
        );
    }
}