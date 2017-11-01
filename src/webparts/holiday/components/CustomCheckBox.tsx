import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'holidayStrings';
import styles from '../components/Holiday.module.scss';
import { Toggle } from 'office-ui-fabric-react';
import { CustomCheckBoxProps } from '../components/CustomCheckBoxProps';
import { CustomCheckBoxState } from '../components/CustomCheckBoxState';

export class CustomCheckBox extends React.Component<CustomCheckBoxProps, CustomCheckBoxState> {

    public render(): React.ReactElement<CustomCheckBoxProps> {
        return (
            <Toggle
                //label={strings.ChoiceHolidayOrHospital}
                onText={strings.ChoiceHoliday}
                offText={strings.ChoiceHospital}
                checked={this.props.checked}
                onChanged={this.props.onChanged}
                disabled={this.props.disabled}
            />
        );
    }
}