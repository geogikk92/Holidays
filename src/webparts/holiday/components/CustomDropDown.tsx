import * as React from 'react';
import { Label, Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps, TooltipHost, DirectionalHint, TooltipDelay } from 'office-ui-fabric-react';
import { Spinner } from 'office-ui-fabric-react/lib/components/Spinner';

import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from '../components/Holiday.module.scss';
import { CustomDropDownProps } from '../components/CustomDropDownProps';
import { CustomDropDownState } from '../components/CustomDropDownState';
import * as strings from 'holidayStrings';

export class CustomDropDown extends React.Component<CustomDropDownProps, CustomDropDownState> {

    constructor(props: CustomDropDownProps, state: CustomDropDownState) {
        super(props);
        this.state = {
            loading: false,
            options: undefined,
            error: undefined
        };
    }

    public componentDidMount(): void {
        this.loadOptions();
    }

    public componentDidUpdate(prevProps: CustomDropDownProps, prevState: CustomDropDownState): void {
        if (this.props.disabled !== prevProps.disabled ||
            this.props.stateKey !== prevProps.stateKey) {
            this.loadOptions();
        }
    }

    private loadOptions(): void {
        this.setState({
            loading: true,
            error: undefined,
            options: undefined
        });

        this.props.loadOptions()
            .then((options: IDropdownOption[]): void => {
                this.setState({
                    loading: false,
                    error: undefined,
                    options: options
                    //defaultValue: options[1].key.toString()
                });
            }, (error: any): void => {
                this.setState((prevState: CustomDropDownState, props: CustomDropDownProps): CustomDropDownState => {
                    prevState.loading = false;
                    prevState.error = error;
                    return prevState;
                });
            });
    }

    public render() {

        if (!this.state.loading) {
            return (
                <div className='dropdownExample'>
                    <div className={`ms-Grid-row  ${styles.row}`}>
                        {(this.props.selectedKey === 'undefined|undefined') ?
                            <Label required={true}>{this.props.label}</Label>
                            :
                            <Label>Избран тип: <strong>{this.props.selectedKey.toString().split("|")[0]}</strong></Label>

                        }
                    </div>
                    <div className={`ms-Grid-row  ${styles.row}`}>
                        <TooltipHost
                            tooltipProps={{
                                onRenderContent: () => {
                                    return (
                                        <div>
                                            {(this.props.selectedKey !== 'undefined|undefined') ?
                                                <ul style={{ margin: 0, padding: 0 }}>
                                                    <li>Тип: {this.props.selectedKey.toString().split("|")[0]}</li>
                                                    <li>Подтип: {this.props.selectedKey.toString().split("|")[1]}</li>
                                                </ul>
                                                : "Моля изберете тип отпуск"}
                                        </div>
                                    );
                                }
                            }}
                        >
                            <Dropdown
                                selectedKey={this.props.selectedKey}
                                onChanged={this.props.onChanged}
                                placeHolder='Моля, изберете тип отпуск'
                                options={this.state.options} />
                        </TooltipHost>
                    </div>
                </div >
            );
        }
        else {
            return (
                <div className={`ms-Grid-row  ${styles.row}`}>
                    <Spinner label={'Зареждане на типове отпуск....'} />
                </div>
            );
        }
    }
}