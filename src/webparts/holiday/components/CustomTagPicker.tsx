import * as React from 'react';
import * as ReactDom from 'react-dom';

import { escape } from '@microsoft/sp-lodash-subset';
import { BaseComponent, assign, autobind } from 'office-ui-fabric-react/lib/Utilities';
import { TextField, Dropdown, Label, DefaultButton, IButtonProps, Toggle, TagPicker, MessageBar, MessageBarType } from 'office-ui-fabric-react';

import styles from '../components/Holiday.module.scss';
import { CustomTagPickerProps } from './CustomTagPickerProps';
import { CustomTagPickerState } from './CustomTagPickerState';
import { BaseEmployee } from '../models/BaseEmployee';
import { Employee } from '../models/Employee';


export class CustomTagPicker extends React.Component<CustomTagPickerProps, CustomTagPickerState> {
    constructor() {
        super();
        this.state = {
            isPickerDisabled: false,
            tags: undefined
        };
    }

    public componentDidMount(): void {
        let su = 'http://sf-spsdev07:42325/api/Employees/GetNomenclatureUser?l=1026';
        this.getEmployee(su, this.props.uName, this.props.uPosition)
            .then((e: Employee) => {
                this.setState({
                    isPickerDisabled: false,
                    tags: e.Replacements.map(item => ({ key: item.UserPosition, name: item.UserName }))
                });
            })
            .catch(err => {
                // Something went wrong. Save the error in state and re-render.
                this.setState({
                    isPickerDisabled: true,
                    tags: undefined
                });
            });
    }

    private getEmployee(requestUrl: string, userName: string, userPosition: string): Promise<Employee> {
        let employee: BaseEmployee = {
            UserName: userName,
            UserPosition: userPosition
        };

        let sUrl = 'http://sf-spsdev07:34006/sites/BG/SPFxSamples';
        return new Promise<Employee>((resolve: (item: Employee) => void, reject: (error: any) => void): void => {
            const parseFetchResponse = response => response.json().then(text => ({
                json: text,
                meta: response,
            }));
            window.fetch(requestUrl, {
                method: 'POST',
                body: JSON.stringify({
                    SiteUrl: sUrl,
                    BaseEmp: employee
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            })
                .then(parseFetchResponse)
                .then(({ json, meta }) => {
                    return json;
                })
                .then((item: Employee) => {
                    const e: Employee = item;
                    resolve(e);
                })
                , (error: any): void => { reject(error); };
        });
    }


    public render(): React.ReactElement<CustomTagPickerProps> {
        return (
            <div>
                <div className={`ms-Grid-row  ${styles.row}`}>
                    <Label required={true} disabled={false}>{this.props.children}</Label>
                </div>
                <div className={`ms-Grid-row  ${styles.row}`}>
                    <TagPicker
                        ref='tagPicker'
                        className='#002050'
                        inputProps={{ disabled: this.state.isPickerDisabled }}
                        onResolveSuggestions={this._onFilterChanged}
                        onChange={this._onChange.bind(this)}
                        getTextFromItem={(item: any) => { return item.key; }}
                        pickerSuggestionsProps={
                            {
                                suggestionsHeaderText: 'Предпочитан заместник',
                                noResultsFoundText: 'Няма намерени резултати'
                            }
                        }
                    />
                </div>
            </div>
        );
    }

    private _onChange() {
        this.setState({ isPickerDisabled: !this.state.isPickerDisabled });
    }


    @autobind
    private _onFilterChanged(filterText: string, tagList: { key: string, name: string }[]) {
        if (this.state.tags === undefined) {
            return [];
        }

        if (this.state.tags.length > 0) {
            let tgs = this.state.tags;
            return filterText ? tgs.filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0) : [];
        }
        else {
            return [];
        }
    }
}