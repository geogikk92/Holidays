import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './components/Holiday.module.scss';
import { IHolidayProps } from './IHolidayProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Holiday } from './models/Holiday';
import { Employee } from './models/Employee';
import { Label } from 'office-ui-fabric-react';
import * as pnp from 'sp-pnp-js';
import * as strings from 'holidayStrings';
import { CustomButton } from './components/CustomButton';
import { CustomTextField } from './components/CustomTextField';
import { CustomMessage } from './components/CustomMessage';
import { autobind } from '@uifabric/utilities';
import { CustomDataSavedSuccess } from './components/CustomDataSavedSuccess';

export interface FormApproveState {
    dtFrom?: string;
    dtTo?: string;
    description?: string;
    typeRequest?: string;
    type?: string;
    fullName?: string;
    days?: number;
    dataWasLoaded?: boolean;
    urlToHolidaySite?: string;
}

export interface FormApproveProps {
    URLAddressToHolidaySite: string;
}



export default class ApproveForm extends React.Component<FormApproveProps, FormApproveState> {
    constructor(props) {
        super(props);
        this.state = {
            dataWasLoaded: false,
            urlToHolidaySite: this.props.URLAddressToHolidaySite
        }
    }

    public componentWillMount(): void {
        let ID: number = Number(this.getParameterByName("myID", document.location.href));
        pnp.sp.web.lists.getByTitle("Отпуски").items.getById(ID).get().then(
            (holidayItem: any) => {
                pnp.sp.web.currentUser.get().then(currentUser => {
                    if (holidayItem.LirexHolidaySupervisorId === currentUser.Id) {
                        this.setState({ dataWasLoaded: true });
                    }
                    else {
                        pnp.sp.web.siteUsers.getById(currentUser.Id).groups.getByName("Човешки ресурси").get().then(
                            (item: any) => {
                                this.setState({ dataWasLoaded: true });
                            },
                            (error: any) => {
                                this.setState({ dataWasLoaded: false });
                                ReactDom.render(<CustomDataSavedSuccess messageType={5} messageText="Нямате права за да извършите тази операция!" redirectTo={this.state.urlToHolidaySite} />, document.getElementById('formContent'));
                            }
                        );
                    }

                    let dateFrom: Date = new Date(holidayItem.LirexHolidayDateFrom);
                    let dateTo: Date = new Date(holidayItem.LirexHolidayDateTo);
                    this.setState((prevState: FormApproveState): FormApproveState => {
                        prevState.dtFrom = dateFrom.getDate() + '.' + (dateFrom.getMonth() + 1) + '.' + dateFrom.getFullYear();
                        prevState.dtTo = dateTo.getDate() + '.' + (dateTo.getMonth() + 1) + '.' + dateTo.getFullYear();
                        prevState.days = holidayItem.LirexHolidayDays;
                        prevState.typeRequest = holidayItem.LirexHolidayTypeRequest;
                        prevState.fullName = holidayItem.LirexHolidayEmpFullName;
                        prevState.type = holidayItem.LirexHolidayType;
                        return prevState;
                    });
                });
            },
            (error: any) => {
                ReactDom.render(<CustomDataSavedSuccess messageType={5} messageText={error.message} redirectTo={this.state.urlToHolidaySite} />, document.getElementById('formContent'));
            }
        );
    }

    public render(): React.ReactElement<{}> {
        return (
            <div className={styles.helloWorld} >
                <div id="formContent" className={styles.container}>
                    {(this.state.dataWasLoaded) ?
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">{this.state.fullName}</div>
                                <strong> <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.typeRequest}</div></strong>
                            </div>
                            <hr />
                            {(this.state.typeRequest == strings.holiday) ?
                                <div className="ms-Grid-row">
                                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Тип:</div>
                                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.type}</div>
                                </div>
                                : null
                            }
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Начална дата: </div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.dtFrom}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Крайна дата:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.dtTo}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Брой работни дни:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.days}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Описание към заявката:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">
                                    <CustomTextField
                                        isMultiline={true}
                                        isRequared={false}
                                        value={this.state.description}
                                        onChanged={this._onChangedHolidayDescription} />
                                </div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm4">
                                    <CustomButton
                                        value={"Одобрявам"}
                                        type="reset"
                                        onClick={this._handleApprove}
                                    />
                                </div>
                                <div className="ms-Grid-col ms-sm4">
                                    <CustomButton
                                        value={'Неодобрявам'}
                                        type="reset"
                                        onClick={this._handleReject}
                                    />
                                </div>
                                <div className="ms-Grid-col ms-sm4">
                                    <CustomButton
                                        value={"Отказвам се"}
                                        type="reset"
                                        onClick={this._handleRedirect}
                                    />
                                </div>
                            </div>
                        </div>
                        : null}
                </div>
            </div>
        );
    }

    private getParameterByName(name, url): string {
        if (!url) url = document.location.href;
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }

    @autobind
    private _handleRedirect(): void {
        window.location.href = this.state.urlToHolidaySite;
    }

    @autobind
    private _handleApprove(): void {
        window.location.href = this.state.urlToHolidaySite;
    }

    @autobind
    private _handleReject(): void {
        window.location.href = this.state.urlToHolidaySite;
    }

    @autobind
    public _onChangedHolidayDescription(newValue: string): void {
        this.setState((prevState: FormApproveState): FormApproveState => {
            prevState.description = newValue;
            return prevState;
        });
    }
}
