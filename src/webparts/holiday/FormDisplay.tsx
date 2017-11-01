import * as React from 'react';
import styles from './components/Holiday.module.scss';
import { IHolidayProps } from './IHolidayProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Holiday } from './models/Holiday';
import { Employee } from './models/Employee';
import { Label } from 'office-ui-fabric-react';
import * as pnp from 'sp-pnp-js';
import * as strings from 'holidayStrings';
import { CustomButton } from './components/CustomButton';
import { autobind } from '@uifabric/utilities';

export interface FormDisplayState {
    holiday?: Holiday;
    employee?: Employee
    dataWasLoaded?: boolean
}

export default class DisplayForm extends React.Component<{}, FormDisplayState> {
    constructor(props) {
        super(props);
        this.state = {
            employee: new Employee(),
            holiday: new Holiday(),
            dataWasLoaded: false
        }
    }

    public componentWillMount(): void {
        let ID: number = Number(this.getParameterByName("ID", document.location.href));

        pnp.sp.web.lists.getByTitle("Отпуски").items.getById(ID).get().then((item: any) => {
            console.log(item);
            this.setState((prevState: FormDisplayState): FormDisplayState => {
                prevState.holiday.Id = ID;
                prevState.holiday.Title = item.Title;
                prevState.holiday.Address = item.LirexHolidayAddress;
                prevState.holiday.DateFrom = item.LirexHolidayDateFrom;
                prevState.holiday.DateTo = item.LirexHolidayDateTo;
                prevState.holiday.Days = item.LirexHolidayDays;
                prevState.holiday.Description = item.LirexHolidayDescription;
                prevState.holiday.Mobile = item.LirexHolidayMobile;
                prevState.holiday.SubType = item.LirexHolidaySubType;
                prevState.holiday.Type = item.LirexHolidayType;
                prevState.holiday.Status = item.LirexHolidayStatus;
                prevState.holiday.TypeRequest = item.LirexHolidayTypeRequest;
                prevState.employee.FullName = item.LirexHolidayEmpFullName;
                prevState.employee.UserPosition = item.LirexHolidayEmployeePosition;

                prevState.dataWasLoaded = true;
                return prevState;
            });

            // nomenclatureWeb.lists.getByTitle('Служители').items.filter(`LirexEmplUser eq  + '${holidayItem.LirexHolidayReplacementId}'`).top(1).get().then(replacement => {
            // });
        });
    }

    public render(): React.ReactElement<{}> {
        return (
            (this.state.dataWasLoaded) ?
                <div className={styles.helloWorld} >
                    <div id="formContent" className={styles.container}>
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Тип на заявката:</div>
                                <strong><div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.TypeRequest}</div></strong>
                            </div>
                            <hr />
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Статус:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.Status}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Три имена на служителя:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.employee.FullName}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Позиция на служителя:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.employee.UserPosition}</div>
                            </div>
                            {(this.state.holiday.TypeRequest !== strings.hospital) ?
                                <div className="ms-Grid-row">
                                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Адрес:</div>
                                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.Address}</div>
                                </div>
                                : null}
                            {(this.state.holiday.TypeRequest !== strings.hospital) ?
                                <div className="ms-Grid-row">
                                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Мобилен номер:</div>
                                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.Mobile}</div>
                                </div>
                                : null}
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Описание към заявката:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.Description}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Избрани дни за {this.state.holiday.TypeRequest.toLowerCase()}:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.Days}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Дата от:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.DateFrom.toLocaleString()}</div>
                            </div>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Дата до:</div>
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.DateTo.toLocaleString()}</div>
                            </div>
                            {(this.state.holiday.TypeRequest !== strings.hospital) ?
                                <div className="ms-Grid-row">
                                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Тип на отпуската:</div>
                                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.Type}</div>
                                </div>
                                : null}
                            {(this.state.holiday.TypeRequest !== strings.hospital) ?
                                <div className="ms-Grid-row">
                                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">Подтип на отпуската:</div>
                                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg6">{this.state.holiday.SubType}</div>
                                </div>
                                : null}
                            <div className="ms-Grid-col ms-sm8 ms-smPush2">
                                <CustomButton
                                    value={strings.btnClear}
                                    type="reset"
                                    onClick={this._handleRedirect}
                                />
                            </div>
                        </div>
                    </div>
                </div>
                : null
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
        window.location.href = document.referrer;
    }
}
