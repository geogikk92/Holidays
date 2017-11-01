import * as React from 'react';
import * as ReactDom from 'react-dom';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './components/Holiday.module.scss';
import * as strings from 'holidayStrings';
import {
  TextField, Image, autobind, Icon, Label, IChoiceGroupOption, IDropdownOption, IPersonaProps, PersonaPresence,
  Spinner, SpinnerSize, SpinnerType, IBasePickerSuggestionsProps, IImageProps, ImageFit, CompoundButton
}
  from 'office-ui-fabric-react';
import { FormEditProps } from './FormEditProps';
import { FormEditState } from './FormEditState';
import { Employee } from './models/Employee';
import { Holiday } from './models/Holiday';
import { BaseEmployee } from './models/BaseEmployee';
import { CustomDateTimePicker } from './components/CustomDateTimePicker';
import { CustomTable } from './components/CustomTable';
import { CustomTextField } from './components/CustomTextField';
import { CustomPeoplePicker } from './components/CustomPeoplePicker';
import { CustomButton } from './components/CustomButton';
import { CustomCheckBox } from './components/CustomCheckBox';
import { CustomMessage } from './components/CustomMessage';
import { CustomDropDown } from './components/CustomDropDown';
import { CustomPeoplePickerCommonUtility } from './components/CustomPeoplePickerCommonUtility';
import { NomenclatureTypes } from './models/NomenclatureTypes';

import { Web } from "sp-pnp-js";
import pnp from "sp-pnp-js";
import { CustomDataSavedSuccess } from './components/CustomDataSavedSuccess';


export default class FormEdit extends React.Component<FormEditProps, FormEditState> {
  constructor(props) {
    super(props);
    this.state = {
      urlToWebAPI: this.props.URLAddressToWebAPI,
      urlToHolidaySite: this.props.URLAddressToHolidaySite,
      ListName: this.props.ListTitle,
      isHoliday: true,
      isEmplDataLoaded: false,
      employee: new Employee(),
      holiday: new Holiday()
    };
  }

  public componentWillMount(): void {
    let ID: number = Number(this.getParameterByName("ID", document.location.href));
    // get a specific item by id
    pnp.sp.web.lists.getByTitle(this.state.ListName).items.getById(ID).get().then((holidayItem: any) => {
      this.setState((prevState: FormEditState): FormEditState => {
        prevState.holiday.Id = ID;
        prevState.holiday.Title = holidayItem.Title;
        prevState.holiday.Address = holidayItem.LirexHolidayAddress;
        prevState.holiday.DateFrom = holidayItem.LirexHolidayDateFrom;
        prevState.holiday.DateTo = holidayItem.LirexHolidayDateTo;
        prevState.holiday.Days = holidayItem.LirexHolidayDays;
        prevState.holiday.Description = holidayItem.LirexHolidayDescription;
        prevState.holiday.Mobile = holidayItem.LirexHolidayMobile;
        prevState.holiday.SubType = holidayItem.LirexHolidaySubType;
        prevState.holiday.Type = holidayItem.LirexHolidayType;
        prevState.holiday.Status = holidayItem.LirexHolidayStatus;
        prevState.holiday.TypeRequest = holidayItem.LirexHolidayTypeRequest;
        prevState.employee.FullName = holidayItem.LirexHolidayEmpFullName;
        prevState.employee.UserName = holidayItem.LirexHolidayEmployeeUsername;
        prevState.employee.UserPosition = holidayItem.LirexHolidayEmployeePosition;
        return prevState;
      });

      if (holidayItem.LirexHolidayTypeRequest == strings.hospital) {
        this.setState({
          isHoliday: false,
          isEmplDataLoaded: true
        });
        return;
      }

      let nomenclatureWeb = new Web(this.props.URLAddresToNomenclatures);
      nomenclatureWeb.lists.getByTitle('Служители').items.filter(`LirexEmpUserName eq  + '${holidayItem.LirexHolidayEmployeeUsername}'`).top(1).get().then(holidayEmplData => {
        this.setState((prevState: FormEditState): FormEditState => {
          prevState.employee.Id = holidayEmplData[0].ID;
          prevState.employee.UserName = holidayItem.LirexHolidayEmployeeUsername;
          prevState.employee.FullName = holidayEmplData[0].Title;
          prevState.employee.UserPosition = holidayEmplData[0].LirexEmplPosition;
          prevState.employee.Department = holidayEmplData[0].LirexEmplDirection;
          prevState.employee.Days = holidayEmplData[0].LirexEmplDays;
          prevState.employee.DaysLeftNextYear = holidayEmplData[0].NextYearDays;
          return prevState;
        });
      });

      nomenclatureWeb.lists.getByTitle('Служители').items.filter(`LirexEmplUser eq  + '${holidayItem.LirexHolidayReplacementId}'`).top(1).get().then(replacement => {

        let repl: BaseEmployee = {
          UserName: replacement[0].LirexEmpUserName,
          UserPosition: replacement[0].LirexEmplPosition
        };

        this.setState((prevState: FormEditState): FormEditState => {
          this.state.holiday.Replacement = repl;
          prevState.isEmplDataLoaded = true;
          prevState.Replacement = [{
            primaryText: replacement[0].Title,
            secondaryText: replacement[0].LirexEmplPosition,
            tertiaryText: replacement[0].Title,
            imageInitials: CustomPeoplePickerCommonUtility.getInitials(replacement[0].Title),
            presence: PersonaPresence.none
          }];
          return prevState;
        });
      });
    });
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

  public render(): React.ReactElement<FormEditProps> {
    if (this.state.isEmplDataLoaded) {
      return (
        <div className={styles.helloWorld} >
          <div className={styles.container}>
            <div className="ms-Grid">
              <div className={`ms-Grid-row ms-borderColor-themePrimary ms-fontColor-white ${styles.row}`}>
                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg12">
                  <Label className="ms-fontSize-m ms-fontColor-blueDark">{this.state.holiday.Title} - {this.state.holiday.Status}</Label>
                </div>
              </div>
              <CustomCheckBox
                checked={this.state.isHoliday}
                disabled={true}
              />
              {(this.state.isHoliday) ?
                <CustomTable
                  employeeData={this.state.employee}
                />
                : null}
              <form onSubmit={this._handleSubmit}>
                <fieldset id="formContent">
                  <CustomTextField
                    label={strings.lblEmployee}
                    isMultiline={false}
                    isRequared={false}
                    disabled={true}
                    value={this.state.employee.FullName}
                  />
                  {(this.state.isHoliday) ?
                    <CustomDropDown
                      disabled={false}
                      label={strings.lblHolidayTypes}
                      stateKey={"1"}
                      onChanged={this._onChangedHolidayType}
                      loadOptions={(): Promise<IDropdownOption[]> => {
                        return this.getHolidayTypes(this.props.URLAddressToWebAPI + strings.ActionTypeSubtypeItems,
                          this.state.urlToHolidaySite + '/Nomenclatures',
                          'HolidayType'
                        );
                      }}
                      selectedKey={this.state.holiday.Type + "|" + this.state.holiday.SubType}
                    /> : null}
                  {(this.state.isHoliday) ?
                    <CustomTextField
                      label={strings.lblAddress}
                      isMultiline={false}
                      isRequared={true}
                      value={this.state.holiday.Address}
                      onChanged={this._onChangedHolidayAddress}
                    />
                    : null}
                  {(this.state.isHoliday) ?
                    <CustomTextField
                      label={strings.lblMobile}
                      isMultiline={false}
                      isRequared={true}
                      value={this.state.holiday.Mobile}
                      onChanged={this._onChangedMobile}
                    />
                    : null}
                  <CustomTextField
                    label={strings.lblDescription}
                    isMultiline={true}
                    isRequared={false}
                    value={this.state.holiday.Description}
                    onChanged={this._onChangedHolidayDescription}
                  />
                  <CustomDateTimePicker
                    label={strings.lblDateFrom}
                    initialDateTime={this.state.holiday.DateFrom}
                    onChanged={this._onChangedHolidayDateFrom}
                  />
                  <CustomDateTimePicker
                    label={strings.lblDateTo}
                    initialDateTime={this.state.holiday.DateTo}
                    onChanged={this._onChangedHolidayDateTo}
                  />
                  <Label>{strings.lblTotalDays} {this.state.holiday.TypeRequest}: <strong>{this.state.holiday.Days}</strong></Label>
                  <div id='datesErrorMsg' />
                  {(this.state.isHoliday) ?
                    <CustomPeoplePicker
                      uName={this.state.employee.UserName}
                      uPosition={this.state.employee.UserPosition}
                      label={strings.lblReplacement}
                      onChanged={this._onChangedHolidayReplacement}
                      loadReplacements={true}
                      suggestionProps={this.getSuggestionProps(strings.suggestReplacement)}
                      defaultItem={this.state.Replacement}
                    />
                    : null}
                  <div id="replacementErrorMsg"> </div>
                  <div className={`ms-Grid-row  ${styles.row}`}>
                    <hr />
                  </div>
                  <br />
                  <div className={`ms-Grid-row  ${styles.row}`}>
                    <div className="ms-Grid-col ms-sm4 ">
                      <CustomButton
                        value={strings.btnSubmit}
                        type="Submit"
                        onClick={this._handleSubmit}
                      />
                    </div>
                    <div className="ms-Grid-col ms-sm8 ms-smPush2">
                      <CustomButton
                        value={strings.btnClear}
                        type="reset"
                        onClick={this._handleRedirect}
                      />
                    </div>
                  </div>
                  <div className={`ms-Grid-row  ${styles.row}`}>
                    <div id='formErrorMsg' />
                  </div>
                </fieldset>
              </form>
            </div>
          </div>
        </div >
      );
    }
    else {
      return (
        <div>
          <Spinner type={SpinnerType.large} size={SpinnerSize.large} label={'Зареждане на Вашите данни.....'} />
        </div>
      );
    }
  }

  @autobind
  private _handleSubmit(event): void {
    event.preventDefault();

    pnp.sp.web.currentUser.get().then(currentUser => {
      let nomenclatureWeb = new Web(this.props.URLAddresToNomenclatures);
      nomenclatureWeb.lists.getByTitle('Служители').items.filter(`LirexEmplUserId eq  + '${currentUser.Id}'`).top(1).get().then(currentUserData => {

        let requestor: BaseEmployee = {
          UserName: currentUser.LoginName.substring(currentUser.LoginName.indexOf('\\') + 1),
          UserPosition: currentUserData[0].LirexEmplPosition
        };

        let totalDaysForUser: number = Number(this.state.employee.Days) + Number(this.state.employee.DaysLeftNextYear);

        let holiday: Holiday = {
          Id: this.state.holiday.Id,
          Title: this.state.holiday.Title, //todo: implement get number from webapi by current user. (base user.)
          Address: (this.state.isHoliday) ? this.state.holiday.Address : null,
          Days: this.state.holiday.Days,
          DateFrom: this.state.holiday.DateFrom,
          DateTo: this.state.holiday.DateTo,
          Replacement: (this.state.isHoliday) ? this.state.holiday.Replacement : null,
          Description: this.state.holiday.Description,
          TypeRequest: (this.state.isHoliday) ? strings.holiday : strings.hospital,
          Type: (this.state.isHoliday) ? this.state.holiday.Type : null,
          SubType: (this.state.isHoliday) ? this.state.holiday.SubType : null,
          Mobile: (this.state.isHoliday) ? this.state.holiday.Mobile : null,
          TotalAllowedDays: totalDaysForUser,
          Employee: this.state.employee,
          Requestor: requestor,
          SiteUrl: this.state.urlToHolidaySite
        };

        const parseFetchResponse = response => response.json().then(text => ({
          json: text,
          meta: response,
        }));

        const cnst = JSON.stringify(holiday);

        window.fetch(this.props.URLAddressToWebAPI + strings.ActionUpdate, {
          method: 'POST',
          body: JSON.stringify(holiday),
          headers: {
            'Content-Type': 'application/json'
          }
        })
          .then(parseFetchResponse)
          .then(({ json, meta }) => {
            const msg: string = json;
            if (meta.ok) {
              this.setState((prevState: FormEditState): FormEditState => {
                prevState.employee.Days = prevState.employee.Days - prevState.holiday.Days;
                return prevState;
              });

              ReactDom.render(<CustomDataSavedSuccess messageType={4} messageText="Данните са редактирани успешно!" redirectTo={this.props.URLAddressToHolidaySite} />, document.getElementById('formContent'));
            }
            else {
              ReactDom.render(<CustomMessage messageType={5} messageText={msg} messageVisible={true} />, document.getElementById('formErrorMsg'));
            }
          }, (error: any): void => {
            ReactDom.render(<CustomMessage messageType={5} messageText={error.message} messageVisible={true} />, document.getElementById('formErrorMsg'));
          });
      });
    });
  }

  @autobind
  private _handleRedirect(): void {
    window.location.href = this.props.URLAddressToHolidaySite;
  }

  @autobind
  public _onChangedHolidayAddress(newValue: string): void {
    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.Address = newValue;
      return prevState;
    });
  }

  @autobind
  public _onChangedHolidayDescription(newValue: string): void {
    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.Description = newValue;
      return prevState;
    });
  }

  @autobind
  private _onChangedHolidayDateFrom(dateFrom: Date): void {
    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.DateFrom = dateFrom;
      return prevState;
    });

    if (this.state.holiday.DateTo === undefined)
      return;

    let hDateFrom: Date = dateFrom;
    let hDateTo: Date = new Date(this.state.holiday.DateTo);

    let requestUrl = this.state.urlToWebAPI + strings.ActionCalculateWorkingDays;
    const parseFetchResponse = response => response.json().then(text => ({
      json: text,
      meta: response,
    }));
    window.fetch(requestUrl, {
      method: 'POST',
      body: JSON.stringify({
        dFirst: hDateFrom.toDateString(),
        dSecond: hDateTo.toDateString(),
        SiteUrl: this.props.URLAddressToHolidaySite
      }),
      headers: {
        'Content-Type': 'application/json'
      }
    })
      .then(parseFetchResponse)
      .then(({ json, meta }) => {
        if (meta.ok) {

          this.setState((prevState: FormEditState): FormEditState => {
            prevState.holiday.Days = json;
            return prevState;
          })

          let currentEmployee: BaseEmployee = {
            UserName: this.state.employee.UserName,
            UserPosition: this.state.employee.UserPosition
          };

          if (this.state.holiday.Replacement !== undefined) {
            this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, currentEmployee, this.state.holiday.Replacement);
            ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('datesErrorMsg'));
          }
          else {
            ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('replacementErrorMsg'));
          }
        }
        else {
          this.setState((prevState: FormEditState): FormEditState => {
            prevState.holiday.Days = 0;
            return prevState;
          })
          ReactDom.render(<CustomMessage messageType={1} messageText={json} messageVisible={true} />, document.getElementById('datesErrorMsg'));
        }
        return json;
      });
  }

  @autobind
  private _onChangedHolidayDateTo(dateTo: Date): void {
    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.DateTo = dateTo;
      return prevState;
    });

    if (this.state.holiday.DateFrom === undefined)
      return;

    let hDateFrom: Date = new Date(this.state.holiday.DateFrom);
    let hDateTo: Date = dateTo;

    let requestUrl = this.state.urlToWebAPI + strings.ActionCalculateWorkingDays;
    const parseFetchResponse = response => response.json().then(text => ({
      json: text,
      meta: response,
    }));
    window.fetch(requestUrl, {
      method: 'POST',
      body: JSON.stringify({
        dFirst: hDateFrom.toDateString(),
        dSecond: hDateTo.toDateString(),
        SiteUrl: this.state.urlToHolidaySite
      }),
      headers: {
        'Content-Type': 'application/json'
      }
    })
      .then(parseFetchResponse)
      .then(({ json, meta }) => {
        if (meta.ok) {
          this.setState((prevState: FormEditState): FormEditState => {
            prevState.holiday.Days = json;
            return prevState;
          });

          let currentEmployee: BaseEmployee = {
            UserName: this.state.employee.UserName,
            UserPosition: this.state.employee.UserPosition
          };

          if (this.state.holiday.Replacement !== undefined)
            this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, currentEmployee, this.state.holiday.Replacement);
          else {
            ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('replacementErrorMsg'));
          }
          ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('datesErrorMsg'));
        }
        else {
          this.setState((prevState: FormEditState): FormEditState => {
            prevState.holiday.Days = 0;
            return prevState;
          })
          ReactDom.render(<CustomMessage messageType={1} messageText={json} messageVisible={true} />, document.getElementById('datesErrorMsg'));
        }
        return json;
      });
  }

  @autobind
  private _onChangedHolidayReplacement(replacement: BaseEmployee[]): void {

    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.Replacement = replacement[0];
      return prevState;
    });

    if (replacement[0] !== undefined && this.state.holiday.Days != undefined) {

      let currentEmployee: BaseEmployee = {
        UserName: this.state.employee.UserName,
        UserPosition: this.state.employee.UserPosition
      };

      this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, currentEmployee, replacement[0])
    }
  }

  @autobind
  private _onChangedHolidayEmployee(employee: BaseEmployee[]): void {
    if (employee.length !== 0) {
      if (this.state.employee.UserName !== employee[0].UserName) {
        this.setState({
          isEmplDataLoaded: false
        });

        this.getEmployee(this.props.URLAddressToWebAPI + strings.ActionGetNomenclatureUser, employee[0].UserName, employee[0].UserPosition).then((item: Employee) => {
          this.setState({
            employee: item,
            isEmplDataLoaded: true,
          });
        });

      }
    }
  }

  @autobind
  public _onChangedHolidayType(item: IDropdownOption) {
    let selectedItem: string[] = item.key.toString().split('|', 2);

    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.Type = selectedItem[0];
      prevState.holiday.SubType = selectedItem[1];
      return prevState;
    });

  }

  @autobind
  public _onChangedMobile(newValue: string): void {
    this.setState((prevState: FormEditState): FormEditState => {
      prevState.holiday.Mobile = newValue;
      return prevState;
    });

  }

  private getCurrentEmployee(requestUrl: string, userName: string): Promise<Employee> {
    let sUrl = this.props.URLAddressToHolidaySite;
    return new Promise<Employee>((resolve: (item: Employee) => void, reject: (error: any) => void): void => {
      const parseFetchResponse = response => response.json().then(text => ({
        json: text,
        meta: response,
      }))
      window.fetch(requestUrl, {
        method: 'POST',
        body: JSON.stringify({
          SiteUrl: sUrl,
          Username: userName
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
          resolve(item);
        })
        , (error: any): void => { reject(error); }
    });
  }

  private getEmployee(requestUrl: string, userName: string, userPosition: string): Promise<Employee> {
    let employee: BaseEmployee = {
      UserName: userName,
      UserPosition: userPosition
    }
    let sUrl = this.props.URLAddressToHolidaySite;
    return new Promise<Employee>((resolve: (item: Employee) => void, reject: (error: any) => void): void => {
      const parseFetchResponse = response => response.json().then(text => ({
        json: text,
        meta: response,
      }))
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
        , (error: any): void => { reject(error); }
    });
  }

  private getHolidayTypes(requestUrl: string, nomenclatureUrl: string, columnInternalName: string): Promise<IDropdownOption[]> {
    const parseFetchResponse = response => response.json().then(text => ({
      json: text,
      meta: response,
    }));

    return window.fetch(requestUrl, {
      method: 'POST',
      body: JSON.stringify({
        Url: nomenclatureUrl,
        ColumnInternalName: columnInternalName
      }),
      headers: {
        'Content-Type': 'application/json'
      }
    })
      .then(parseFetchResponse)
      .then(({ json, meta }) => {
        const ttt: NomenclatureTypes = json;
        return json;
      })
      .then((items: NomenclatureTypes[]) => {

        const options: IDropdownOption[] = [];
        let couter: number = 0;

        items.forEach(option => {
          couter++;
          if (option.SubType === '') {
            options.push({
              key: couter,
              text: option.Type,
              itemType: 2
            });

            items.forEach(subOption => {
              if (option.Type === subOption.Type && subOption.SubType !== '')
                options.push({
                  key: option.Type + '|' + subOption.SubType,
                  text: subOption.SubType
                });
            });
          }
        });
        return options;
      });
  }

  private checkReplacementIsFree(requestUrl: string, currentEmployee: BaseEmployee, replacementEmployee: BaseEmployee): void {

    ReactDom.render(<Spinner label={strings.spinerCheckReplacementIsFree} />, document.getElementById('replacementErrorMsg'))

    let sUrl = this.props.URLAddressToHolidaySite;
    const parseFetchResponse = response => response.json().then(text => ({
      json: text,
      meta: response,
    }));
    window.fetch(requestUrl, {
      method: 'POST',
      body: JSON.stringify({
        SiteUrl: sUrl,
        Mode: 2, //1 Add, 2 Edit
        dFrom: this.state.holiday.DateFrom,
        dTo: this.state.holiday.DateTo,
        CurrentUser: currentEmployee,
        ReplacementUser: replacementEmployee,
        ItemId: this.state.holiday.Id
      }),
      headers: {
        'Content-Type': 'application/json'
      }
    })
      .then(parseFetchResponse)
      .then(({ json, meta }) => {
        if (meta.ok) {
          ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('replacementErrorMsg'));
        }
        else {
          ReactDom.render(<CustomMessage messageType={1} messageText={json} messageVisible={true} />, document.getElementById('replacementErrorMsg'));
        }
        return json;
      });
  }

  private getSuggestionProps(headerText: string): IBasePickerSuggestionsProps {
    const suggestionPropsReplacement: IBasePickerSuggestionsProps = {
      suggestionsHeaderText: headerText,
      noResultsFoundText: strings.suggestNoResultsFoundText,
      loadingText: strings.suggestloadingText
    };
    return suggestionPropsReplacement;
  }
}


