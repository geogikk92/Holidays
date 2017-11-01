import * as React from 'react';
import * as ReactDom from 'react-dom';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './components/Holiday.module.scss';
import * as strings from 'holidayStrings';
import { FormAddProps } from './FormAddProps';
import { FormAddState } from './FormAddState';
import {
  TextField, Image, autobind, Icon, Label, IChoiceGroupOption, IDropdownOption, IPersonaProps, PersonaPresence,
  Spinner, SpinnerSize, SpinnerType, IBasePickerSuggestionsProps, IImageProps, ImageFit, CompoundButton
} from 'office-ui-fabric-react';

import * as pnp from 'sp-pnp-js';

import { Employee } from './models/Employee';
import { Holiday } from './models/Holiday';
import { BaseEmployee } from './models/BaseEmployee';
import { NomenclatureTypes } from './models/NomenclatureTypes';

import { CustomDateTimePicker } from './components/CustomDateTimePicker';
import { CustomTable } from './components/CustomTable';
import { CustomTextField } from './components/CustomTextField';
import { CustomPeoplePicker } from './components/CustomPeoplePicker';
import { CustomButton } from './components/CustomButton';
import { CustomCheckBox } from './components/CustomCheckBox';
import { CustomMessage } from './components/CustomMessage';
import { CustomDropDown } from './components/CustomDropDown';
import { CustomDataSavedSuccess } from './components/CustomDataSavedSuccess';
import { CustomPeoplePickerCommonUtility } from './components/CustomPeoplePickerCommonUtility';
import { Web } from 'sp-pnp-js';


export default class AddForm extends React.Component<FormAddProps, FormAddState> {
  constructor(props) {
    super(props);
    this.state = {
      urlToWebAPI: this.props.URLAddressToWebAPI,
      urlToHolidaySite: this.props.URLAddressToHolidaySite,
      isHoliday: true,
      isEmplDataLoaded: false,
      loadLastHoliday: false,
      userHasHoliday: true,
      currentUserName: this.props.EmpNickName.substring(this.props.EmpNickName.indexOf('\\') + 1),
      currentUserPosition: null,

      employee: new Employee(),
      holiday: new Holiday()
    };
  }

  public componentWillMount() {

    //Проверка дали потребителя има създадени отпуски
    //ако има бутона за използване на отпуск е включен
    this.checkUserHasHoliday(this.state.currentUserName);

    //Използване на RestFull API-то на SharePoint за зареждане на даанните за текущия потребител от списъка 'Служители'
    let nomenclatureWeb = new Web(this.props.URLAddresToNomenclatures);
    nomenclatureWeb.lists.getByTitle('Служители').items.filter(`LirexEmpUserName eq  + '${this.state.currentUserName}'`).top(1).get().then(
      currentEmplData => {
        this.setState((prevState: FormAddState): FormAddState => {
          prevState.employee.FullName = currentEmplData[0].Title;
          prevState.employee.Department = currentEmplData[0].LirexEmplDirection;
          prevState.employee.Days = currentEmplData[0].LirexEmplDays;
          prevState.employee.UserPosition = currentEmplData[0].LirexEmplPosition;
          prevState.currentUserPosition = currentEmplData[0].UserPosition;
          prevState.employee.UserName = this.state.currentUserName;
          prevState.employee.DaysLeftNextYear = currentEmplData[0].NextYearDays;
          prevState.holiday.TotalAllowedDays = currentEmplData[0].LirexEmplDays;
          prevState.holiday.TypeRequest = strings.holiday;
          prevState.isEmplDataLoaded = true;
          return prevState;
        });
      },
      error => {
        ReactDom.render(<CustomMessage messageType={1} messageText={error.message} messageVisible={true} />, document.getElementById('formContent'));
      });

    //Използване на Custrom API за зареждане на даанните за текущия потребител от списъка 'Служители'
    // this.getCurrentEmployee(this.props.URLAddressToWebAPI + strings.ActionGetNomenclatureUserByUsername, this.state.currentUserName).then((item: Employee) => {
    //   this.setState((prevState: FormAddState): FormAddState => {
    //     prevState.employee = item;
    //     prevState.currentUserPosition = item.UserPosition;
    //     prevState.holiday.TotalAllowedDays = item.Days;
    //     prevState.holiday.TypeRequest = strings.holiday;
    //     prevState.isEmplDataLoaded = true;
    //     return prevState;
    //   });
    // });
  }

  public render(): React.ReactElement<FormAddProps> {
    return (
      <div className={styles.helloWorld} >
        <div id="formContent" className={styles.container}>
          {(this.state.isEmplDataLoaded) ?
            <div className="ms-Grid">
              <div className={`ms-Grid-row ms-borderColor-themePrimary ms-fontColor-white ${styles.row}`}>
                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg12">
                  <Label className="ms-fontSize-m ms-fontColor-blueDark">{escape(this.props.Description)}</Label>
                </div>
              </div>
              <CustomCheckBox
                onChanged={this._onChangedType}
                checked={this.state.isHoliday}
                disabled={this.state.loadLastHoliday}
              />
              {(this.state.isHoliday) ?
                <CustomTable
                  employeeData={this.state.employee}
                />
                : null}
              <div className={`ms-Grid-row  ${styles.row}`}>
                {(this.state.isHoliday) ?
                  <CustomButton
                    onClick={this._handleLastHoliday}
                    value={strings.btnUseLastHoliday}
                    loadSpinner={this.state.loadLastHoliday}
                    disabled={this.state.userHasHoliday}
                  />
                  : null}
              </div>
              {(this.state.loadLastHoliday) ?
                null
                :
                <form onSubmit={this._handleSubmit}>
                  <fieldset className={styles.fieldSet}>
                    <CustomPeoplePicker
                      isRequared={true}
                      uName={this.state.currentUserName}
                      uPosition={this.state.currentUserPosition}
                      label={strings.lblEmployee}
                      loadReplacements={false}
                      defaultItem={this.getDefaultPersona()}
                      onChanged={this._onChangedHolidayEmployee}
                      suggestionProps={this.getSuggestionProps(strings.suggestEmployee)}
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
                        isRequared={true}
                        uName={this.state.employee.UserName}
                        uPosition={this.state.employee.UserPosition}
                        label={strings.lblReplacement}
                        onChanged={this._onChangedHolidayReplacement}
                        loadReplacements={true}
                        suggestionProps={this.getSuggestionProps(strings.suggestReplacement)}
                        defaultItem={this.state.lastReplacement}
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
                </form>}
            </div>
            : <Spinner type={SpinnerType.large} size={SpinnerSize.large} label={strings.spinerLoadEmployeeData} />
          }
        </div>
      </div >
    );
  }

  @autobind
  private _onChangedType(): void {
    this.setState((prevState: FormAddState): FormAddState => {
      prevState.isHoliday = !this.state.isHoliday;
      prevState.holiday.TypeRequest = (!this.state.isHoliday) ? strings.holiday : strings.hospital;
      return prevState;
    });
  }

  @autobind
  private _handleLastHoliday(): void {

    this.setState({ loadLastHoliday: true });

    let employee: BaseEmployee = {
      UserName: this.state.employee.UserName,
      UserPosition: this.state.employee.UserPosition
    };

    this.getLastHoliday(this.props.URLAddressToWebAPI + strings.ActionGetLastHolidayByUser, employee, this.state.urlToHolidaySite)
      .then((lastHoliday: Holiday) => {
        this.setState((prevState: FormAddState): FormAddState => {
          prevState.holiday = lastHoliday;
          this.state.holiday.DateFrom = lastHoliday.DateFrom;
          this.state.holiday.DateTo = lastHoliday.DateTo;
          return prevState;
        });

        this.getEmployee(this.props.URLAddressToWebAPI + strings.ActionGetNomenclatureUser, lastHoliday.Replacement.UserName, lastHoliday.Replacement.UserPosition)
          .then((replacementEmpl: Employee) => {
            this.setState((prevState: FormAddState): FormAddState => {
              prevState.lastReplacement = [{
                primaryText: replacementEmpl.FullName,
                secondaryText: replacementEmpl.UserPosition,
                tertiaryText: replacementEmpl.UserName,
                imageInitials: CustomPeoplePickerCommonUtility.getInitials(replacementEmpl.FullName),
                presence: PersonaPresence.none
              }];
              prevState.loadLastHoliday = false;
              return prevState;
            });

            let replacement: BaseEmployee = {
              UserName: replacementEmpl.UserName,
              UserPosition: replacementEmpl.UserPosition
            };

            this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, employee, replacement);
          });
      }).catch(err => {
        ReactDom.render(<CustomMessage messageType={1} messageText={err.message} messageVisible={true} />, document.getElementById('formContent'));
      });
  }

  @autobind
  private _handleSubmit(event): void {
    event.preventDefault();

    let totalDaysForUser: number = this.state.employee.Days + this.state.employee.DaysLeftNextYear;

    let requestor: BaseEmployee = {
      UserName: this.state.currentUserName,
      UserPosition: this.state.currentUserPosition
    };

    let holiday: Holiday = {
      Id: 0,
      Title: '1', //todo: implement get number from webapi by current user. (base user.)
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
    window.fetch(this.props.URLAddressToWebAPI + strings.ActionCreate, {
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
          this.setState((prevState: FormAddState): FormAddState => {
            prevState.employee.Days = prevState.employee.Days - prevState.holiday.Days;
            return prevState;
          });

          ReactDom.render(<CustomDataSavedSuccess messageType={4} messageText="Данните са записани успешно" redirectTo={this.props.URLAddressToHolidaySite} />, document.getElementById('formContent'));
        }
        else {
          ReactDom.render(<CustomMessage messageType={5} messageText={msg} messageVisible={true} />, document.getElementById('formErrorMsg'));
        }
      }, (error: any): void => {
        ReactDom.render(<CustomMessage messageType={5} messageText={error.message} messageVisible={true} />, document.getElementById('formErrorMsg'));
      });
  }

  @autobind
  private _handleRedirect(): void {
    window.location.href = this.props.URLAddressToHolidaySite;
  }

  @autobind
  public _onChangedHolidayAddress(newValue: string): void {
    this.setState((prevState: FormAddState): FormAddState => {
      prevState.holiday.Address = newValue;
      return prevState;
    });
  }

  @autobind
  public _onChangedHolidayDescription(newValue: string): void {
    this.setState((prevState: FormAddState): FormAddState => {
      prevState.holiday.Description = newValue;
      return prevState;
    });
  }

  @autobind
  private _onChangedHolidayDateFrom(dateFrom: Date): void {
    this.setState((prevState: FormAddState): FormAddState => {
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

          this.setState((prevState: FormAddState): FormAddState => {
            prevState.holiday.Days = json;
            return prevState;
          });

          let currentEmployee: BaseEmployee = {
            UserName: this.state.employee.UserName,
            UserPosition: this.state.employee.UserPosition
          };

          if (this.state.holiday.Replacement !== undefined && this.state.isHoliday) {
            this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, currentEmployee, this.state.holiday.Replacement);
            ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('datesErrorMsg'));
          }
          else {
            ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('replacementErrorMsg'));
          }
        }
        else {
          this.setState((prevState: FormAddState): FormAddState => {
            prevState.holiday.Days = 0;
            return prevState;
          });
          ReactDom.render(<CustomMessage messageType={1} messageText={json} messageVisible={true} />, document.getElementById('datesErrorMsg'));
        }
        return json;
      });
  }

  @autobind
  private _onChangedHolidayDateTo(dateTo: Date): void {
    this.setState((prevState: FormAddState): FormAddState => {
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
          this.setState((prevState: FormAddState): FormAddState => {
            prevState.holiday.Days = json;
            return prevState;
          });

          let currentEmployee: BaseEmployee = {
            UserName: this.state.employee.UserName,
            UserPosition: this.state.employee.UserPosition
          };

          if (this.state.holiday.Replacement !== undefined && this.state.isHoliday)
            this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, currentEmployee, this.state.holiday.Replacement);
          else {
            ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('replacementErrorMsg'));
          }
          ReactDom.render(<CustomMessage messageVisible={false} />, document.getElementById('datesErrorMsg'));
        }
        else {
          this.setState((prevState: FormAddState): FormAddState => {
            prevState.holiday.Days = 0;
            return prevState;
          });
          ReactDom.render(<CustomMessage messageType={1} messageText={json} messageVisible={true} />, document.getElementById('datesErrorMsg'));
        }
        return json;
      });
  }

  @autobind
  private _onChangedHolidayReplacement(replacement: BaseEmployee[]): void {

    this.setState((prevState: FormAddState): FormAddState => {
      prevState.holiday.Replacement = replacement[0];
      return prevState;
    });

    if (replacement[0] !== undefined && this.state.holiday.Days !== undefined) {

      let currentEmployee: BaseEmployee = {
        UserName: this.state.employee.UserName,
        UserPosition: this.state.employee.UserPosition
      };

      this.checkReplacementIsFree(this.props.URLAddressToWebAPI + strings.ActionValidateReplacementForPeriod, currentEmployee, replacement[0]);
    }
  }

  @autobind
  private _onChangedHolidayEmployee(employee: BaseEmployee[]): void {
    if (employee.length !== 0) {
      if (this.state.employee.UserName !== employee[0].UserName) {
        this.setState({
          isEmplDataLoaded: false
        });

        this.checkUserHasHoliday(employee[0].UserName);

        let nomenclatureWeb = new Web(this.props.URLAddresToNomenclatures);
        nomenclatureWeb.lists.getByTitle('Служители').items.filter(`LirexEmpUserName eq  + '${employee[0].UserName}'`).top(1).get().then(holidayEmplData => {
          this.setState((prevState: FormAddState): FormAddState => {
            prevState.employee.FullName = holidayEmplData[0].Title;
            prevState.employee.Department = holidayEmplData[0].LirexEmplDirection;
            prevState.employee.Days = holidayEmplData[0].LirexEmplDays;
            prevState.employee.UserPosition = holidayEmplData[0].LirexEmplPosition;
            prevState.currentUserPosition = holidayEmplData[0].UserPosition;
            prevState.employee.UserName = holidayEmplData[0].LirexEmpUserName;
            prevState.employee.DaysLeftNextYear = holidayEmplData[0].NextYearDays;
            prevState.holiday.TotalAllowedDays = holidayEmplData[0].LirexEmplDays;
            prevState.holiday.TypeRequest = strings.holiday;
            prevState.isEmplDataLoaded = true;

            return prevState;
          });
        });


        // this.getEmployee(this.props.URLAddressToWebAPI + strings.ActionGetNomenclatureUser, employee[0].UserName, employee[0].UserPosition).then((item: Employee) => {
        //   this.setState({
        //     employee: item,
        //     isEmplDataLoaded: true,
        //   });
        // });
      }
    }
  }

  @autobind
  public _onChangedHolidayType(item: IDropdownOption) {
    let selectedItem: string[] = item.key.toString().split('|', 2);

    this.setState((prevState: FormAddState): FormAddState => {
      prevState.holiday.Type = selectedItem[0];
      prevState.holiday.SubType = selectedItem[1];
      return prevState;
    });
  }

  @autobind
  public _onChangedMobile(newValue: string): void {
    this.setState((prevState: FormAddState): FormAddState => {
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
      }));
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
        , (error: any): void => { reject(error); };
    });
  }

  private getEmployee(requestUrl: string, userName: string, userPosition: string): Promise<Employee> {
    let employee: BaseEmployee = {
      UserName: userName,
      UserPosition: userPosition
    };

    let sUrl = this.props.URLAddressToHolidaySite;
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
        , (error: any): void => {
          ReactDom.render(<CustomMessage messageType={1} messageText={error.message} messageVisible={true} />, document.getElementById('formContent'));
        };
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

  private getLastHoliday(requestUrl: string, baseEmployee: BaseEmployee, siteURL: string): Promise<Holiday> {
    return new Promise<Holiday>((resolve: (item: Holiday) => void, reject: (error: any) => void): void => {
      const parseFetchResponse = response => response.json().then(text => ({
        json: text,
        meta: response,
      }));
      window.fetch(requestUrl, {
        method: 'POST',
        body: JSON.stringify({
          SiteUrl: siteURL,
          BaseEmp: baseEmployee
        }),
        headers: {
          'Content-Type': 'application/json'
        }
      })
        .then(parseFetchResponse)
        .then(({ json, meta }) => {
          return json;
        })
        .then((item: Holiday) => {
          resolve(item);
        })
        , (error: any): void => {
          ReactDom.render(<CustomMessage messageType={1} messageText={error.message} messageVisible={true} />, document.getElementById('formContent'));
        };
    });
  }

  private checkReplacementIsFree(requestUrl: string, currentEmployee: BaseEmployee, replacementEmployee: BaseEmployee): void {

    ReactDom.render(<Spinner label={strings.spinerCheckReplacementIsFree} />, document.getElementById('replacementErrorMsg'));

    let sUrl = this.props.URLAddressToHolidaySite;
    const parseFetchResponse = response => response.json().then(text => ({
      json: text,
      meta: response,
    }));
    window.fetch(requestUrl, {
      method: 'POST',
      body: JSON.stringify({
        SiteUrl: sUrl,
        Mode: 1, //1 Add, 2 Edit
        dFrom: this.state.holiday.DateFrom,
        dTo: this.state.holiday.DateTo,
        CurrentUser: currentEmployee,
        ReplacementUser: replacementEmployee,
        ItemId: 0
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

  private checkUserHasHoliday(userNickName: string): void {
    this.setState({
      userHasHoliday: true
    });

    pnp.sp.web.lists.getByTitle('Отпуски').items.select('LirexHolidayTypeRequest')
      .filter(`LirexHolidayEmployeeUsername eq  + '${userNickName}'`).get().then((holiday: any) => {
        if (holiday.length != 0)
          holiday.forEach(element => {
            if (element.LirexHolidayTypeRequest === strings.holiday) {
              this.setState({
                userHasHoliday: false
              });
              return;
            }
          });
      }).catch(err => {
        ReactDom.render(<CustomMessage messageType={1} messageText={err.message} messageVisible={true} />, document.getElementById('formContent'));
      });
  }

  private getDefaultPersona(): IPersonaProps[] {

    let result: IPersonaProps[] = [{
      primaryText: this.state.employee.FullName,
      secondaryText: this.state.employee.UserPosition,
      tertiaryText: this.state.employee.UserName,
      imageInitials: CustomPeoplePickerCommonUtility.getInitials(this.state.employee.FullName),
      presence: PersonaPresence.none
    }];

    return result;
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


