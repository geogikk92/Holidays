import * as React from 'react';


import { IPersonaProps, IBasePickerSuggestionsProps, NormalPeoplePicker, BasePicker, PersonaPresence, Label, autobind, TextField, Spinner, SpinnerType, SpinnerSize } from 'office-ui-fabric-react';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import styles from '../components/Holiday.module.scss';

import { CustomPeoplePickerProps } from './CustomPeoplePickerProps';
import { CustomPeoplePickerState } from './CustomPeoplePickerState';
import { CustomPeoplePickerCommonUtility } from './CustomPeoplePickerCommonUtility';

import { Employee } from '../models/Employee';
import { BaseEmployee } from '../models/BaseEmployee';

export interface IPeoplePickerExampleState {
  currentPicker?: number | string;
  delayResults?: boolean;
}

export class CustomPeoplePicker extends React.Component<CustomPeoplePickerProps, CustomPeoplePickerState> {
  constructor(props: CustomPeoplePickerProps) {
    super(props);

    this.state = {
      items: [],
      personas: [],
      isPickersDisabled: true
    };
  }

  public componentDidMount(): void {
    let suEmployee: string = 'http://sf-spsdev07:42325/api/Employees/GetNomenclatureUser?l=1026';
    let suEmployees: string = 'http://sf-spsdev07:42325/api/Employees/GetNomenclatureEmployeesItemsBySpecificUser?l=1026';

    if (this.props.loadReplacements) {
      this.getEmployee(suEmployee, this.props.uName, this.props.uPosition)
        .then((e: Employee) => {
          this.setState({
            isPickersDisabled: false,
            personas: e.Replacements.map(item => ({
              primaryText: item.FullName,
              secondaryText: item.UserPosition,
              tertiaryText: item.UserName,
              imageInitials: CustomPeoplePickerCommonUtility.getInitials(item.FullName),
              presence: PersonaPresence.none
            }))
          });
        })
        .catch(err => {
          this.setState({
            items: [],
            personas: []
          });
        });
    }
    else {
      this.getEmployees(suEmployees, this.props.uName, this.props.uPosition)
        .then((e: Employee[]) => {
          this.setState({
            isPickersDisabled: false,
            personas: e.map(item => ({
              primaryText: item.FullName,
              secondaryText: item.UserPosition,
              tertiaryText: item.UserName,
              imageInitials: CustomPeoplePickerCommonUtility.getInitials(item.FullName),
              presence: PersonaPresence.none
            }))
          });
        })
        .catch(err => {
          this.setState({
            items: [],
            personas: []
          });
        });
    }
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

  private getEmployees(requestUrl: string, userName: string, userPosition: string): Promise<Employee[]> {
    let employee: BaseEmployee = {
      UserName: userName,
      UserPosition: userPosition
    };

    let sUrl = 'http://sf-spsdev07:34006/sites/BG/SPFxSamples';
    return new Promise<Employee[]>((resolve: (item: Employee[]) => void, reject: (error: any) => void): void => {

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
        .then((item: Employee[]) => {
          const e: Employee[] = item;
          resolve(e);
        })
        , (error: any): void => { reject(error); };
    });
  }

  public render(): React.ReactElement<CustomPeoplePickerProps> {

    return (
      <div >
        <div className={`ms-Grid-row  ${styles.row}`}>
          <Label required={this.props.isRequared} >{this.props.label}</Label>
        </div>
        <div className={`ms-Grid-row  ${styles.row}`}>

          <NormalPeoplePicker
            disabled={this.state.isPickersDisabled}
            defaultSelectedItems={this.props.defaultItem}
            onChange={this._onChangePeoplePicker}
            onResolveSuggestions={this._onFilterChangedPeoplePicker}
            getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
            pickerSuggestionsProps={this.props.suggestionProps}
            className={'ms-PeoplePicker'}
            key={'normal'}
          />
        </div>
      </div>
    );
  }

  @autobind
  private _onChangePeoplePicker(items?: IPersonaProps[]): void {

    /** Empty the array */
    this.state.items = new Array<BaseEmployee>();

    /** Fill it with new items */
    items.forEach((i: IPersonaProps) => {

      let be: BaseEmployee = {
        UserName: i.tertiaryText,
        UserPosition: i.secondaryText
      };

      this.state.items.push(be);
    });
    this.setState(this.state);

    if (this.props.onChanged != null) {
      this.props.onChanged(this.state.items);
    }
  }

  @autobind
  private _onFilterChangedPeoplePicker(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {

    if (filterText && currentPersonas.length == 0) {
      return this.state.personas.filter(p => p.primaryText.toLowerCase().indexOf(filterText.toLowerCase()) === 0 ||
        p.tertiaryText.toLowerCase().indexOf(filterText.toLowerCase()) === 0);
    }
    else {
      return [];
    }
  }
}