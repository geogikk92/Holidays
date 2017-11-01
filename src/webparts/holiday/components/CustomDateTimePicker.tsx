import * as React from 'react';

import { BaseComponent, assign, autobind } from 'office-ui-fabric-react/lib/Utilities';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Label, TextField } from 'office-ui-fabric-react/';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import * as strings from 'holidayStrings';

import { CustomDateTimePickerProps } from './CustomDateTimePickerProps';
import { CustomDateTimePickerState } from './CustomDateTimePickerState';

import styles from '../components/Holiday.module.scss';



const DayPickerStrings: IDatePickerStrings = {
  months: [
    'Януари',
    'Февруари',
    'Март',
    'Април',
    'Май',
    'Юни',
    'Юли',
    'Август',
    'Септември',
    'Октомври',
    'Ноември',
    'Декември'
  ],

  shortMonths: [
    'Ян',
    'Фев',
    'Мар',
    'Апр',
    'Май',
    'Юни',
    'Юли',
    'Авг',
    'Сеп',
    'Окт',
    'Ное',
    'Дек'
  ],

  days: [
    'Неделя',
    'Понеделник',
    'Вторник',
    'Сряда',
    'Четвъртък',
    'Петък',
    'Събота'
  ],

  shortDays: [
    'Н',
    'П',
    'В',
    'С',
    'Ч',
    'П',
    'С'
  ],

  goToToday: 'Днес',
  isRequiredErrorMessage: 'Полето е задължително.',
  invalidInputErrorMessage: 'Некоректен формат на датата.'
};

export class CustomDateTimePicker extends React.Component<CustomDateTimePickerProps, CustomDateTimePickerState> {
  constructor(props) {
    super(props);
    this.state = {
      date: (this.props.initialDateTime !== undefined) ? new Date(this.props.initialDateTime) : null
    };
  }

  public render(): React.ReactElement<CustomDateTimePickerProps> {
    this.state = {
      date: (this.props.initialDateTime !== undefined) ? new Date(this.props.initialDateTime) : null
    };

    return (
      <div>
        <div className={`ms-Grid-row  ${styles.row}`}>
          <Label required={true} disabled={false}>{this.props.label}</Label>
        </div>
        <div className="ms-Grid-row">
          <div className={"ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"}>
            <DatePicker
              allowTextInput={false}
              value={this.state.date}
              onSelectDate={this._dateSelected}
              isRequired={this.props.isRequired}
              formatDate={(date: Date) => (date.getDate() + '.' + (date.getMonth() + 1) + '.' + date.getFullYear())}
              firstDayOfWeek={DayOfWeek.Monday}
              strings={DayPickerStrings}
            />
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private _dateSelected(date: Date): void {
    if (date == null)
      return;
    this.state.date = date;
    this.setState(this.state);
    this.saveFullDate();
  }


  private saveFullDate(): void {
    if (this.state.date == null)
      return;
    var finalDate = new Date(this.state.date.toDateString());

    if (finalDate != null) {
      var finalDateAsString: string = '';
      if (this.props.formatDate) {
        finalDateAsString = this.props.formatDate(finalDate);
      }
      else {
        finalDateAsString = finalDate.toString();
      }
    }
    this.state.fullDate = finalDateAsString;
    this.setState(this.state);

    if (this.props.onChanged != null) {
      this.props.onChanged(finalDate);
    }
  }
}