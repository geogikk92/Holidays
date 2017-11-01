import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from '../components/Holiday.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { CustomTableProps } from './CustomTableProps';
import { CustomTableState } from './CustomTableState';
import * as strings from 'holidayStrings';

export class CustomTable extends React.Component<CustomTableProps, CustomTableState> {

  public render() {
    return (
      <div className={`ms-Grid-row  ${styles.row}`}>
        <table className={styles.table}>
          <tr>
            <th>{strings.EmplTableHeaderThreeNames}</th>
            <th>{strings.EmplTableHeaderDays}</th>
            <th>{strings.EmplTableHeaderPosition}</th>
            <th>{strings.EmplTableHeaderDepartment}</th>
          </tr>
          <tr>
            <td>{this.props.employeeData.FullName}</td>
            <td>{this.props.employeeData.Days}</td>
            <td>{this.props.employeeData.UserPosition}</td>
            <td>{this.props.employeeData.Department}</td>
          </tr>
        </table>
      </div>
    );
  }
}
