import { IChoiceGroupOption, IDropdownOption, IPersonaProps } from 'office-ui-fabric-react';
import { BaseEmployee } from './models/BaseEmployee';
import { Employee } from './models/Employee';
import { Holiday } from './models/Holiday';

export interface FormEditState {

    urlToWebAPI?: string;
    urlToHolidaySite?: string;
    isEmplDataLoaded?: boolean;
    isHoliday?: boolean;
    ListName?: string;

    employee?: Employee;
    holiday?: Holiday;
    currentUserName?: string;
    currentUserPosition?: string;
    Replacement?: IPersonaProps[];
}
