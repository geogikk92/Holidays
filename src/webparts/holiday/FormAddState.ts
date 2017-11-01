import { IChoiceGroupOption, IDropdownOption, IPersonaProps } from 'office-ui-fabric-react';
import { BaseEmployee } from './models/BaseEmployee';
import { Employee } from './models/Employee';
import { Holiday } from './models/Holiday';

export interface FormAddState {

    urlToWebAPI?: string;
    urlToHolidaySite?: string;
    isEmplDataLoaded?: boolean;
    isHoliday?: boolean;
    loadLastHoliday?: boolean;
    userHasHoliday?: boolean;

    employee?: Employee;
    holiday?: Holiday;
    currentUserName?: string;
    currentUserPosition?: string;
    lastReplacement?: IPersonaProps[];
    // holidayDescription?: string;
    // holidayMobile?: string;
    // holidayType?: string;
    // holidaySubType?: string;
    // holidayDateFrom?: Date;
    // holidayDateTo?: Date;
    // holidayCalcDays?: string;
    // holidayReplacement?: BaseEmployee;
}
