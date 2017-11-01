import { IPersonaProps, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { BaseEmployee } from '../models/BaseEmployee';

export interface CustomPeoplePickerState {
    items?: BaseEmployee[];
    personas?: IPersonaProps[];
    isPickersDisabled?: boolean;
}