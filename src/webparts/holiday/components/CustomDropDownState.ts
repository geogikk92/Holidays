import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface CustomDropDownState {
  loading?: boolean;
  options?: IDropdownOption[];
  error?: string;
  defaultValue?: string;
}