import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';


export interface CustomDropDownProps {
    label: string;
    loadOptions: () => Promise<IDropdownOption[]>;
    onChanged: (option: IDropdownOption, index?: number) => void;
    selectedKey: string | number;
    defaultSelectedKey?: string;
    disabled: boolean;
    stateKey: string;
}