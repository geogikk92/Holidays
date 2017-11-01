export interface CustomTextFieldProps {
    value?: string;
    label?: string;
    isRequared?: boolean;
    isMultiline?: boolean;
    disabled?: boolean;
    defaultValue?: string;
    onChanged?: (newValue: any) => void;
    onErrorMsg?: (newValue: string) => string;
}