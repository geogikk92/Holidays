export interface CustomDateTimePickerProps {
    isRequired?: boolean;
    label?: string;
    initialDateTime?: Date;
    formatDate?: (date: Date) => string;
    onChanged?: (newValue: Date) => void;
}