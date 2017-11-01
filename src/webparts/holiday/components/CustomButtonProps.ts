export interface CustomButtonProps {
    value?: string;
    type?: string;
    onClick?: (e) => void;
    loadSpinner?: boolean;
    disabled?: boolean;
}