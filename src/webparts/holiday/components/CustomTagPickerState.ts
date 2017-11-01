export interface CustomTagPickerState {
    isPickerDisabled: boolean;
    tags?: CustomTags[];
}

export interface CustomTags {
    key: string;
    name: string;
} 