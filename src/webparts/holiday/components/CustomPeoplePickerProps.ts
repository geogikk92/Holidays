import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IPersonaProps } from 'office-ui-fabric-react';
import { BaseEmployee } from '../models/BaseEmployee';
import { IBasePickerSuggestionsProps } from 'office-ui-fabric-react';

export interface CustomPeoplePickerProps {

  suggestionProps: IBasePickerSuggestionsProps;
  label: string;
  placeholder?: string;
  required?: boolean;
  loadReplacements?: boolean;
  uName?: string;
  isRequared?: boolean;
  uPosition?: string;
  defaultItem?: IPersonaProps[];


  onChanged?: (items: BaseEmployee[]) => void;

}