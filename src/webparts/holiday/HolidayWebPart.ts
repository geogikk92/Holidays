import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'holidayStrings';
import AddForm from './FormAdd';
import EditForm from './FormEdit';
import DisplayForm from './FormDisplay';
import ApproveForm from './FormApprove';
import FormApproveProps from './FormApprove';
import { IHolidayProps } from './IHolidayProps';
import { FormAddProps } from './FormAddProps';
import { FormEditProps } from './FormEditProps';
import { IHolidaysWebPartProps } from './IHolidayWebPartProps';
import * as pnp from 'sp-pnp-js';

export default class HolidaysWebPart extends BaseClientSideWebPart<IHolidaysWebPartProps> {

  public render(): void {
    const addElement: React.ReactElement<FormAddProps> = React.createElement(
      AddForm,
      {
        DisplayMode: this.properties.DisplayMode,
        Description: this.properties.Description,
        EmpNickName: this.context.pageContext.user.loginName,
        URLAddressToWebAPI: this.properties.URLAddressToWebAPI,
        URLAddressToHolidaySite: this.context.pageContext.web.absoluteUrl,
        URLAddresToNomenclatures: this.context.pageContext.web.absoluteUrl + "/Nomenclatures"
      });

    const editElement: React.ReactElement<FormEditProps> = React.createElement(
      EditForm,
      {
        DisplayMode: this.properties.DisplayMode,
        Description: this.properties.Description,
        URLAddressToWebAPI: this.properties.URLAddressToWebAPI,
        URLAddressToHolidaySite: this.context.pageContext.web.absoluteUrl,
        URLAddresToNomenclatures: this.context.pageContext.web.absoluteUrl + "/Nomenclatures",
        ListTitle: "Отпуски"
      });

    const displayElement: React.ReactElement<any> = React.createElement(
      DisplayForm,
      {
        URLAddressToHolidaySite: this.context.pageContext.web.absoluteUrl,
      });

    const approveElement: React.ReactElement<FormApproveProps> = React.createElement(
      ApproveForm,
      {
        URLAddressToHolidaySite: this.context.pageContext.web.absoluteUrl,
      });


    const displayModeKey: string = this.properties.DisplayMode;

    if (displayModeKey === '1') {
      ReactDom.render(addElement, this.domElement);
    }
    else if (displayModeKey === '2') {
      ReactDom.render(editElement, this.domElement);
    }
    else if (displayModeKey === '3') {
      ReactDom.render(displayElement, this.domElement);
    }
    else if (displayModeKey === '4') {
      ReactDom.render(approveElement, this.domElement);
    }
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('DisplayMode', {
                  label: strings.ChoiceAFormForVisualization,
                  options: [
                    { key: '1', text: 'Добавяне' },
                    { key: '2', text: 'Редактиране' },
                    { key: '3', text: 'Преглед' },
                    { key: '4', text: 'Одобряване' }
                  ]
                }),
                PropertyPaneTextField('Description', {
                  label: 'Описание към уеб частта',
                  multiline: true,
                  placeholder: 'Моля въведете описание',
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.AdvanceGroupName,
              groupFields: [
                PropertyPaneTextField('URLAddressToWebAPI', {
                  label: 'URL адрес до Web API',
                  placeholder: 'Въведете URL адрес до Web API',
                  onGetErrorMessage: this.simpleTextBoxValidationMethod,
                }),
                PropertyPaneTextField('URLAddressToHolidaySite', {
                  label: 'URL адрес сайта на отпуски',
                  placeholder: 'Въведете URL адрес досайта на отпуски',
                  onGetErrorMessage: this.simpleTextBoxValidationMethod,
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private simpleTextBoxValidationMethod(value: string): string {
    if (value.length < 5) {
      return "Value must be more than 5 characters!";
    } else {
      return "";
    }
  }
}
