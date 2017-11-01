declare interface IHolidayStrings {

  ActionGetNomenclatureUserByUsername: string;
  ActionGetLastHolidayByUser: string;
  ActionGetNomenclatureUser: string;
  ActionValidateReplacementForPeriod: string;
  ActionCreate: string;
  ActionUpdate: string;
  ActionCalculateWorkingDays: string;
  ActionTypeSubtypeItems: string;

  holiday: string;
  hospital: string;

  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AdvanceGroupName: string;
  ChoiceAFormForVisualization: string;

  ChoiceHolidayOrHospital: string;
  ChoiceHoliday: string;
  ChoiceHospital: string;

  EmplTableHeaderThreeNames: string;
  EmplTableHeaderDays: string;
  EmplTableHeaderPosition: string;
  EmplTableHeaderDepartment: string;

  btnUseLastHoliday: string;
  lblEmployee: string;
  lblHolidayTypes: string;
  lblAddress: string;
  lblDescription: string;
  lblDateFrom: string;
  lblDateTo: string;
  lblMobile: string;
  lblTotalDays: string;
  lblReplacement: string;
  suggestReplacement: string;
  suggestEmployee: string;
  suggestNoResultsFoundText: string;
  suggestloadingText: string;
  btnSubmit: string;
  btnClear: string;

  spinerLoadEmployeeData: string;
  spinerCheckReplacementIsFree: string;

}

declare module 'holidayStrings' {
  const strings: IHolidayStrings;
  export = strings;
}
