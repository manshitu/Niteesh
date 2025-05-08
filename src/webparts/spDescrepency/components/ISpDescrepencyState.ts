import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISpDescrepencyProps {
  context: WebPartContext; 
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext?: boolean;
  userDisplayName: string;
}

export interface ISpDescrepencyState {
  style: string;
  uploadStatus: string;
  selectedFile?: File;
  agencyOptions: { key: string; text: string }[];
  isLoading: boolean;
  errorMessage: string;
  selectedAgency?: string;
  masterData: IExcelRow[];
  validAgencyData: IExcelRow[];
  currentPage: number;
  recordsPerPage: number;
  isPopupVisible: boolean;
  selectedRow?: IExcelRow;
  descrepencyReport: IDiscrepancyResult[];
  activeTab:
    | "MasterData"
    | "DiscrepancyReport"
    | "DiscrepancyDetails"
    | "Admin"
    | "Director";
  selectedDiscrepancy?: string;
  filteredDiscrepancyData: IExcelRow[];
  isAdmin: boolean;
  isDirector: boolean;
  isHR: boolean;
  showAgencyDropdown: boolean;
  isSaving: boolean;
  saveStatus: string;
  saveAdminStatus: string;
  saveDirectorStatus: string;
  userLocalityName: string;
  userFIPS: string;
  allDiscrepancyData: Record<string, IExcelRow[]>; 
  enableSaveButton: boolean;
  isLateSubmission : boolean;
  adminFormData: {
    fips: string;
    month: string;
    certifiedCycle: string;
    certifyAccurate: boolean;
    certifyException: boolean;
    adminPrintName: string;
    adminDate: string;
    directorPrintName: string;
    directorDate: string;
    adminSignatureCompleted: boolean;
    directorSignatureCompleted: boolean;
    directorComment: string;
    dateLabel: string;
  };
  currentLoggedInUser: string;
}

export interface IExcelRow {
  BureauFIPS?: string;
  PayrollPositionNumber?: string;
  JobTitle?: string;
  StateJobTitle?: string;
  EmployeeLastName?: string;
  EmployeeFirstName?: string;
  EmployeeMiddleInitial?: string;
  Salary?: string;
  FTE?: string;
  ReimbursementPercentage?: string;
  ReimbursementStatusCode?: string;
  LastName?: string;//Added but not used
  FirstName?: string;//Added but not used
  EmployeeSalary?: string;
  PositionTimeStatus?: string;
  EmployeeStatus?: string;
  DeviationCode?: string;
  AssigPercentageTimeToPosition?: string;
  EmployeeExpectedJobEndDate?: string;
  RatingDate?: string;
  ProbationExpectedEndDate?: string;
  StatePositionNumber?: string;
}

export interface IDiscrepancyResult {
  LetsPositions: number; //LETSPositionFilledAndVacant
  DescLetsPositions: string; //DescLETSPositionFilledAndVacant
  VacantLetsPositions: number; //VacantLETSpositions
  DescVacantLetsPositions: string; //DescVacantLETSpositions
  FilledLetsPositions: number; //FilledLETSpositions
  DescFilledLetsPositions: string; //DescFilledLETSpositions
  EmployeeLetsNotFoundLocal: number; //EmployeesinLETSNotFoundLocalFile
  DescEmployeeLetsNotFoundLocal: string; //DescEmployeesinLETSNotFoundLocalFile
  VacantPositionsLets: number; //VacantPositionsInLETSThatMayBeImproperlyVacant
  DescVacantPositionsLets: string; //DescVacantPositionsInLETSThatMayBeImproperlyVacant
  NumberofLocalPositions: number; //LocalPositionsFilledAndVacant
  DescNumberofLocalPositions: string; //DescLocalPositionsFilledAndVacant
  //NumberOfVacantLocalPositions: number; //NOT USED
  //DescNumberOfVacantLocalPositions: string; //NOT USED
  NumberOfFilledLocalPositions: number; //FilledLocalPositions
  DescNumberOfFilledLocalPositions: string; //DescFilledLocalPositions
  NumberOfEmployeesInLocalNotFoundInLets: number; // EmployeesInLocalNotFoundInLETSdata
  DescNumberOfEmployeesInLocalNotFoundInLets: string; //DescEmployeesInLocalNotFoundInLETSdata
  NumberOfEmployeeWithSignificantSalary: number; //EmployeesWithSignificantSalaryVariancesBetweenLETSAndLocalData
  DescNumberOfEmployeeWithSignificantSalary: string; // DescEmployeesWithSignificantSalaryVariancesBetweenLETSAndLocalData
  NumberOfLocalPositionsInLETS: number; // LocalPositionsThatAreAlsoInLETSwithDifferentStateTitles
  DescNumberOfLocalPositionsInLETS: string; //DescLocalPositionsThatAreAlsoInLETSwithDifferentStateTitles
  LetsLocalPositionBlank: number; //LETSLocalPositionIsBlank
  DescLetsLocalPositionBlank: string; //DescLETSLocalPositionIsBlank
  NumberOfEmployeeWithPastDueProbation: number; //EmployeeswithPastDueProbationEndingDate
  DescNumberOfEmployeeWithPastDueProbation: string; //DescEmployeeswithPastDueProbationEndingDate
  NumberOfEmployeeWithPastDueAnnual: number; //EmployeesWithPastDueAnnualEvaluationDate
  DescNumberOfEmployeeWithPastDueAnnual: string; //DescEmployeesWithPastDueAnnualEvaluationDate
  NumberOfEmployeeInExpiredPositions: number; //EmployeesInExpiredPositions
  DescNumberOfEmployeeInExpiredPositions: string; //DescEmployeesInExpiredPositions
  NumberOfPositionsWithInvalidRSC: number;// PositionsWithInvalidRSCValues
  DescNumberOfPositionsWithInvalidRSC: string; //DescPositionsWithInvalidRSCValues
  EmployeeslistedbutNoEESalary: number; //EmployeeslistedbutNoEESalary
  DescEmployeeslistedbutNoEESalary: string; //DescEmployeeslistedbutNoEESalary
  EmployeeslistedButNoEETimeStatus: number; //EmployeeslistedButNoEETimeStatus
  DescEmployeeslistedButNoEETimeStatus: string; //DescEmployeeslistedButNoEETimeStatus
  PartTimeEmployeesWithSalary: number; //PartTimeEmployeesWithSalary
  DescPartTimeEmployeesWithSalary: string; //DescPartTimeEmployeesWithSalary
  FullTimeEmployeesWithHourlyRate: number; //FullTimeEmployeesWithHourlyRate
  DescFullTimeEmployeesWithHourlyRate: string; //DescFullTimeEmployeesWithHourlyRate
  EmployeesWithDeviationCodePoint5: number; //EmployeesWithDeviationCodePoint5
  DescEmployeesWithDeviationCodePoint5: string; //DescEmployeesWithDeviationCodePoint5
  EmployeesWithBlankAssignTime: number; //EmployeesWithBlankAssignTime
  DescEmployeesWithBlankAssignTime: string; //DescEmployeesWithBlankAssignTime
  EmployeeswithBlankEmployeeStatus: number; //EmployeeswithBlankEmployeeStatus
  DescEmployeeswithBlankEmployeeStatus: string; //DescEmployeeswithBlankEmployeeStatus
}
//Match it with Master
export interface IMasterDataItem {
  ID: number;
  FIPS: string;
  REGION: string;
  PersonNumber: string;
  FirstName: string;
  LastName: string;
  MiddleName: string;
  EmployeeStatus: string;
  EmployeePositionBeginDate: string;
  EmployeeSalary: string;
  AssigPercentageTimeToPosition: string;
  StatePositionNumber: string;
  LocalPositionNumber: string;
  OTDCode: string;
  OTD: string;
  DeviationCode: string;
  PositionDuration: string;
  PositionTimeStatus: string;
  PositionStatus: string;
  PositionCLStartDate: string;
  EffectiveDateFrom: string;
  ExpectedPositionEndDate: string;
  PositionEndDate: string;
  ReimbursementStatusCode: string;
  RatingDate: string;
  EmployeeExpectedJobEndDate: string;
  ProbationExpectedEndDate: string;
  
}
export interface IDiscrepancyData {
  DiscrepancyName: string;
  Count: number;
}