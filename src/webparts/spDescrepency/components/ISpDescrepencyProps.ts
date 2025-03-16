import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISpDescrepencyProps {
  context: WebPartContext; // Added context property
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
  userLocalityName: string;
  userFIPS: string;
  allDiscrepancyData: Record<string, IExcelRow[]>; // Added allDiscrepancyData property
  adminFormData: {
    fips: string;
    month: string;
    certifiedCycle: string;
    certifyAccurate: boolean;
    certifyException: boolean;
    adminPrintName: string;
    directorPrintName: string;
    adminSignatureCompleted: boolean;
    directorSignatureCompleted: boolean;
    directorComment: string;
  };
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
  LetsPositions: number;
  DescLetsPositions: string;
  VacantLetsPositions: number;
  DescVacantLETSpositions: string
  FilledLetsPositions: number;
  DescFilledLETSpositions: string;
  EmployeeLetsNotFoundLocal: number;
  VacantPositionsLets: number;
  NumberofLocalPositions: number;
  NumberOfVacantLocalPositions: number;
  NumberOfFilledLocalPositions: number;
  NumberOfEmployeesInLocalNotFoundInLets: number;
  NumberOfEmployeeWithSignificantSalary: number;
  NumberOfLocalPositionsInLETS: number;
  LetsLocalPositionBlank: number;
  NumberOfEmployeeWithPastDueProbation: number;
  NumberOfEmployeeWithPastDueAnnual: number;
  NumberOfEmployeeInExpiredPositions: number;
  NumberOfPositionsWithInvalidRSC: number;
  EmployeeslistedbutNoEESalary: number;
  EmployeeslistedButNoEETimeStatus: number;
  PartTimeEmployeesWithSalary: number;
  FullTimeEmployeesWithHourlyRate: number;
  EmployeesWithDeviationCodePoint5: number;
  EmployeesWithBlankAssignTime: number;
  EmployeeswithBlankEmployeeStatus: number;
}

export interface IDiscrepancyData {
  DiscrepancyName: string;
  Count: number;
}