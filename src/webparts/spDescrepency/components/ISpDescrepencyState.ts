export interface ISpDescrepencyState {
  style: string;
  uploadStatus: string;
  selectedFile?: File;
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
  adminFormData: {
    fips: string;
    month: string;
    certifiedCycle: string;
    certifyAccurate: boolean;
    certifyException: boolean;
    adminPrintName: string;
    directorPrintName: string;
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
}

export interface IDiscrepancyResult {
  LetsPositions: number;
  VacantLetsPositions: number;
  FilledLetsPositions: number;
  EmployeeLetsNotFoundLocal: number;
  VacantPositionsLets: number;
  NumberofLocalPositions: number;
  NumberOfVacantLocalPositions: number;
  NumberOfFilledLocalPositions: number;
  NumberOfEmployeeWithSignificantSalary: number;
  NumberOfLocalPositionsInLETS: number;
  LetsLocalPositionBlank: number;
  NumberOfEmployeeWithPastDueProbation: number;
  NumberOfEmployeeWithPastDueAnnual: number;
  NumberOfEmployeeInExpiredPositions: number;
  NumberOfPositionsWithInvalidRSC: number;
}
