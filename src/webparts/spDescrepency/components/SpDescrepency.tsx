import * as React from "react";
import styles from "./SpDescrepency.module.scss";
import type { ISpDescrepencyProps } from "./ISpDescrepencyProps";
import { sp } from "@pnp/sp/presets/all";
import { read, utils } from "xlsx";
import {
  Spinner,
  SpinnerSize,
  //MessageBar,
  //MessageBarType,
  Dropdown,
  IDropdownOption,
  //Dialog,
  //DialogType,
  DetailsList,
  IColumn
} from "office-ui-fabric-react";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SelectionMode } from "@fluentui/react";

interface ISpDescrepencyState {
  style: string;
  uploadStatus: string;
  selectedFile: File | undefined;
  isLoading: boolean;
  errorMessage: string;
  selectedAgency: string | undefined;
  masterData: IExcelRow[];
  validAgencyData: IExcelRow[];
  currentPage: number;
  isPopupVisible: boolean;
  selectedRow: IExcelRow | undefined;
  descrepencyReport: IDiscrepancyResult[];
  activeTab: "MasterData" | "DiscrepancyReport" | "DiscrepancyDetails" | "Admin" | "Director";
  selectedDiscrepancy?: string | undefined; // Stores selected discrepancy for Tab 3
  filteredDiscrepancyData: IExcelRow[]; // Stores the filtered data for details tab
  isAdmin: boolean;
  isDirector: boolean;
  isHR: boolean;  
  showAgencyDropdown: boolean; // Determines whether to show the dropdown
  isSaving: boolean; // Track save operation
  saveStatus: string; // Show success/error message
  userLocalityName: string; // Stores user's locality name
  adminFormData: { // Stores Admin tab form data    
    fips: string;
    month: string;
    certifiedCycle: string;
    certifyAccurate: boolean;
    certifyException: boolean;
    printName: string;
  };
}

interface IExcelRow {
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

interface IDiscrepancyResult {
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

export default class SpDescrepency extends React.Component<
  ISpDescrepencyProps,
  ISpDescrepencyState
> {
  private agencyOptions: IDropdownOption[] = [
    { key: "1", text: "Agency 1" },
    { key: "2", text: "Agency 2" },
    { key: "3", text: "Agency 3" },
  ];

  constructor(props: ISpDescrepencyProps) {
    super(props);
    this.state = {
      style: "",
      uploadStatus: "",
      selectedFile: undefined,
      isLoading: false,
      errorMessage: "",
      selectedAgency: undefined,
      masterData: [],
      validAgencyData: [],
      currentPage: 1,
      isPopupVisible: false,
      selectedRow: undefined,
      descrepencyReport: [],
      activeTab: "MasterData",
      selectedDiscrepancy: undefined,
      filteredDiscrepancyData: [],
      isAdmin: false,
      isDirector: false,
      isHR: false,
      showAgencyDropdown: false, // Default is hidden, will be updated later
      isSaving: false,
      saveStatus: "",
      userLocalityName: "", // Stores user's locality name
      adminFormData: {
        fips: "",
        month: "",
        certifiedCycle: "",
        certifyAccurate: false,
        certifyException: false,
        printName: "",
      },
    };

    sp.setup({
      spfxContext: this.props.context as IWebPartContext,
    });
  }

  /*
  private resetUI = (): void => {
    console.log("Resetting UI...");
    this.setState({
      selectedAgency: "",
      //uploadStatus: "",
      errorMessage: "",
    });
    const fileInput = document.querySelector(
      `.${styles.fileInput}`
    ) as HTMLInputElement;
    if (fileInput) {
      fileInput.value = "";
    }
  };
  */

  public async componentDidMount(): Promise<void> {    
    await this.checkUserAccess(); // Fetch user role and agency name on load    
    await this.fetchDiscrepancyReportForCurrentMonth(); // Load discrepancy data if available
    await this.fetchAdminFormData(); // Load admin form data if available
  }

  private checkUserAccess = async (): Promise<void> => {    
    try {
      // Get current user's email
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email.toLowerCase();
  
      // Fetch user details from SharePoint list
      const listName = "LDSSProfileSummary"; // Replace with your actual list name
      const items = await sp.web.lists.getByTitle(listName).items.select("*").get();
  
      // Find matching user record
      const userRecord = items.find(item => 
        item.PrimaryAdminEmail?.toLowerCase() === currentUserEmail ||
        item.SecondaryAdminEmail?.toLowerCase() === currentUserEmail ||        
        item.DirectorAsstDirectorEmail?.toLowerCase() === currentUserEmail ||
        item.HREmail?.toLowerCase() === currentUserEmail
      );
  
      if (userRecord) {
        const isAdmin = (userRecord.PrimaryAdminEmail?.toLowerCase() === currentUserEmail) || (userRecord.SecondaryAdminEmail?.toLowerCase() === currentUserEmail);
        const isDirector = userRecord.DirectorAsstDirectorEmail?.toLowerCase() === currentUserEmail;
        const isHR = userRecord.HREmail?.toLowerCase() === currentUserEmail;
        const defaultAgency = userRecord.Title || ""; // Get agency name (Title column)

        this.setState({
          isAdmin: isAdmin, //userRecord.PrimaryAdminEmail?.toLowerCase() === currentUserEmail,
          //isAdmin: (userRecord.PrimaryAdminEmail?.toLowerCase() === currentUserEmail) || (userRecord.SecondaryAdminEmail?.toLowerCase() === currentUserEmail),
          isDirector: isDirector, //!userRecord.DirectorAsstDirectorEmail && userRecord.DirectorAsstDirectorEmail?.toLowerCase() === currentUserEmail,
          isHR: isHR, // Check if user is HR
          //agencyName: defaultAgency, // Get agency name (Title column)
          //selectedAgency: '02',
          selectedAgency: defaultAgency, // Set default agency
          showAgencyDropdown: isHR, // Show dropdown only if HR 
          userLocalityName: userRecord.field_1 || "", // Get locality name
               
        },() => {
          if (!isHR && defaultAgency) {
             // eslint-disable-next-line no-void
             void this.fetchMasterAgencyData(defaultAgency); // Auto-load data if not HR
          }
        }
      );  
      }
    } catch (error) {
      console.error("Error fetching user access:", error);
    }
  };
    
  private handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (file) {
      this.setState({
        selectedFile: file,
        uploadStatus: "File selected. Click 'Show Report' to proceed.",
        style: styles.alertMessage,
      });
    } else {
      this.setState({
        style: styles.errorMessage,
        uploadStatus: "No file selected. Please choose a valid .xlsx file.",
      });
    }
  };

  private handleValidateDescrepencyClick = async (): Promise<void> => {
    const { selectedFile, selectedAgency, masterData } = this.state;

    if (!selectedFile) {
      this.setState({
        style: styles.errorMessage,
        uploadStatus: "No file selected. Please choose a valid .xlsx file.",
      });
      return;
    }

    if (!selectedAgency) {
      this.setState({
        style: styles.errorMessage,
        uploadStatus: "Please select an agency before uploading the file.",
      });
      return;
    }

    try {
      this.setState({
        style: styles.alertMessage,
        uploadStatus: "Processing file, please wait...",
        isLoading: true,
        errorMessage: "",
      });

      // Read and parse the Excel file
      const data = await this.readExcel(selectedFile);

      // Validate the data
      const { validRows, invalidRows } = this.validateExcelData(data);

      if (invalidRows.length > 0) {
        console.warn("Invalid rows:", invalidRows);
      }

      // Upload the file to the SharePoint library
      await this.uploadFileToLibrary(selectedFile);

      // Save the valid data to the SharePoint list
      await this.saveDataToList(validRows);

      this.setState({
        style: styles.successMessage,
        uploadStatus: `File uploaded successfully! ${validRows.length} rows saved. ${invalidRows.length} rows skipped due to validation errors.`,
      });

      // Calculating Discrepancies based on userData and masterData
      const discrepancyResult = this.calculateDiscrepancies(validRows, masterData);

      // Display the discrepancy report
      this.displayDiscrepancyReport([discrepancyResult]);

      // Reset the UI
      //this.resetUI();
    } catch (error) {
      console.error("Error during file upload:", error);
      this.setState({
        errorMessage:
          error.message ||
          "An error occurred while processing the file. Please try again.",
      });
    } finally {
      this.setState({ isLoading: false });
    }
  };

  private calculateDiscrepancies = (validRows: IExcelRow[], masterData: IExcelRow[]): IDiscrepancyResult => {
    const letsPositions = masterData.length;
    const vacantLetsPositions = masterData.filter((master) => master.EmployeeFirstName).length;
    const filledLetsPositions = masterData.filter((master) => !master.EmployeeFirstName).length;

    const employeeLetsNotFoundLocal = 0; //validRows.filter((agency) => !agency.EmployeeFirstName).length;
    const vacantPositionsLets = 0; //masterData.filter((master) => !master.EmployeeFirstName).length;

    const numberofLocalPositions = validRows.length;
    const numberOfVacantLocalPositions = validRows.filter((agency) => agency.EmployeeFirstName).length;
    const numberOfFilledLocalPositions = validRows.filter((agency) => !agency.EmployeeFirstName).length;

    const numberOfEmployeeWithSignificantSalary = 0;
    const numberOfLocalPositionsInLETS = 0;
    const letsLocalPositionBlank = 0;
    const numberOfEmployeeWithPastDueProbation = 0;
    const numberOfEmployeeWithPastDueAnnual = 0;
    const numberOfEmployeeInExpiredPositions = 0;
    const numberOfPositionsWithInvalidRSC = 0;

    return {
      LetsPositions: letsPositions,
      VacantLetsPositions: vacantLetsPositions,
      FilledLetsPositions: filledLetsPositions,
      EmployeeLetsNotFoundLocal: employeeLetsNotFoundLocal,
      VacantPositionsLets: vacantPositionsLets,
      NumberofLocalPositions: numberofLocalPositions,
      NumberOfVacantLocalPositions: numberOfVacantLocalPositions,
      NumberOfFilledLocalPositions: numberOfFilledLocalPositions,
      NumberOfEmployeeWithSignificantSalary: numberOfEmployeeWithSignificantSalary,
      NumberOfLocalPositionsInLETS: numberOfLocalPositionsInLETS,
      LetsLocalPositionBlank: letsLocalPositionBlank,
      NumberOfEmployeeWithPastDueProbation: numberOfEmployeeWithPastDueProbation,
      NumberOfEmployeeWithPastDueAnnual: numberOfEmployeeWithPastDueAnnual,
      NumberOfEmployeeInExpiredPositions: numberOfEmployeeInExpiredPositions,
      NumberOfPositionsWithInvalidRSC: numberOfPositionsWithInvalidRSC,
    };
  };

  private displayDiscrepancyReport = (discrepancies: IDiscrepancyResult[]): void => {
    if (!discrepancies) {
      alert("No discrepancies found. Data matches the master database.");
      return;
    }

    this.setState({
      descrepencyReport: discrepancies,
      activeTab: "DiscrepancyReport", // Switch to Discrepancy Report tab
    });
  };

  private handleAgencyChange = async (agency: string): Promise<void> => {
    this.setState({ selectedAgency: agency });
    await this.fetchMasterAgencyData(agency);
  };

  private fetchMasterAgencyData = async (agency: string): Promise<void> => {
    const listName = "PRS_Master_Data"; // Replace with your list name
    try {
      const items = await sp.web.lists
        .getByTitle(listName)
        .items.filter(`field_4 eq '${agency}'`) // Filter based on the selected agency
        .select("*")
        .top(100) // Adjust the number of rows to fetch
        .get();

      const masterData = items.map((item) => ({
        BureauFIPS: item.Title,
        Region: item.field_3, //StateJobTitle
        PersonNumber: item.field_5, //StateJobTitle
        FirstName: item.field_6, //EmployeeLastName
        LastName: item.field_7, //EmployeeFirstName
        MiddleName: item.field_28,
        FIPS: item.field_4,
        EmployeeStatus: item.field_8,
        EmployeePositionBeginDate: item.field_9,
        EmployeeSalary: item.field_10,
        AssigPercentageTimeToPosition: item.field_11,
        StatePositionNumber: item.field_12,
        LocalPositionNumber: item.field_13,
        OTD: item.field_15,
        OTDCode: item.field_14,
        DeviationCode: item.field_16,
        PositionDuration: item.field_17,
        PositionTimeStatus: item.field_18,
        PositionStatus: item.field_19,
        PositionCLStartDate: item.field_20,
        EffectiveDateFrom: item.field_21,
        ExpectedPositionEndDate: item.field_22,
        PositionEndDate: item.field_23,
        ReimbursementStatusCode: item.field_24,
        RatingDate: item.field_25,
        EmployeeExpectedJobEndDate: item.field_26,
        ProbationExpectedEndDate: item.field_27,
      }));
      this.setState({ masterData });
    } catch (error) {
      console.error("Error fetching list data: ", error);
      throw new Error("Failed to fetch data from the list.");
    }
  };

  private fetchAdminFormData = async (): Promise<void> => {
    try {
      const listName = "CertificationReportData"; // Replace with your actual SharePoint list name
  
      const currentUser = await sp.web.currentUser.get();
      const currentUserName = currentUser.Title;

      // Get current month and year
      const currentDate = new Date();
      const currentMonth = String(currentDate.getMonth() + 1).padStart(2, "0"); // "01" to "12"
      const currentYear = String(currentDate.getFullYear());
  
      // Query SharePoint to get admin form data for the current month and agency
      const items = await sp.web.lists
        .getByTitle(listName)
        .items
        .filter(`field_1 eq '${this.state.selectedAgency}' and field_2 eq '${currentMonth}' and field_3 eq '${currentYear}'`)
        .select("*")
        .get();
  
        if (items.length > 0) {
          // Found data in SharePoint, set form state
          const formData = items[0];
          this.setState({
            adminFormData: {
              fips: formData.field_1 || "",
              month: formData.field_2,
              certifiedCycle: formData.field_3,
              certifyAccurate: formData.CertifyAccurate,
              certifyException: formData.CertifyException,
              printName: formData.AdminPrintName,
            },
          });
        } else {
          // No data found, set default values
          this.setState({
            adminFormData: {
              fips: "00000", // Default FIPS code
              month: currentMonth,
              certifiedCycle: currentMonth + "/" + currentYear,
              certifyAccurate: false, // Default unchecked
              certifyException: false, // Default unchecked
              printName: currentUserName, // Pre-fill with current user's name
            },
          });
        }
    } catch (error) {
      console.error("Error fetching admin form data:", error);
    }
  };

  private renderAgencyDropdown = (): JSX.Element | null => {
    if (!this.state.showAgencyDropdown) return null;
  
    return (
      <Dropdown
        label="Select Agency Name"
        placeholder="Select an agency"
        options={this.agencyOptions}
        onChange={(_, option) => this.handleAgencyChange(option?.key as string)}
        selectedKey={this.state.selectedAgency}
        className={styles.dropdown}
      />
    );
  };
  
  public renderMasterDataGrid(): JSX.Element {
    //const { masterData, isLoading } = this.state;
    const { masterData } = this.state;

    if (this.state.selectedAgency === undefined) {
      return <p>Please select agency to see respective data.</p>;
    } else if (this.state.selectedAgency && masterData.length === 0) {
      return <p>No data available for the selected agency.</p>;
    }

    return (
      <div className={styles.gridContainer}>
        <DetailsList
          items={masterData}
          columns={this.columns}
          selectionMode={SelectionMode.single} // Enforces single row selection
          //onItemInvoked={this.handleShowDetailsClick} // Handles row click
          onActiveItemChanged={this.handleShowDetailsClick}
          compact={true} // Optional: makes the grid more compact
        />
      </div>
    );
  }

  private renderSelectedDiscrepancyDetails = (): JSX.Element => {
    const { selectedDiscrepancy, filteredDiscrepancyData } = this.state;    
  
    if (!selectedDiscrepancy) {
      return <p>Please select a discrepancy from the report.</p>;
    }
  
    if (filteredDiscrepancyData.length === 0) {
      return <p>No records found for selected discrepancy.</p>;
    }
  
    // Define dynamic columns based on the discrepancy type
    let columns: IColumn[] = [];
  
    if (selectedDiscrepancy === "LetsPositions") {
      columns = [
        { key: "1", name: "Employee Name", fieldName: "FirstName", minWidth: 100, maxWidth: 150 },
        { key: "2", name: "Position Number", fieldName: "LocalPositionNumber", minWidth: 100, maxWidth: 150 },
        { key: "3", name: "Salary", fieldName: "EmployeeSalary", minWidth: 80, maxWidth: 120 },
      ];
    } else if (selectedDiscrepancy === "VacantLetsPositions") {
      columns = [
        { key: "1", name: "BureauFIPS", fieldName: "BureauFIPS", minWidth: 100, maxWidth: 150 },
        { key: "2", name: "Job Title", fieldName: "Region", minWidth: 100, maxWidth: 150 },
      ];
    } else if (selectedDiscrepancy === "FilledLetsPositions") {
      columns = [
        { key: "1", name: "Employee Name", fieldName: "FirstName", minWidth: 100, maxWidth: 150 },
        { key: "2", name: "Status", fieldName: "EmployeeStatus", minWidth: 100, maxWidth: 150 },
      ];
    }
  
    return (
      <div className={styles.tabContent}>
        <h3>Details for {selectedDiscrepancy}</h3>
        <DetailsList
          items={filteredDiscrepancyData}
          columns={columns}
          selectionMode={SelectionMode.none}
          compact={true}
        />                  
      </div>
    );
  };    
  
  private saveDiscrepancyReportToSharePoint = async (data: IDiscrepancyResult[]): Promise<void> => {
    if (data.length === 0) {
      this.setState({ saveStatus: "No data to save." });
      return;
    }

    // Get current month in two-digit format (e.g., "02" for February)
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1;//String(currentDate.getMonth() + 1).padStart(2, "0"); // Ensures "01" to "12"
    const currentYear = currentDate.getFullYear(); // Example: 2025
  
    const listName = "LocalWorkforceReconciliationSummary"; // Ensure this matches your SharePoint list name    
    const discrepencyList = sp.web.lists.getByTitle(listName);
    this.setState({ isSaving: true, saveStatus: "Saving to SharePoint..." });
  
    try {
      await Promise.all(data.map(async (item) => {
        await discrepencyList.items.add({
          //Title: item.LetsPositions, // Assuming 'Title' stores the discrepancy name
          field_1: this.state.selectedAgency,
          field_2: currentYear,
          field_3: currentMonth,
          field_4: item.LetsPositions,
          field_5: item.VacantLetsPositions,
          field_6: item.FilledLetsPositions,
          field_7: item.EmployeeLetsNotFoundLocal,
          field_8: item.VacantPositionsLets,
          field_9: item.NumberofLocalPositions,
          field_10: item.NumberOfVacantLocalPositions,
          field_11: item.NumberOfFilledLocalPositions,
          field_12: item.NumberOfFilledLocalPositions,
          field_13: item.NumberOfEmployeeWithSignificantSalary,
          field_14: item.NumberOfLocalPositionsInLETS,
          field_15: item.LetsLocalPositionBlank,
          field_16: item.NumberOfEmployeeWithPastDueProbation,
          field_17: item.NumberOfEmployeeWithPastDueAnnual,
          field_18: item.NumberOfEmployeeInExpiredPositions,
          field_19: item.NumberOfPositionsWithInvalidRSC
          //DateReported: new Date().toISOString() // Adding timestamp
        });
      }));
  
      this.setState({ isSaving: false, saveStatus: "Discrepancy report saved successfully!" });
    } catch (error) {
      console.error("Error saving discrepancy report to SharePoint:", error);
      this.setState({ isSaving: false, saveStatus: "Error saving to SharePoint. Please try again." });
    }
  };

  private fetchDiscrepancyReportForCurrentMonth = async (): Promise<void> => {
    const { selectedAgency } = this.state;
    try {
      const listName = "LocalWorkforceReconciliationSummary"; // Ensure this matches your SharePoint list name
  
      // Get current month in "02" format and year as string
      const currentDate = new Date();
      const currentMonth = currentDate.getMonth() + 1; // Ensures "01" to "12"
      const currentYear = currentDate.getFullYear(); // Convert to string
  
      // Query SharePoint to get discrepancy data for the current month and year
      const items = await sp.web.lists
        .getByTitle(listName)
        .items
        .filter(`field_1 eq '${selectedAgency}' and field_3 eq '${currentMonth}' and field_2 eq '${currentYear}'`)
        .select("*")
        .get();
  
      if (items.length > 0) {
        // Map SharePoint data to IDiscrepancyResult structure
        const descrepencyReport: IDiscrepancyResult[] = items.map(item => ({
          DiscrepancyName: item.Title, // Assuming Title holds the discrepancy name
          LetsPositions: item.field_4,
          VacantLetsPositions: item.field_5,
          FilledLetsPositions: item.field_6,
          EmployeeLetsNotFoundLocal: item.field_7,
          VacantPositionsLets: item.field_8,
          NumberofLocalPositions: item.field_9,
          NumberOfVacantLocalPositions: item.field_10,
          NumberOfFilledLocalPositions: item.field_11,
          //NumberOfEmployeesInLocalNotFoundInLets: Number(item.NumberOfEmployeesInLocalNotFoundInLets) || 0        
          NumberOfEmployeeWithSignificantSalary: item.field_12,
          NumberOfLocalPositionsInLETS: item.field_13,
          LetsLocalPositionBlank: item.field_14,
          NumberOfEmployeeWithPastDueProbation: item.field_15,
          NumberOfEmployeeWithPastDueAnnual: item.field_16,
          NumberOfEmployeeInExpiredPositions: item.field_17,
          NumberOfPositionsWithInvalidRSC: item.field_18
        }));
          
        // Update state with fetched data
        this.setState({ descrepencyReport: descrepencyReport });
      }
    } catch (error) {
      console.error("Error fetching discrepancy report from SharePoint:", error);
    }
  };

  private renderForm = (): JSX.Element => {
    //const { adminFormData, isSaving, saveStatus } = this.state;
    const { adminFormData } = this.state;
    return (      
      <div>
        {/* Locality Information */}
        <div className={styles.formGroup}>
          <label>Locality Name (City or County):</label>          
          <input type="text" className={styles.formInput} value={this.state.userLocalityName} placeholder="Enter Locality Name"
            onChange={(e) => this.setState({ userLocalityName: e.target.value })} 
          />
          <label>FIPS:</label>
          <input type="text" className={styles.formInputSmall} value={adminFormData.fips}
            onChange={(e) => this.setState({ adminFormData: { ...adminFormData, fips: e.target.value } })} 
          />          
          <label>Month:</label>
          <input type="text" className={styles.formInputSmall} value={adminFormData.month}
            onChange={(e) => this.setState({ adminFormData: { ...adminFormData, month: e.target.value } })} 
          />          
        </div>
    
        <div className={styles.formGroup}>
          <label>Cycle being Certified (Month/Year) (example: 05/2023):</label>
          <input type="text" className={styles.formInput} value={adminFormData.certifiedCycle}
            onChange={(e) => this.setState({ adminFormData: { ...adminFormData, certifiedCycle: e.target.value } })} 
          />
        </div>
  
        <div className={styles.formGroup}>          
          <p>
            <div className={styles.subHeading}>Position Reimbursement & Status Report for corresponding LHRC Certification Period</div>
            This report provides employees, positions, and reimbursement status information.  
            Agencies are responsible for ensuring that the information is accurate.  Additional reference 
            resources are available on FUSION here : 
          <a href="https://fusion.dss.virginia.gov/hr/HR-Home/Local-Agency-Home/Local-HR-Connect-Project" target="_blank" rel="noreferrer">Local HR Connect Information</a>
        </p>
        </div>
  
        {/* Certification Section */}
        <h4>Certify the Report</h4>
        <p>Check the appropriate box (double-click the box for selection options):</p>
  
        <div className={styles.checkboxGroup}>
          <label>
          <input type="checkbox" className={styles.checkbox} 
              checked={adminFormData.certifyAccurate} 
              onChange={(e) => this.setState({ adminFormData: { ...adminFormData, certifyAccurate: e.target.checked } })}
            />
            I have reviewed the LHRC Position Reimbursement & Status Report and I certify that it is accurate.
          </label>
        </div>
  
        <div className={styles.checkboxGroup}>
          <label>
          <input type="checkbox" className={styles.checkbox} 
              checked={adminFormData.certifyException} 
              onChange={(e) => this.setState({ adminFormData: { ...adminFormData, certifyException: e.target.checked } })}
            />
            I have reviewed the LHRC Position Reimbursement & Status Report and I certify that all data except the employee / position information noted on the attached Reconciliation Summary Report.
          </label>
        </div>
  
        <p>
          By signing this report, I certify that the information has been reconciled between the Payroll system and Local HR Connect (LHRC). All reconciling differences have been identified and are reflected on the attached Reconciliation Summary Report. Upon request, explanations and supporting documentation for reconciling items are available for review.
        </p>
  
        <hr />
  
        {/* Administrator Signature Section */}
        <h4>Completed by LDSS Office Manager or LHRC Administrator</h4>
        <div className={styles.formGroup}>
          <label>Print Name:</label>
          <input type="text" className={styles.formInput} value={adminFormData.printName}
            onChange={(e) => this.setState({ adminFormData: { ...adminFormData, printName: e.target.value } })} 
          />
        </div>
        {/* <button disabled={true} className={styles.submitButton}>Submit</button> */}
      </div>
    );
  };
    
  private renderAdminForm = (): JSX.Element => {
    const { isSaving, saveStatus } = this.state;
    return (
      <div className={styles.formContainer}>        
        <div className={styles.formTitle}>Local HR Connect (LHRC) Certification Report - Admin</div>
        {this.renderForm()} 

        <button className={styles.saveButton} onClick={this.saveAdminFormToSharePoint} disabled={isSaving}>
          {isSaving ? "Saving..." : "Save to SharePoint"}
        </button>  
        
        {saveStatus && <p className={styles.statusMessage}>{saveStatus}</p>}
      </div>
    );
  };
  
  private renderDirectorForm = (): JSX.Element => {
    return (
      <div className={styles.formContainer}>        
        <div className={styles.formTitle}>Local HR Connect (LHRC) Certification Report - Director</div>
        {this.renderForm()} {/* Reuse the same form */}

        <hr />

        {/* Director Signature Section */}
        <h4>Reviewed by LDSS Director or Assistant Director</h4>
        <div className={styles.formGroup}>
          <label>Print Name: </label>
          <input type="text" className={styles.formInput} placeholder="Enter Name" />
        </div>
      </div>
    );
  };

  private saveAdminFormToSharePoint = async (): Promise<void> => {
    const { adminFormData, selectedAgency } = this.state;
    if (!adminFormData.certifyAccurate || !adminFormData.fips || !adminFormData.certifyException || !adminFormData.printName) {
      this.setState({ saveStatus: "Please fill all fields before saving." });
      return;
    }
  
    // Get current month and year
    const currentDate = new Date();
    const currentMonth = String(currentDate.getMonth() + 1).padStart(2, "0"); // "01" to "12"
    const currentYear = String(currentDate.getFullYear());
  
    const listName = "CertificationReportData"; // Replace with your SharePoint list name
    this.setState({ isSaving: true, saveStatus: "Saving to SharePoint..." });
  
    try {
      await sp.web.lists.getByTitle(listName).items.add({
        Title: selectedAgency,
        field_2: currentMonth,
        field_3: currentYear,
        field_1: adminFormData.fips,
        CertifyAccurate: adminFormData.certifyAccurate,
        CertifyException: adminFormData.certifyException,
        //field_7: adminFormData.certifiedBy,
        field_7: currentDate, // Adding timestamp
        field_8: true, // Assuming checkbox values are boolean
        AdminPrintName: adminFormData.printName,
      });
  
      this.setState({ isSaving: false, saveStatus: "Admin form saved successfully!" });
    } catch (error) {
      console.error("Error saving admin form:", error);
      this.setState({ isSaving: false, saveStatus: "Error saving. Please try again." });
    }
  };

  private renderDiscrepancyReport = (): JSX.Element => {
    const { descrepencyReport, isSaving, saveStatus } = this.state;

    if (this.state.descrepencyReport.length === 0) {
      return <p>No discrepancies found.</p>;
    }

    return (
      <table className={styles.reportTable}>
        <thead>
          <tr>
            <th className={styles.tableValue}>Discrepancy Name</th>
            <th>Count</th>
          </tr>
        </thead>
        <tbody>
          {this.state.descrepencyReport.map((report, index) => (
            <React.Fragment key={index}>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("LetsPositions");
                    }}
                  >
                    LETS positions (filled and vacant)
                  </a>
                </td>
                <td>{report.LetsPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("VacantLetsPositions");
                    }}
                  >
                    Vacant LETS positions
                  </a>
                </td>
                <td>{report.VacantLetsPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("FilledLetsPositions");
                    }}
                  >
                    Filled LETS positions
                  </a>
                </td>
                <td>{report.FilledLetsPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("EmployeeLetsNotFoundLocal");
                    }}
                  >
                    Employees in LETS not found local file
                  </a>
                </td>
                <td>{report.EmployeeLetsNotFoundLocal}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("VacantPositionsLets");
                    }}
                  >
                    Vacant positions in LETS that may be improperly vacant (i.e.
                    there is an equivalent filled position in local data)
                  </a>
                </td>
                <td>{report.VacantPositionsLets}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("NumberofLocalPositions");
                    }}
                  >
                    # of local positions (filled and vacant)
                  </a>
                </td>
                <td>{report.NumberofLocalPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfVacantLocalPositions"
                      );
                    }}
                  >
                    # of filled local positions
                  </a>
                </td>
                <td>{report.NumberOfVacantLocalPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfFilledLocalPositions"
                      );
                    }}
                  >
                    # of employees in local not found in LETS data
                  </a>
                </td>
                <td>{report.NumberOfFilledLocalPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfEmployeeWithSignificantSalary"
                      );
                    }}
                  >
                    # of employees with significant (&gt; $1.00) salary
                    variances between LETS and local data
                  </a>
                </td>
                <td>{report.NumberOfEmployeeWithSignificantSalary}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfLocalPositionsInLETS"
                      );
                    }}
                  >
                    # of local positions that are also in LETS with different
                    state titles
                  </a>
                </td>
                <td>{report.NumberOfLocalPositionsInLETS}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick("LetsLocalPositionBlank");
                    }}
                  >
                    LETS local position is blank
                  </a>
                </td>
                <td>{report.LetsLocalPositionBlank}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfEmployeeWithPastDueProbation"
                      );
                    }}
                  >
                    # of Employees with Past Due Probation Ending Date
                  </a>
                </td>
                <td>{report.NumberOfEmployeeWithPastDueProbation}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfEmployeeWithPastDueAnnual"
                      );
                    }}
                  >
                    # of Employees with Past Due Annual Evaluation Date
                  </a>
                </td>
                <td>{report.NumberOfEmployeeWithPastDueAnnual}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfEmployeeInExpiredPositions"
                      );
                    }}
                  >
                    # of Employees in Expired Positions
                  </a>
                </td>
                <td>{report.NumberOfEmployeeInExpiredPositions}</td>
              </tr>
              <tr>
                <td>
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.handleDiscrepancyClick(
                        "NumberOfPositionsWithInvalidRSC"
                      );
                    }}
                  >
                    # of Positions with Invalid RSC values
                  </a>
                </td>
                <td>{report.NumberOfPositionsWithInvalidRSC}</td>
              </tr>
            </React.Fragment>
          ))}
        </tbody>
        {/* Save to SharePoint Button */}
        <button className={styles.saveButton} onClick={() => this.saveDiscrepancyReportToSharePoint(descrepencyReport)}
          disabled={isSaving} // Disable while saving
        >
          {isSaving ? "Saving..." : "Save to SharePoint"}
        </button>

        {/* Show status message after saving */}
        {saveStatus && <p className={styles.statusMessage}>{saveStatus}</p>}
      </table>
    );
  };

  private handleDiscrepancyClick = (discrepancyName: string): void => {
    const { masterData, validAgencyData } = this.state; // Get master data from state
    let filteredData: IExcelRow[] = [];
    
    switch (discrepancyName) {      
      case "LetsPositions":
        filteredData = masterData; //.filter(row => row.EmployeeStatus === "Active"); // Example filter
        break;
  
      case "VacantLetsPositions":
        filteredData = masterData.filter(row => row.EmployeeFirstName); // Filter vacant positions
        break;

      case "FilledLetsPositions":
        filteredData = masterData.filter(row => !row.EmployeeFirstName); // Filter vacant positions
        break;
  
      case "EmployeeLetsNotFoundLocal":
        filteredData = masterData; //.filter(row => row.EmployeeStatus === "Not Found");
        break;

      case "VacantPositionsLets":
        filteredData = masterData; //.filter(row => row.EmployeeStatus === "Not Found");
        break;
      
      case "NumberofLocalPositions":
          filteredData = validAgencyData; //.filter(row => row.EmployeeStatus === "Not Found");
          break;
      
      case "NumberOfVacantLocalPositions":
        filteredData = validAgencyData.filter((agency) => agency.EmployeeFirstName);
        break;
      
      case "NumberOfFilledLocalPositions":
        filteredData = validAgencyData.filter((agency) => !agency.EmployeeFirstName);
        break;
      
      default:
        filteredData = [];
    }
    
    this.setState({
      selectedDiscrepancy: discrepancyName,
      filteredDiscrepancyData: filteredData,
      activeTab: "DiscrepancyDetails",
    });
  };  

  private handleShowDetailsClick = (item: IExcelRow): void => {
    //console.log("Button clicked for row:", item.BureauFIPS);
    this.setState({ selectedRow: item, isPopupVisible: true });
  };

  private columns: IColumn[] = [
    {
      key: "column1",
      name: "FIPS",
      fieldName: "BureauFIPS",
      minWidth: 50,
      maxWidth: 60,
      isResizable: true,
    },
    {
      key: "column2",
      name: "Region",
      fieldName: "Region",
      minWidth: 80,
      maxWidth: 120,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Per. Num.",
      fieldName: "PersonNumber",
      minWidth: 50,
      maxWidth: 120,
      isResizable: true,
    },
    {
      key: "column4",
      name: "First Name",
      fieldName: "FirstName",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: "column5",
      name: "Last Name",
      fieldName: "LastName",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: "actions",
      name: "Actions",
      minWidth: 100,
      maxWidth: 150,
      isResizable: false,
      onRender: (item: IExcelRow) => (
        <a
          href="#"
          className={styles.inlineLink}
          onClick={(e) => {
            e.preventDefault(); // Prevent default link behavior
            this.handleShowDetailsClick(item);
          }}
        >
          show details
        </a>
      ),
    },
  ];

  private closePopup = (): void => {
    this.setState({ selectedRow: undefined });
  };
  
  private readExcel = (file: File): Promise<IExcelRow[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const binaryStr = e.target?.result as string;
          const workbook = read(binaryStr, { type: "binary" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const data: IExcelRow[] = utils.sheet_to_json(worksheet);
          resolve(data);
        } catch (error) {
          reject(`Failed to parse Excel file. ${error}`);
        }
      };
      reader.onerror = () => reject("Error reading the file.");
      reader.readAsBinaryString(file);
    });
  };

  private renderPopup = (): JSX.Element | null => {
    const { isPopupVisible, selectedRow } = this.state;

    if (isPopupVisible && selectedRow) {
      const entries = Object.entries(selectedRow);

      // Group data into chunks of 2
      const groupedEntries = [];
      for (let i = 0; i < entries.length; i += 2) {
        groupedEntries.push(entries.slice(i, i + 2));
      }

      return (
        <>
          <div className={styles.popupOverlay} />
          <div className={styles.fullPopup}>
            <button className={styles.closeButton} onClick={this.closePopup}>
              Close
            </button>
            <h3>Row Details</h3>
            <table>
              <thead>
                <tr>
                  <th>Column</th>
                  <th>Value</th>
                  <th>Column</th>
                  <th>Value</th>
                </tr>
              </thead>
              <tbody>
                {groupedEntries.map((group, rowIndex) => (
                  <tr key={rowIndex}>
                    {group.map(([key, value], colIndex) => (
                      <React.Fragment key={colIndex}>
                        <td>
                          <strong>{key}</strong>
                        </td>
                        <td>{value}</td>
                      </React.Fragment>
                    ))}
                    {Array.from({ length: 2 - group.length }).map(
                      (_, index) => (
                        <React.Fragment key={`empty-${index}`}>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                        </React.Fragment>
                      )
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      );
    }    
    return null;
  };

  private validateExcelData(data: IExcelRow[]): { validRows: IExcelRow[]; invalidRows: IExcelRow[] } {    
    const validRows: IExcelRow[] = [];
    const invalidRows: IExcelRow[] = [];

    data.forEach((row) => {
      if (row.PayrollPositionNumber && row.JobTitle && row.Salary) {
        validRows.push(row);
      } else {
        invalidRows.push(row);
      }
    });
    this.setState({ validAgencyData: validRows });
    return { validRows, invalidRows };
  }

  private uploadFileToLibrary = async (file: File): Promise<void> => {
    const libraryRelativePath = "/devlab/DocQ"; // Replace with your library path
    const folder = sp.web.getFolderByServerRelativeUrl(libraryRelativePath);

    try {
      // Upload the file
      const uploadResponse = await folder.files.add(file.name, file, true);

      // Set metadata for the uploaded file
      await uploadResponse.file.getItem().then(async (item) => {
        await item.update({
          Agency: this.state.selectedAgency || "Unknown",
        });
      });

      console.log(
        `File "${file.name}" uploaded successfully to "${libraryRelativePath}" with agency metadata.`
      );
    } catch (error) {
      if (error?.message?.includes("404")) {
        throw new Error(
          `Library not found at path: "${libraryRelativePath}". Please verify the library name and location.`
        );
      } else {
        throw error;
      }
    }
  };

  private saveDataToList = async (data: IExcelRow[]): Promise<void> => {
    const listName = "PRS_User_Data"; // Replace with your list name
    const list = sp.web.lists.getByTitle(listName);

    try {
      for (const item of data) {
        await list.items.add({
          Agency: this.state.selectedAgency || "",
          Title: String(item.BureauFIPS) || "",
          field_1: String(item.PayrollPositionNumber) || "", //PayrollPositionNumber
          field_2: item.JobTitle || "", //JobTitle
          field_3: item.StateJobTitle || "", //StateJobTitle
          field_4: item.EmployeeLastName || "", //EmployeeLastName
          field_5: item.EmployeeFirstName || "", //EmployeeFirstName
          field_6: item.EmployeeMiddleInitial || "", //EmployeeMiddleInitial
          field_7: String(item.Salary) || "", //Salary
          field_8: String(item.FTE) || "", //FTE
          field_9: String(item.ReimbursementPercentage) || "", //ReimbursementPercentage
        });
      }
      console.log(
        `Successfully added ${data.length} items to the "${listName}" list.`
      );
    } catch (error) {
      throw new Error(`Failed to save data to list "${listName}". ${error}`);
    }
  };

  public render(): React.ReactElement<ISpDescrepencyProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section
        className={`${styles.spDescrepency} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <header className={styles.header}>
          <h2 className={styles.headerTitle}>
            Discrepancy Management Dashboard
          </h2>
          <p className={styles.headerSubtitle}>
            Identify and manage discrepancies efficiently
          </p>
        </header>

        <main className={styles.mainContent}>
          <div className={styles.uploadSection}>            
            {/* Conditionally Render Agency Dropdown */}
            {this.renderAgencyDropdown()}

            <input
              type="file"
              accept=".xlsx"
              disabled={this.state.masterData.length < 1}
              className={styles.fileInput}
              onChange={this.handleFileUpload}
            />
            <button
              className={styles.uploadButton}
              disabled={
                !this.state.selectedFile ||
                this.state.masterData.length < 1 ||
                this.state.isLoading
              }
              onClick={this.handleValidateDescrepencyClick}
            >
              Show Report
            </button>
          </div>

          <div className={styles.tabContainer}>
            <button
              className={
                this.state.activeTab === "MasterData"
                  ? styles.activeTab
                  : styles.tab
              }
              onClick={() => this.setState({ activeTab: "MasterData" })}
            >
              Master Data
            </button>
            <button
              className={
                this.state.activeTab === "DiscrepancyReport"
                  ? styles.activeTab
                  : styles.tab
              }
              //disabled={!this.state.selectedDiscrepancy} // Disable if no discrepancy selected
              onClick={() => this.setState({ activeTab: "DiscrepancyReport" })}
            >
              Discrepancy Report
            </button>
            <button
              className={
                this.state.activeTab === "DiscrepancyDetails"
                  ? styles.activeTab
                  : styles.tab
              }
              disabled={!this.state.selectedDiscrepancy} // Disable if no discrepancy selected
              onClick={() => this.setState({ activeTab: "DiscrepancyDetails" })}
            >
              Discrepancy Details
            </button>
            {/* Show Admin tab if user is an Admin or HR */}
            {(this.state.isAdmin || this.state.isHR) && (
              <button className={this.state.activeTab === "Admin" ? styles.activeTab : styles.tab }
                onClick={() => this.setState({ activeTab: "Admin" })}
              >
                Admin Tab
              </button>
            )}
            {/* Show Director Tab only if user is authorized */}
            {this.state.isDirector || this.state.isHR && (
              <button
                className={
                  this.state.activeTab === "Director"
                    ? styles.activeTab
                    : styles.tab
                }
                onClick={() => this.setState({ activeTab: "Director" })}
              >
                Director Tab
              </button>
            )}
          </div>

          <div className={styles.tabContent}>
            {this.state.isLoading ? (
              <Spinner
                size={SpinnerSize.large}
                label="Please wait while loading data..."
              />
            ) : (
              <>
                {this.state.activeTab === "MasterData" && this.renderMasterDataGrid()}
                {this.state.activeTab === "DiscrepancyReport" && this.renderDiscrepancyReport()}
                {this.state.activeTab === "DiscrepancyDetails" && this.renderSelectedDiscrepancyDetails()}
                {/* Only render Admin Tab if user is Admin or HR */}
                {(this.state.isAdmin || this.state.isHR) && this.state.activeTab === "Admin" && this.renderAdminForm()}
                {/* Only render Director Tab if user is Director or HR */}
                {(this.state.isDirector || this.state.isHR) && this.state.activeTab === "Director" && this.renderDirectorForm()}
              </>
            )}
          </div>

          {/* Popup for Row Details */}
          {this.renderPopup()}

          {/* 
          {this.state.isLoading && (
            <Spinner size={SpinnerSize.medium} label="Processing..." />
          )}
          
          {this.state.errorMessage && (
            <MessageBar messageBarType={MessageBarType.error}>
              {this.state.errorMessage}
            </MessageBar>
          )}
          
          <p className={this.state.style}>{this.state.uploadStatus}</p>
          */}
        </main>

        <footer className={styles.footer}>
          <p>&copy; {new Date().getFullYear()} Discrepancy Management System</p>
        </footer>
      </section>
    );
  }
}