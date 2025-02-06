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
  isDirector: boolean; // New state to control access
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
      isDirector: false, // Default to false until we check permissions
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
    await this.checkDirectorAccess(); // Check user access on load
  }

  private checkDirectorAccess = async (): Promise<void> => {
    try {
      //const currentUser = await sp.web.currentUser.get();
      //const userGroups = await sp.web.currentUser.groups.get();
  
      // Check if user belongs to "Directors" group
      //const isDirector = userGroups.some(group => group.Title === "Directors");
  
      this.setState({ isDirector: true });
    } catch (error) {
      console.error("Error checking user permissions:", error);
    }
  };
  
  private handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
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

  private handleValidateDescrepencyClick = async () => {
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

  private calculateDiscrepancies = (
    validRows: IExcelRow[],
    masterData: IExcelRow[]
  ): IDiscrepancyResult => {
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

  private handleAgencyChange = async (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> => {
    const selectedAgency = option?.key as string;

    this.setState({ selectedAgency, isLoading: true, activeTab: "MasterData", masterData: [] });
    
    if (selectedAgency) {
      try {
        const data = await this.fetchMasterAgencyData(selectedAgency);
        this.setState({ masterData: data, errorMessage: "" });
      } catch (error) {
        this.setState({
          errorMessage:
            "Error fetching data for the selected agency. Please try again." +
            error,
        });
      } finally {
        this.setState({ isLoading: false });
      }
    }
  };

  private fetchMasterAgencyData = async (
    agency: string
  ): Promise<IExcelRow[]> => {
    const listName = "PRS_Master_Data"; // Replace with your list name
    try {
      const items = await sp.web.lists
        .getByTitle(listName)
        .items.filter(`field_4 eq '${agency}'`) // Filter based on the selected agency
        .select("*")
        .top(100) // Adjust the number of rows to fetch
        .get();

      return items.map((item) => ({
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
    } catch (error) {
      console.error("Error fetching list data: ", error);
      throw new Error("Failed to fetch data from the list.");
    }
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
  /*
  private renderApprovalForm = (): JSX.Element => {
    return (
      <div className={styles.formContainer}>
        <h3>Discrepancy Resolution Form</h3>
        <input type="text" placeholder="Field 1" className={styles.formInput} />
        <input type="text" placeholder="Field 2" className={styles.formInput} />
        <input type="text" placeholder="Field 3" className={styles.formInput} />
        <input type="text" placeholder="Field 4" className={styles.formInput} />
  
        <select className={styles.formDropdown}>
          <option value="">Select Option 1</option>
          <option value="Option A">Option A</option>
          <option value="Option B">Option B</option>
        </select>
  
        <select className={styles.formDropdown}>
          <option value="">Select Option 2</option>
          <option value="Option X">Option X</option>
          <option value="Option Y">Option Y</option>
        </select>
  
        <button className={styles.submitButton}>Submit</button>
      </div>
    );
  };
  */

  private renderForm = (): JSX.Element => {
    return (      
      <div>
        {/* Locality Information */}
        <div className={styles.formGroup}>
          <label>Locality Name (City or County):</label>
          <input type="text" className={styles.formInput} placeholder="Enter Locality Name" />
          <label>FIPS:</label>
          <input type="text" className={styles.formInputSmall} placeholder="Enter FIPS" />
          <label>Month:</label>
          <input type="text" className={styles.formInputSmall} placeholder="Enter Month" />
        </div>
    
        <div className={styles.formGroup}>
          <label>Cycle being Certified (Month/Year) (example: 05/2023):</label>
          <input type="text" className={styles.formInput} placeholder="MM/YYYY" />
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
            <input type="checkbox" className={styles.checkbox} /> I have reviewed the LHRC Position Reimbursement & Status Report and I certify that it is accurate.
          </label>
        </div>
  
        <div className={styles.checkboxGroup}>
          <label>
            <input type="checkbox" className={styles.checkbox} /> I have reviewed the LHRC Position Reimbursement & Status Report and I certify that all data except the employee / position information noted on the attached Reconciliation Summary Report.
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
          <input type="text" className={styles.formInput} placeholder="Enter Name" />
        </div>
        {/* <button disabled={true} className={styles.submitButton}>Submit</button> */}
      </div>
    );
  };
  
  private renderAdminForm = (): JSX.Element => {
    return (
      <div className={styles.formContainer}>        
        <div className={styles.formTitle}>Local HR Connect (LHRC) Certification Report - Admin</div>
        {this.renderForm()} {/* Reuse the form */}
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

  private renderDiscrepancyReport = (): JSX.Element => {
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
                  <a href="#" onClick={(e) => {e.preventDefault(); this.handleDiscrepancyClick("LetsPositions"); }}>
                    LETS positions (filled and vacant)
                  </a>
                </td>
                <td>{report.LetsPositions}</td>
              </tr>
              <tr>
                <td>                  
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("VacantLetsPositions"); }}>  
                    Vacant LETS positions
                  </a>
                </td>
                <td>{report.VacantLetsPositions}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("FilledLetsPositions"); }}>
                    Filled LETS positions
                  </a>
                </td>
                <td>{report.FilledLetsPositions}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("EmployeeLetsNotFoundLocal"); }}>
                    Employees in LETS not found local file
                  </a>
                </td>
                <td>{report.EmployeeLetsNotFoundLocal}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("VacantPositionsLets"); }}>
                    Vacant positions in LETS that may be improperly vacant (i.e.
                    there is an equivalent filled position in local data)
                  </a>
                </td>
                <td>{report.VacantPositionsLets}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberofLocalPositions"); }}>
                    # of local positions (filled and vacant)
                  </a>
                </td>
                <td>{report.NumberofLocalPositions}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfVacantLocalPositions"); }}>
                    # of filled local positions
                  </a>
                </td>
                <td>{report.NumberOfVacantLocalPositions}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfFilledLocalPositions"); }}>
                    # of employees in local not found in LETS data
                  </a>
                </td>
                <td>{report.NumberOfFilledLocalPositions}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfEmployeeWithSignificantSalary"); }}>
                    # of employees with significant (&gt; $1.00) salary
                    variances between LETS and local data
                  </a>
                </td>
                <td>{report.NumberOfEmployeeWithSignificantSalary}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfLocalPositionsInLETS"); }}>
                    # of local positions that are also in LETS with different
                    state titles
                  </a>
                </td>
                <td>{report.NumberOfLocalPositionsInLETS}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("LetsLocalPositionBlank"); }}>
                    LETS local position is blank
                  </a>
                </td>
                <td>{report.LetsLocalPositionBlank}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfEmployeeWithPastDueProbation"); }}>
                    # of Employees with Past Due Probation Ending Date
                  </a>
                </td>
                <td>{report.NumberOfEmployeeWithPastDueProbation}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfEmployeeWithPastDueAnnual"); }}>
                    # of Employees with Past Due Annual Evaluation Date
                  </a>
                </td>
                <td>{report.NumberOfEmployeeWithPastDueAnnual}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfEmployeeInExpiredPositions"); }}>
                    # of Employees in Expired Positions
                  </a>
                </td>
                <td>{report.NumberOfEmployeeInExpiredPositions}</td>
              </tr>
              <tr>
                <td>
                  <a href="#" onClick={(e) => { e.preventDefault(); this.handleDiscrepancyClick("NumberOfPositionsWithInvalidRSC"); }}>
                    # of Positions with Invalid RSC values
                  </a>
                </td>
                <td>{report.NumberOfPositionsWithInvalidRSC}</td>
              </tr>
            </React.Fragment>
          ))}
        </tbody>
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

  private validateExcelData(data: IExcelRow[]) {
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
            <Dropdown
              label="Select Agency Name"
              title="Select an agency"
              placeholder="Select an agency"
              options={this.agencyOptions}
              onChange={this.handleAgencyChange}
              selectedKey={this.state.selectedAgency}
              className={styles.dropdown}
            />
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
            <button className={
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
            <button className={this.state.activeTab === "DiscrepancyDetails" ? styles.activeTab : styles.tab}
              disabled={!this.state.selectedDiscrepancy} // Disable if no discrepancy selected
              onClick={() => this.setState({ activeTab: "DiscrepancyDetails" })}>
              Discrepancy Details
            </button>
            <button className={this.state.activeTab === "Admin" ? styles.activeTab : styles.tab}
              onClick={() => this.setState({ activeTab: "Admin" })}>
              Admin Tab
            </button>
            {/* Show Director Tab only if user is authorized */}
            {this.state.isDirector && (
              <button className={this.state.activeTab === "Director" ? styles.activeTab : styles.tab}
                onClick={() => this.setState({ activeTab: "Director" })}>
                Director Tab
              </button>
            )}
          </div>

          <div className={styles.tabContent}>
            {this.state.isLoading ? (
              <Spinner size={SpinnerSize.large} label="Please wait while loading data..." />
            ) : (
              <>
                {this.state.activeTab === "MasterData" && this.renderMasterDataGrid()}
                {this.state.activeTab === "DiscrepancyReport" && this.renderDiscrepancyReport()}
                {this.state.activeTab === "DiscrepancyDetails" && this.renderSelectedDiscrepancyDetails()}
                {this.state.activeTab === "Admin" && this.renderAdminForm()}  {/* Admin Tab */}
                {/* Only render Director Tab if user is authorized */}
                {this.state.isDirector && this.state.activeTab === "Director" && this.renderDirectorForm()}  {/* Director Tab */}                
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