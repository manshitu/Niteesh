import * as React from "react";
import styles from "./SpDescrepency.module.scss";
import type { ISpDescrepencyProps } from "./ISpDescrepencyProps";
import { sp } from "@pnp/sp/presets/all";
import { read, utils } from "xlsx";
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
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
  currentPage: number;
  isPopupVisible: boolean;
  selectedRow: IExcelRow | undefined;
  descrepencyReport: IDiscrepancyResult[];
  showDescrepencyPopup: boolean;
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
      currentPage: 1,
      isPopupVisible: false,
      selectedRow: undefined,
      descrepencyReport: [],
      showDescrepencyPopup: false,
    };

    sp.setup({
      spfxContext: this.props.context as IWebPartContext,
    });
  }

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
  
    return {
      LetsPositions: letsPositions,
      VacantLetsPositions: vacantLetsPositions,
      FilledLetsPositions: filledLetsPositions,
    };
  };

  private displayDiscrepancyReport = (discrepancies: IDiscrepancyResult[]): void => {
    if (!discrepancies) {
      alert("No discrepancies found. Data matches the master database.");
      return;
    }

    // Navigate to another screen or display a modal
    this.setState({
      descrepencyReport: discrepancies,
      showDescrepencyPopup: true, // Example for showing a popup
    });
  };

  private handleAgencyChange = async (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> => {
    const selectedAgency = option?.key as string;

    this.setState({ selectedAgency, isLoading: true, masterData: [] });

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
        FIPS: item.field_4
      }));
    } catch (error) {
      console.error("Error fetching list data: ", error);
      throw new Error("Failed to fetch data from the list.");
    }
  };

  public renderMasterDataGrid(): JSX.Element {
    const { masterData, isLoading } = this.state;

    if (isLoading) {
      return <Spinner size={SpinnerSize.medium} label="Loading data..." />;
    }

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

  private closeDescrepencyPopup = (): void => {
    this.setState({
      descrepencyReport: [],
      showDescrepencyPopup: false, // Example for showing a popup
    });
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
    const { isPopupVisible, selectedRow, showDescrepencyPopup, descrepencyReport } = this.state;
  
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
                        <td><strong>{key}</strong></td>
                        <td>{value}</td>
                      </React.Fragment>
                    ))}
                    {/* Fill empty cells for incomplete rows */}
                    {Array.from({ length: 2 - group.length }).map((_, index) => (
                      <React.Fragment key={`empty-${index}`}>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </React.Fragment>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      );
    }
  
    if (showDescrepencyPopup) {
      return (
        <>
        <div className={styles.popupOverlay} />
        <div className={styles.popup}>
            <button className={styles.closeButton} onClick={this.closeDescrepencyPopup}>
              Close
            </button>
            <h3>Discrepancy Report</h3>            
            <table>
              <thead>
                <tr>
                  <th className={styles.tableValue}> Descrepency Name </th>                  
                  <th> Count </th>
                </tr>
              </thead>
              <tbody>
                {descrepencyReport.map((report, index) => (
                  <>
                    <tr>
                      <td><a href="#">LETS positions (filled and vacant)</a></td>
                      <td>{report.LetsPositions}</td>
                    </tr>
                    <tr>
                      <td><a href="#">Vacant LETS positions</a></td>
                      <td>{report.VacantLetsPositions}</td>
                    </tr>
                    <tr>
                      <td><a href="#">Filled LETS positions</a></td>
                      <td>{report.FilledLetsPositions}</td>
                    </tr>                    
                  </>
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
    return { validRows, invalidRows };
  }

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
                !this.state.selectedAgency ||
                this.state.isLoading
              }
              onClick={this.handleValidateDescrepencyClick} // Same function as the input's onChange
            >              
              Show Report
            </button>
          </div>

          <div>           
            {this.renderMasterDataGrid()}
            {this.renderPopup()}
          </div>

          {this.state.isLoading && (
            <Spinner size={SpinnerSize.medium} label="Processing..." />
          )}

          {this.state.errorMessage && (
            <MessageBar messageBarType={MessageBarType.error}>
              {this.state.errorMessage}
            </MessageBar>
          )}
          
          <p className={this.state.style}>{this.state.uploadStatus}</p>
        </main>

        <footer className={styles.footer}>
          <p>&copy; {new Date().getFullYear()} Discrepancy Management System</p>
        </footer>
      </section>
    );
  }
}