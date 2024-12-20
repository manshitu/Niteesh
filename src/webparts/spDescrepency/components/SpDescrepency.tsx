import * as React from "react";
import styles from "./SpDescrepency.module.scss";
import type { ISpDescrepencyProps } from "./ISpDescrepencyProps";
import { sp } from "@pnp/sp/presets/all";
import { read, utils } from "xlsx";

interface ISpDescrepencyState {
  uploadStatus: string;
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

export default class SpDescrepency extends React.Component<ISpDescrepencyProps, ISpDescrepencyState> {
  constructor(props: ISpDescrepencyProps) {
    super(props);
    this.state = {
      uploadStatus: "",
    };
  }

  private handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];

    if (!file) {
      this.setState({ uploadStatus: "No file selected. Please choose a valid .xlsx file." });
      return;
    }

    try {
      this.setState({ uploadStatus: "Processing file, please wait..." });

      // Read and parse the Excel file
      const data = await this.readExcel(file);

      // Upload the file to the SharePoint library
      //await this.uploadFileToLibrary(file);

      // Save the extracted data to the SharePoint list
      await this.saveDataToList(data);

      this.setState({ uploadStatus: "File uploaded and data saved successfully!" });
    } catch (error) {
      console.error("Error during file upload:", error);
      this.setState({ uploadStatus: "An error occurred while processing the file. Please try again." });
    }
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
          reject(`Failed to parse Excel file.  "${error}"`);
        }
      };
      reader.onerror = () => reject("Error reading the file.");
      reader.readAsBinaryString(file);
    });
  };

  private uploadFileToLibrary = async (file: File): Promise<void> => {  
    debugger;  
    const libraryName = "DocQ"; // Replace with your library name
    const folder = sp.web.getFolderByServerRelativeUrl(libraryName);
    console.log(folder);
    //var endpoint = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/RootFolder/Files/add(url='${fileName}', overwrite=true)`;

    try {
      await folder.files.add(file.name, file, true);
    } catch (error) {
      throw new Error(`Failed to upload file to library "${libraryName}". "${error}"`);
    }
  };
    
  private saveDataToList = async (data: IExcelRow[]): Promise<void> => {
    debugger;
    const listName = "Discrepancy"; // Replace with your list name
    const list = sp.web.lists.getByTitle(listName);

    try {
      for (const item of data) {
        await list.items.add({
          Title: item.BureauFIPS || "",
          PayrollPositionNumber: item.PayrollPositionNumber || "",
          JobTitle: item.JobTitle || "",          
          StateJobTitle: item.StateJobTitle || "",
          EmployeeLastName: item.EmployeeLastName || "",
          EmployeeFirstName: item.EmployeeFirstName || "",
          EmployeeMiddleInitial: item.EmployeeMiddleInitial || "",
          FTE: item.FTE || "",
          ReimbursementPercentage: item.ReimbursementPercentage || "",
        });
      }
    } catch (error) {
      throw new Error(`Failed to save data to list "${listName}". "${error}"`);
    }
  };

  public render(): React.ReactElement<ISpDescrepencyProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section
        className={`${styles.spDescrepency} ${hasTeamsContext ? styles.teams : ""}`}
      >
        <header className={styles.header}>
          <h1 className={styles.headerTitle}>Discrepancy Management Dashboard</h1>
          <p className={styles.headerSubtitle}>Identify and manage discrepancies efficiently</p>
        </header>

        <main className={styles.mainContent}>
          <div className={styles.uploadSection}>
            <h4>Upload and Analyze Data</h4>
            <input
              type="file"
              accept=".xlsx"
              className={styles.fileInput}
              onChange={this.handleFileUpload}
            />
            <br />
            <p>{this.state.uploadStatus}</p>
          </div>
        </main>

        <footer className={styles.footer}>
          <p>&copy; {new Date().getFullYear()} Discrepancy Management System</p>
        </footer>
      </section>
    );
  }
}