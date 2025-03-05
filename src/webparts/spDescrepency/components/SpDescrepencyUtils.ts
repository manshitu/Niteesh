import { read, utils } from "xlsx";
import { IExcelRow } from "./ISpDescrepencyProps";
import { IColumn } from "@fluentui/react";

export const readExcel = (file: File): Promise<IExcelRow[]> => {
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

export const getDescrepencyColumns = (descrepencyName: string): Promise<IColumn[]> => {
    let columns: IColumn[] = [];
    if (descrepencyName === "LetsPositions") {
        columns = [
          {
            key: "1",
            name: "Employee Name",
            fieldName: "FirstName",
            minWidth: 100,
            maxWidth: 150,
          },
          {
            key: "2",
            name: "Position Number",
            fieldName: "LocalPositionNumber",
            minWidth: 100,
            maxWidth: 150,
          },
          {
            key: "3",
            name: "Salary",
            fieldName: "EmployeeSalary",
            minWidth: 80,
            maxWidth: 120,
          },
        ];
      } else if (descrepencyName === "VacantLetsPositions") {
        columns = [
          {
            key: "1",
            name: "BureauFIPS",
            fieldName: "BureauFIPS",
            minWidth: 100,
            maxWidth: 150,
          },
          {
            key: "2",
            name: "Job Title",
            fieldName: "Region",
            minWidth: 100,
            maxWidth: 150,
          },
        ];
      } else if (descrepencyName === "FilledLetsPositions") {
        columns = [
          {
            key: "1",
            name: "Employee Name",
            fieldName: "FirstName",
            minWidth: 100,
            maxWidth: 150,
          },
          {
            key: "2",
            name: "Status",
            fieldName: "EmployeeStatus",
            minWidth: 100,
            maxWidth: 150,
          },
        ];
      }
      return Promise.resolve(columns);
  };