using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class PTSheet : ExcelFile
    {
        string folderName = "PT";
        string projectCode = string.Empty;
        string countryCode = string.Empty;
        public PTSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {
            projectCode = args[0];
            countryCode = args[1];
            this.AdditionalColumns.Add(new AdditionalColumns() { ColumnName = "PT_CODE", ColumnValue = projectCode });
        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {

            switch (excelSheetName)
            {
                case "SIV Tracker":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_ST.csv", TableName = "HTS_PT_ST", FolderName = folderName, HeaderRow = 2 };
                    }
                case "Patient List":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_PL.csv", TableName = "HTS_PT_PL", FolderName = folderName, HeaderRow = 2 };
                    }
                case "Nurse List":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_NL.csv", TableName = "HTS_PT_NL", FolderName = folderName, HeaderRow = 2 };
                    }
                case "Visit Tracker":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_VT.csv", TableName = "HTS_PT_VT", FolderName = folderName, HeaderRow = 2 };
                    }
                case "MRN Internal DCF Tracker":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_DCF.csv", TableName = "HTS_PT_DCF", FolderName = folderName, HeaderRow = 2 };
                    }
                case "Study Contacts":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_SC.csv", TableName = "HTS_PT_SC", FolderName = folderName, HeaderRow = 1 };
                    }
                case "CRA List & Training":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"HTS_PT_{projectCode}_{countryCode}_CLT.csv", TableName = "HTS_PT_CLT", FolderName = folderName };
                    }
                default:
                    {
                        return null;
                    }
            }
        }
    }
}
