using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class PILFolderSheet : ExcelFile
    {

        string folderName = "PIL";
        string projectNumber = string.Empty;
        public PILFolderSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {
            projectNumber = args[0];
            this.AdditionalColumns.Add(new AdditionalColumns() { ColumnName = "PROJECT NUMBER", ColumnValue = projectNumber });
        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            switch (excelSheetName)
            {
                case "Issues and Deviations":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"IssueLog_Issues_and_Deviations_{projectNumber}.csv", TableName = "IssueLog_Issues_and_Deviations", FolderName = folderName, HeaderRow = 1 };

                    }
                case "Clinical Incidents":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"IssueLog_Clinical_Incidents_{projectNumber}.csv", TableName = "IssueLog_Clinical_Incident", FolderName = folderName, HeaderRow = 1 };

                    }

                default:
                    {
                        return null;
                    }
            }
        }
    }
}
