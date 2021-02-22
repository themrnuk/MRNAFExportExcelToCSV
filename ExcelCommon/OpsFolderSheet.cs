using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class OpsFolderSheet : ExcelFile
    {
        string folderName = "OPS";
        string sessionid = string.Empty;
        public OpsFolderSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {
             sessionid = args[0];
            this.AdditionalColumns.Add(new AdditionalColumns() { ColumnName = "SESSION_ID", ColumnValue = sessionid });
        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {      
            if (excelSheetName == "Unit")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"Ops_Unit_{sessionid}.csv", TableName = "Ops_Unit", FolderName = folderName, HeaderRow = 2 };
            }
            return null;

        }
    }
}
