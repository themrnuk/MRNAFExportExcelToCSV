using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class NoBSessionsSheet : ExcelFile
    {
        string folderName = "Nursing";
        string region = string.Empty;
        public NoBSessionsSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {
            region = args[0];
            this.AdditionalColumns.Add(new AdditionalColumns() { ColumnName = "REGION", ColumnValue = region });
        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            switch (excelSheetName)
            {
                case "Sessions":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"Nurse_Sessions_{region}.csv", TableName = "Nurse_Sessions", FolderName = folderName };
                    }
                case "Training":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"Nurse_Training_{region}.csv", TableName = "Nurse_Training", FolderName = folderName };
                    }
                case "Nurse Docs":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"Nurse_Docs_{region}.csv", TableName = "Nurse_Docs", FolderName = folderName };
                    }
                default:
                    {
                        return null;
                    }
            }



        }
    }
}
