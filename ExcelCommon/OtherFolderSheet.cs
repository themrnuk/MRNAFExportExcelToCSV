using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class OtherFolderSheet : ExcelFile
    {
        string folderName = "Other";
        public OtherFolderSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            if (excelSheetName == "Sheet3")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"Historic_Projects.csv", TableName = "Historic_Projects", FolderName = folderName, FixedColumn = true, ColumnCount = 2 };
            }
            return null;
        }
    }
}
