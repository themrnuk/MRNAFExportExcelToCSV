using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class ReferenceTablesSheet : ExcelFile
    {

        string folderName = "Reference";
        public ReferenceTablesSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            if (excelSheetName == "STATIC_DECODES")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "Reference_Codes.csv", TableName = "Reference_Codes", FolderName = folderName };
            }
            return null;
        }
    }
}
