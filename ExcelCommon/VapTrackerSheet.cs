using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class VapTrackerSheet : ExcelFile
    {
        string folderName = "Nursing";
        public VapTrackerSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }

        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            if (excelSheetName == "VAP Table")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "Nurse_VAP_Table_.csv", TableName = "Nurse_VAP_Table", FolderName = folderName };
            }
            return null;
        }
    }
}
