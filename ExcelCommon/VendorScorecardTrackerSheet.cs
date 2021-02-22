using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class VendorScorecardTrackerSheet : ExcelFile
    {
        string folderName = "Vendors";

        public VendorScorecardTrackerSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }

        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            if (excelSheetName == "Scorecard_Tracker")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "Vendor_Scorecard.csv", TableName = "Vendor_Scorecard", FolderName = folderName };
            }
            return null;
        }
    }
}
