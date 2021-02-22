using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class VendorGovernanceTrackerSheet : ExcelFile
    {
        string folderName = "Vendors";

        public VendorGovernanceTrackerSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }

        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            if (excelSheetName == "Govn_Tracker")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "Vendor_Governance.csv", TableName = "Vendor_Governance", FolderName = folderName };
            }
            return null;
        }

    }
}
