using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class NurseListSheet : ExcelFile
    {
        string folderName = "Nursing";
        public NurseListSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }

        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            if (excelSheetName == "Nurse List")
            {
                return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "Nurse_List_.csv", TableName = "Nurse_List", FolderName = folderName };
            }
            return null;
        }
    }
}
