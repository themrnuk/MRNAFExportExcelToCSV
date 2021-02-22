using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class INNRecruitmentSheet : ExcelFile
    {
        string folderName = "INNUS Recruitment";
        public INNRecruitmentSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {

        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {

            switch (excelSheetName)
            {
                case "Hired":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "INN_Recruitment_Hired.csv", TableName = "INN_Recruitment_Hired", FolderName = folderName };
                    }
                case "Open Locations":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "INN_Recruitment_OL.csv", TableName = "INN_Recruitment_OL", FolderName = folderName };

                    }
                case "Onboarding Tracking":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "INN_Recruitment_OBT.csv", TableName = "INN_Recruitment_OBT", HeaderRow = 4, FolderName = folderName };
                    }
                case "Applicant Tracking":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "INN_Recruitment_AT.csv", TableName = "INN_Recruitment_AT", FolderName = folderName };

                    }
                case "Not Onbaorded":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "INN_Recruitment_NOB.csv", TableName = "INN_Recruitment_NOB", FolderName = folderName };

                    }
                case "Historic":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = "INN_Recruitment_Historic.csv", TableName = "INN_Recruitment_Historic", FixedColumn = true, ColumnCount = 7, FolderName = folderName };

                    }
                default:
                    {
                        return null;
                    }

            }

        }
    }
}
