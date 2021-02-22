using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class INNFolderSheet : ExcelFile
    {
        string folderName = "INN";
        string countryCode = string.Empty;
        public INNFolderSheet(string excelFileName, string[] args) : base(excelFileName, args)
        {
            countryCode = args[0];
            this.AdditionalColumns.Add(new AdditionalColumns() { ColumnName = "COUNTRY CODE", ColumnValue = countryCode });
        }
        public override ExcelSheets GetCsvSheet(string excelSheetName)
        {
            switch (excelSheetName)
            {
                case "Referral Tracker":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_RT1.csv", TableName = "INN_PT_RT1", FolderName = folderName };

                    }
                case "Patient Nurse List":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_PNL.csv", TableName = "INN_PT_PNL", FolderName = folderName };

                    }
                case "Visit Scheduler":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_VS1.csv", TableName = "INN_PT_VS1", FolderName = folderName };

                    }
                case "Nurse Database":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_ND1.csv", TableName = "INN_PT_ND1", FolderName = folderName };

                    }
                case "All Projects Information":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_API.csv", TableName = "INN_PT_API", FolderName = folderName };

                    }
                case "SNS Referral Tracker":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_SRT.csv", TableName = "INN_PT_SRT", FolderName = folderName };

                    }
                case "Site Nurse List":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_SNL.csv", TableName = "INN_PT_SNL", FolderName = folderName };

                    }
                case "SNS Visit Scheduler":
                    {
                        return new ExcelSheets() { SheetName = excelSheetName, CsvFileName = $"INN_PT_{countryCode}_SVS.csv", TableName = "INN_PT_SVS", FolderName = folderName };
                    }
                default:
                    {
                        return null;
                    }
            }



        }
    }
}
