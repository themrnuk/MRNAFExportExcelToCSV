using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public abstract class ExcelFile
    {
        public string excelFileName = string.Empty;
        public List<AdditionalColumns> AdditionalColumns;
        protected string[] args;
        public ExcelFile(string _excelFileName, string[] _args)
        {
            excelFileName = _excelFileName;
            args = _args;
            AdditionalColumns = new List<AdditionalColumns>(){new AdditionalColumns() { ColumnName = "SOURCE_SPREAD_SHEET", ColumnValue = _excelFileName },
                                                              new AdditionalColumns() { ColumnName = "TIMEDATE_SNAPSHOT", ColumnValue = DateTime.Now.ToString("dd'/'MM'/'yyyy HH:mm:ss") },
                                                              new AdditionalColumns() { ColumnName = "SOURCE_SHEET", ColumnValue = string.Empty },
                                                              new AdditionalColumns() { ColumnName = "CELL_RANGE", ColumnValue = string.Empty }
        };
        }
        public abstract ExcelSheets GetCsvSheet(string excelSheetName);
    }
}
