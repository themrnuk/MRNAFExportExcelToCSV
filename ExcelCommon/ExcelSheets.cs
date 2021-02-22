using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class ExcelSheets
    {

        public string SheetName { get; set; }
        public string CsvFileName { get; set; }
        public string TableName { get; set; }
        public string FolderName { get; set; }
        public string TableColumns { get; set; }
        public int HeaderRow { get; set; }
        public bool FixedColumn { get; set; }
        public int ColumnCount { get; set; }
        public ExcelFile CsvFile { get; set; }
    }
}
