using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    class ExcelToCSVErrorLog
    {
        public string ExcelFileName { get; set; }
        public string ExcelSheetName { get; set; }
        public string CSVName { get; set; }
        public string ExcelRowNumber { get; set; }
        public string ErrorMessage { get; set; }

        public string ExceptionMessage { get; set; }
        public DateTime DateCreated { get; set; }
    }
}
