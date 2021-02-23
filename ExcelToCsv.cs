using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
//using System.Threading;
using ExcelDataReader;
using ExcelNumberFormat;
//using Microsoft.Azure.Storage;
//using Microsoft.Azure.Storage.Blob;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.WindowsAzure.Storage;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MRNAFExportExcelToCSV
{
    public static class ExcelToCsv
    {
        [FunctionName("ExcelToCsv")]
        public static async void Run([BlobTrigger("%ContainerName%/Excels/{name}", Connection = "AzureWebJobsStorage")] Stream excelFileInput, Binder binder, string name, ILogger log, ExecutionContext context)
        {
            log.LogInformation($"C# Blob trigger function executed at: {DateTime.Now}");

            var config = new ConfigurationBuilder()
           .SetBasePath(context.FunctionAppDirectory)
           .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            string defaultContainerName = config.GetConnectionStringOrSetting("ContainerName");




            //string folderPath = string.Empty;

            string folderName = string.Empty;
            if (name.Contains("/"))
            {
                return;
            }
            string errMsg = string.Empty;
            ExcelFile csvFile = null;
            if (!name.EndsWith(".xls") && !name.EndsWith(".xlsx") && !name.EndsWith(".xlsm"))
            {
                errMsg = $"unable to convert file: {name} into csv.Please upload upload a file with extension (.xls)(.xlsx)(.xlsm)";
            }
            else
            {

                if (name.ToLower().StartsWith("vendor governance tracker"))
                {
                    csvFile = new VendorGovernanceTrackerSheet(name, null);
                }
                else if (name.ToLower().StartsWith("vendor scorecard tracker"))
                {
                    csvFile = new VendorScorecardTrackerSheet(name, null);
                }
                else if (name.ToLower().StartsWith("applicant & onboarding tracker"))
                {
                    csvFile = new INNRecruitmentSheet(name, null);
                }

                else if (name.ToLower().StartsWith("reference_tables"))
                {
                    csvFile = new ReferenceTablesSheet(name, null);
                }
                else if (name.ToLower().StartsWith("nurse list"))
                {
                    csvFile = new NurseListSheet(name, null);
                }
                else if (name.ToLower().StartsWith("vap tracker"))
                {
                    csvFile = new VapTrackerSheet(name, null);
                }
                else if (name.ToLower().StartsWith("nob sessions"))
                {
                    var namearray = name.Split(" ");
                    if (namearray.Length < 3)
                    {
                        log.LogError($"unable to convert file: {name} into csv");
                        return;
                    }
                    csvFile = new NoBSessionsSheet(name, new string[] { namearray[2] });

                }
                else if (name.ToUpper().StartsWith("INN"))
                {
                    var namearray = name.Split(" ");
                    if (namearray.Length < 2)
                    {
                        errMsg = ($"unable to convert file: {name} into csv.Invalid file Name.");
                    }
                    else
                    {
                        string countryCode = namearray[0].Substring(3);
                        csvFile = new INNFolderSheet(name, new string[] { countryCode });
                    }
                }
                else if (name.ToLower().Contains("_project issue log_"))
                {
                    var namearray = name.Split("_");
                    if (namearray.Length < 2)
                    {
                        errMsg = ($"unable to convert file: {name} into csv.Invalid file Name.");
                    }
                    else
                    {
                        string projectNumber = namearray[0];
                        csvFile = new PILFolderSheet(name, new string[] { projectNumber });
                    }
                }
                else if (name.ToLower().StartsWith("old hts opportunities"))
                {
                    csvFile = new OtherFolderSheet(name, null);
                }
                else if (name.ToLower().Contains("ops sheet"))
                {
                    var namearray = name.Split("#");
                    if (namearray.Length < 2)
                    {
                        errMsg = ($"unable to convert file: {name} into csv.Invalid file Name.");
                    }
                    else
                    {
                        string sessionNumber = string.Empty;
                        for (int digit = 0; digit < namearray[1].Length; digit++)
                        {
                            if (Char.IsDigit(namearray[1][digit]))
                            {
                                sessionNumber += namearray[1][digit];
                            }

                            else
                            {
                                break;
                            }
                        }
                        csvFile = new OpsFolderSheet(name, new string[] { sessionNumber });
                    }
                }
                else
                {
                    var namearray = name.Split("_");
                    if (namearray.Length < 3 && !(Array.IndexOf(namearray, "Project Plan Tracker") > 1))
                    {
                        errMsg = ($"unable to convert file: {name} into csv.Invalid file Name.");
                    }
                    else
                    {
                        csvFile = new PTSheet(name, namearray);
                    }
                }

            }

            if (csvFile != null && string.IsNullOrWhiteSpace(errMsg))
            {
                log.LogInformation($"Do your processing on the excelFileInput file here.");
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (csvFile.GetType() == typeof(INNRecruitmentSheet))
                {
                    ProcessINNRecruitmentSheet(csvFile, excelFileInput, binder, name, defaultContainerName, log);
                }
                else
                {
                    using (excelFileInput)
                    {

                        IExcelDataReader reader = null;
                        if (name.EndsWith(".xls"))
                        {
                            reader = ExcelReaderFactory.CreateBinaryReader(excelFileInput);
                        }
                        else if (name.EndsWith(".xlsx") || name.EndsWith(".xlsm"))
                        {
                            try
                            {
                                reader = ExcelReaderFactory.CreateOpenXmlReader(excelFileInput);
                            }
                            catch (Exception ex)
                            {
                                log.LogError(ex.Message);
                                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                obj.ExcelFileName = name;
                                obj.ExcelSheetName = reader.Name;
                                obj.CSVName = "";
                                obj.ExcelRowNumber = "";
                                obj.ErrorMessage = "An errror occured during reading excel file.";
                                obj.ExceptionMessage = ex.Message;
                                obj.DateCreated = DateTime.Now;
                                SaveErrorLogsToTable(obj, log);
                            }
                        }

                        do
                        {
                            ExcelSheets currentSheet = new ExcelSheets();
                            try
                            {
                                currentSheet = csvFile.GetCsvSheet(reader.Name.Trim());
                                if (currentSheet != null)
                                {

                                    string csvFilePath = $"{defaultContainerName}/{currentSheet.FolderName}/{currentSheet.CsvFileName}";
                                    var csvContent = string.Empty;

                                    List<AdditionalColumns> additionalColumns = csvFile.AdditionalColumns.ToList();

                                    int rowIndex = 0;
                                    List<int> writablecolumns = new List<int>();
                                    try
                                    {
                                        string additionaHeaders = string.Empty;
                                        string additionaHeadersValues = string.Empty;
                                        int patient_columnindex = -1;
                                        string patient_columndefaultvalue = "99999999";
                                        while (reader.Read())
                                        {
                                            try
                                            {
                                                List<string> arr = new List<string>();
                                                int lastNotEmptyCellIndex = 0;
                                                string headerColumnRange = string.Empty;
                                                if (rowIndex == currentSheet.HeaderRow)
                                                {
                                                    AdditionalColumns sourceSheetColumn = additionalColumns.LastOrDefault(c => c.ColumnName == "SOURCE_SHEET");
                                                    sourceSheetColumn.ColumnValue = reader.Name;
                                                    AdditionalColumns cellRangeColumn = additionalColumns.LastOrDefault(c => c.ColumnName == "CELL_RANGE");

                                                    int fieldCount = currentSheet.FixedColumn ? currentSheet.ColumnCount : reader.FieldCount;
                                                    fieldCount = Math.Min(fieldCount, 500);
                                                    for (int headerCellIndex = 0; headerCellIndex < fieldCount; headerCellIndex++)
                                                    {
                                                        string cellText = string.Empty;
                                                        try
                                                        {
                                                            cellText = GetFormattedValue(reader, headerCellIndex, CultureInfo.InvariantCulture, log);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            log.LogError(ex.Message);
                                                        }

                                                        if (currentSheet.SheetName == "Unit" && headerCellIndex == 0)
                                                        {
                                                            cellText = "TYPE";
                                                        }
                                                        cellText = cellText.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
                                                        cellText = Regex.Replace(cellText, @"\s+", " ");

                                                        if (!string.IsNullOrWhiteSpace(cellText))
                                                        {
                                                            writablecolumns.Add(headerCellIndex);
                                                            if (cellText.Contains("|") || cellText.Contains("\"") || cellText.Contains("\f") || cellText.Contains("\b") || cellText.Contains("\t"))
                                                            {
                                                                cellText = string.Format("\"{0}\"", cellText.Replace("\"", "\"\"").ToUpper().Trim());
                                                            }
                                                            else
                                                            {
                                                                cellText = cellText.ToUpper().Trim();
                                                            }
                                                            if (arr.Count == 0)
                                                            {
                                                                cellRangeColumn.ColumnValue = $"{GetExcelColumnName(headerCellIndex + 1)}{currentSheet.HeaderRow + 1}";
                                                            }
                                                            lastNotEmptyCellIndex = headerCellIndex;
                                                            arr.Add(cellText);
                                                        }

                                                    }


                                                    if (currentSheet.CsvFileName.StartsWith("INN_PT_"))
                                                    {
                                                        if (additionalColumns.Any(c => c.ColumnName == "PATIENT ID"))
                                                        {
                                                            var patientColumn = additionalColumns.FirstOrDefault(c => c.ColumnName == "PATIENT ID");
                                                            additionalColumns.Remove(patientColumn);
                                                        }

                                                        if (currentSheet.SheetName == "Referral Tracker" || currentSheet.SheetName == "Patient Nurse List" || currentSheet.SheetName == "Visit Scheduler")
                                                        {
                                                            patient_columnindex = arr.FindIndex(a => a == "PATIENT ID");
                                                            if (patient_columnindex == -1 && !additionalColumns.Any(c => c.ColumnName == "PATIENT ID"))
                                                            {
                                                                additionalColumns.Add(new AdditionalColumns() { ColumnName = "PATIENT ID", ColumnValue = patient_columndefaultvalue });
                                                            }
                                                        }
                                                    }


                                                    cellRangeColumn.ColumnValue = $"{cellRangeColumn.ColumnValue}:{GetExcelColumnName(lastNotEmptyCellIndex + 1)}{currentSheet.HeaderRow + 1}";
                                                    additionaHeaders = string.Join("|", additionalColumns.Select(c => c.ColumnName));
                                                    additionaHeadersValues = string.Join("|", additionalColumns.Select(c => c.ColumnValue));
                                                }
                                                else if (rowIndex >= currentSheet.HeaderRow + 1)
                                                {

                                                    for (int writablecolumnIndex = 0; writablecolumnIndex < writablecolumns.Count; writablecolumnIndex++)
                                                    {
                                                        int dataCellIndex = writablecolumns[writablecolumnIndex];

                                                        string cellText = string.Empty;
                                                        try
                                                        {

                                                            cellText = GetFormattedValue(reader, dataCellIndex, CultureInfo.InvariantCulture, log);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            log.LogError(ex.Message);

                                                        }

                                                        if (cellText.Contains("|") || cellText.Contains("\"") || cellText.Contains("\n") || cellText.Contains("\r") || cellText.Contains("\f") || cellText.Contains("\b") || cellText.Contains("\t"))
                                                        {
                                                            cellText = string.Format("\"{0}\"", cellText.Replace("\"", "\"\""));
                                                        }
                                                        cellText = cellText.Trim();
                                                        if (currentSheet.SheetName == "Unit")
                                                        {
                                                            if (dataCellIndex == 0 && !(cellText.ToUpper() == "NURSE TRAINING" || cellText.ToUpper() == "VISITS"))
                                                            {
                                                                break;
                                                            }
                                                        }
                                                        if (string.IsNullOrWhiteSpace(cellText) && writablecolumnIndex == patient_columnindex && patient_columnindex > -1)
                                                        {
                                                            cellText = patient_columndefaultvalue;
                                                        }
                                                        arr.Add(cellText);
                                                    }

                                                }
                                                else
                                                {

                                                }
                                                //if (arr.Any(a => a.Replace( "\\\"","").Replace("\\n", "").Replace("\\r", "").Replace("\\f", "").Replace("\\b", "").Replace( "\\t", "").Trim() != ""))
                                                if (arr.Any(a => a.Replace("\n", "").Replace("\r", "").Replace("\f", "").Replace("\b", "").Replace("\t", "").Trim() != ""))
                                                {
                                                    if (additionalColumns.Count > 0)
                                                    {
                                                        if (rowIndex == currentSheet.HeaderRow)
                                                        {

                                                            csvContent += additionaHeaders + "|" + string.Join("|", arr) + "\n";
                                                            currentSheet.TableColumns = csvContent;
                                                            CreateTableSchema(currentSheet, name, currentSheet.CsvFileName, log);
                                                        }
                                                        else
                                                        {
                                                            csvContent += additionaHeadersValues + "|" + string.Join("|", arr) + "\n";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        csvContent += string.Join("|", arr) + "\n";
                                                    }

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.LogError(ex.Message);
                                                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                                obj.ExcelFileName = name;
                                                obj.ExcelSheetName = reader.Name;
                                                obj.CSVName = currentSheet.CsvFileName;
                                                obj.ExcelRowNumber = rowIndex.ToString();
                                                obj.ErrorMessage = "An errror occured during processing excel rows.";
                                                obj.ExceptionMessage = ex.Message;
                                                obj.DateCreated = DateTime.Now;
                                                SaveErrorLogsToTable(obj, log);
                                            }
                                            rowIndex++;

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        log.LogError(ex.Message);
                                        ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                        obj.ExcelFileName = name;
                                        obj.ExcelSheetName = reader.Name;
                                        obj.CSVName = currentSheet.CsvFileName;
                                        obj.ExcelRowNumber = string.Empty;
                                        obj.ErrorMessage = "An errror occured before processing excel rows.";
                                        obj.ExceptionMessage = ex.Message;
                                        obj.DateCreated = DateTime.Now;
                                        SaveErrorLogsToTable(obj, log);
                                    }

                                    if (!string.IsNullOrWhiteSpace(csvContent))
                                    {
                                        BlobAttribute blob = new BlobAttribute(csvFilePath, FileAccess.Write);
                                        using (Stream destination = binder.Bind<Stream>(blob))
                                        {
                                            StreamWriter csv = new StreamWriter(destination, Encoding.UTF8);
                                            csv.Write(csvContent);
                                            csv.Close();
                                        }
                                    }
                                    else
                                    {
                                        log.LogError("No header or data are found in csv file");
                                        ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                        obj.ExcelFileName = name;
                                        obj.ExcelSheetName = reader.Name;
                                        obj.CSVName = currentSheet.CsvFileName;
                                        obj.ExcelRowNumber = string.Empty;
                                        obj.ErrorMessage = "No header or data are found in csv file";
                                        obj.ExceptionMessage = "No header or data are found in csv file";
                                        obj.DateCreated = DateTime.Now;
                                        SaveErrorLogsToTable(obj, log);
                                    }

                                }


                            }
                            catch (Exception ex)
                            {
                                log.LogError(ex.Message);
                                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                obj.ExcelFileName = name;
                                obj.ExcelSheetName = reader != null ? reader.Name : "";
                                obj.CSVName = currentSheet.CsvFileName;
                                obj.ExcelRowNumber = string.Empty;
                                obj.ErrorMessage = "Unable to process sheet.";
                                obj.ExceptionMessage = ex.Message;
                                obj.DateCreated = DateTime.Now;
                                SaveErrorLogsToTable(obj, log);

                            }
                        } while (reader.NextResult());

                    }
                }

            }
            else
            {
                log.LogError(errMsg);
                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                obj.ExcelFileName = name;
                obj.ExcelSheetName = string.Empty;
                obj.CSVName = string.Empty;
                obj.ExcelRowNumber = string.Empty;
                obj.ErrorMessage = errMsg;
                obj.ExceptionMessage = errMsg;
                obj.DateCreated = DateTime.Now;
                SaveErrorLogsToTable(obj, log);
            }
            try
            {
                var archiveFolder = config.GetConnectionStringOrSetting("ArchiveFolder");
                string azurestorageconnectionString = config.GetConnectionStringOrSetting("AzureWebJobsStorage");
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(azurestorageconnectionString);
                var blobClient = storageAccount.CreateCloudBlobClient();
                var container = blobClient.GetContainerReference(defaultContainerName);
                var blockBlob = container.GetBlockBlobReference($"Excels/{name}");

                var destBlob = container.GetBlockBlobReference($"Excels/Archive/{name}"); // ==> Copy source blob to destination container

                await destBlob.StartCopyAsync(blockBlob);
                //remove source blob after copy is done.            

                await blockBlob.DeleteIfExistsAsync();// ==> Delete blob
            }
            catch (Exception ex)
            {
                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                obj.ExcelFileName = name;
                obj.ExcelSheetName = string.Empty;
                obj.CSVName = string.Empty;
                obj.ExcelRowNumber = string.Empty;
                obj.ErrorMessage = "Failed to move or delete file.";
                obj.ExceptionMessage = ex.Message;
                obj.DateCreated = DateTime.Now;
                SaveErrorLogsToTable(obj, log);
            }
        }
        private static void CreateTableSchema(ExcelSheets excelSheet, string name, string csvFileName, ILogger log)
        {
            try
            {
                if (excelSheet.FolderName == "INN" || excelSheet.FolderName == "PIL" || excelSheet.FolderName == "PT")
                {
                    var str = Environment.GetEnvironmentVariable("SQL_ConnectionString", EnvironmentVariableTarget.Process);
                    using (SqlConnection conn = new SqlConnection(str))
                    {
                        using (SqlCommand cmd = new SqlCommand("raw.usp_CreateCsvTableSchema", conn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@tablename", SqlDbType.NVarChar).Value = excelSheet.TableName;
                            cmd.Parameters.Add("@foldername", SqlDbType.NVarChar).Value = excelSheet.FolderName;
                            cmd.Parameters.Add("@columns", SqlDbType.NVarChar).Value = excelSheet.TableColumns.Replace(@"\n", "");
                            conn.Open();
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                obj.ExcelFileName = name;
                obj.ExcelSheetName = excelSheet.SheetName;
                obj.CSVName = excelSheet.CsvFileName;
                obj.ExcelRowNumber = string.Empty;
                obj.ErrorMessage = ex.Message;
                obj.ExceptionMessage = ex.Message;
                obj.DateCreated = DateTime.Now;
                SaveErrorLogsToTable(obj, log);
            }
        }
        static string GetFormattedValue(IExcelDataReader reader, int columnIndex, CultureInfo culture, ILogger log)
        {

            try
            {

                var value = reader.GetValue(columnIndex);
                var strValue = Convert.ToString(value);
                if (!string.IsNullOrEmpty(strValue))
                {
                    strValue = strValue.Trim();
                    var formatString = reader.GetNumberFormatString(columnIndex);
                    DateTime dateValue;
                    if (formatString == @"[$-409]d\-mmm\-yy;@" || formatString == "[$-409]d-mmm-yy;@" ||
                                formatString == "[$-409]dd-mmm-yy;@" || formatString == "[$-409]mmmm d, yyyy;@"
                                || formatString == "m/d/yyyy;@" || formatString == "m/d/yyyy"
                                 || formatString == "mm/dd/yy" || formatString == "d-mmm-yy"
                                  || formatString == "yyyy-mm-dd" || formatString == "mm/dd/yy"
                                  || formatString == @"dd\-mmm\-yy"
                                )
                    {

                        if (DateTime.TryParse(strValue, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateValue))
                        {
                            if (dateValue == DateTime.MinValue)
                            {
                                return string.Empty;
                            }
                            else
                            {
                                return dateValue.ToString("dd'/'MM'/'yyyy");
                            }

                        }

                    }
                    else if (formatString == "m/d/yy h:mm" || formatString == "[$-409]m/d/yy h:mm AM/PM;@" ||
                        formatString == "m/d/yy h:mm;@"
                        )
                    {

                        if (DateTime.TryParse(strValue, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateValue))
                        {
                            if (dateValue == DateTime.MinValue)
                            {
                                return string.Empty;
                            }
                            else
                            {
                                return dateValue.ToString("dd'/'MM'/'yyyy HH:mm:ss");
                            }

                        }
                    }
                    if (formatString != null)
                    {
                        var format = new NumberFormat(formatString);
                        return format.Format(value, culture);
                    }
                    return Convert.ToString(value, culture).Trim();
                }

            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
            }
            return string.Empty;
        }

        static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        static async void SaveErrorLogsToTable(ExcelToCSVErrorLog objErrorLog, ILogger log)
        {
            try
            {
                CreateErrorLogTable(log);
                var str = Environment.GetEnvironmentVariable("SQL_ConnectionString", EnvironmentVariableTarget.Process);
                using (SqlConnection conn = new SqlConnection(str))
                {
                    conn.Open();
                    var text = @$"Insert INTO raw.ExcelToCSVErrorLogs values(
                        '{objErrorLog.ExcelFileName}',
                        '{objErrorLog.ExcelSheetName}',
                        '{objErrorLog.CSVName}',
                        '{objErrorLog.ExcelRowNumber}',
                        '{objErrorLog.ErrorMessage}',
                        '{objErrorLog.ExceptionMessage}',
                        '{objErrorLog.DateCreated}')";

                    using (SqlCommand cmd = new SqlCommand(text, conn))
                    {
                        var rows = await cmd.ExecuteNonQueryAsync();
                        log.LogInformation($"{rows} rows were inserted.");
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error occured when writing log - {ex.Message}");
            }
        }

        static async void CreateErrorLogTable(ILogger log)
        {
            try
            {
                string strQuery = @"IF (NOT EXISTS (SELECT * 
                                        FROM INFORMATION_SCHEMA.TABLES
                                        WHERE TABLE_SCHEMA = 'raw'
                                        AND TABLE_NAME = 'ExcelToCSVErrorLogs'))
                                  BEGIN
                                        CREATE TABLE[raw].[ExcelToCSVErrorLogs](
                                        [ExcelFileName][varchar](500) NULL,
                                        [ExcelSheetName] [varchar] (500) NULL,
	                                    [CSVName] [varchar] (500) NULL,
	                                    [RowNumber] [int] NULL,
	                                    [ErrorMessage] [varchar](500) NULL,
                                        [ExceptionMessage] [varchar](max) NULL,
                                        [DateCreated] [datetime] NULL
                                        ) ON[PRIMARY] TEXTIMAGE_ON[PRIMARY]
                                  END";
                var str = Environment.GetEnvironmentVariable("SQL_ConnectionString", EnvironmentVariableTarget.Process);

                using (SqlConnection conn = new SqlConnection(str))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(strQuery, conn))
                    {
                        var rows = await cmd.ExecuteNonQueryAsync();
                        log.LogInformation("Table ExcelToCSVErrorLog created.");
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Failed to create table ExcelToCSVErrorLog - {ex.Message}");
            }
        }

        static void ProcessINNRecruitmentSheet(ExcelFile csvFile, Stream excelFileInput, Binder binder, string name, string defaultContainerName, ILogger log)
        {
            try
            {
                XSSFWorkbook xssfwb;
                using (excelFileInput)
                {
                    xssfwb = new XSSFWorkbook(excelFileInput);
                    if (xssfwb != null && xssfwb.Count > 0)
                    {
                        DataFormatter dataFormatter = new DataFormatter(CultureInfo.InvariantCulture);
                        IFormulaEvaluator formulaEvaluator = WorkbookFactory.CreateFormulaEvaluator(xssfwb);
                        ExcelSheets currentSheet = new ExcelSheets();
                        for (int currentSheetIndex = 0; currentSheetIndex < xssfwb.Count; currentSheetIndex++)
                        {
                            ISheet currentExcelSheet = xssfwb.GetSheetAt(currentSheetIndex);
                            currentSheet = csvFile.GetCsvSheet(xssfwb.GetSheetAt(currentSheetIndex).SheetName);
                            if (currentSheet != null)
                            {
                                if (currentSheet != null)
                                {

                                    string csvFilePath = $"{defaultContainerName}/{currentSheet.FolderName}/{currentSheet.CsvFileName}";
                                    var csvContent = string.Empty;

                                    List<AdditionalColumns> additionalColumns = csvFile.AdditionalColumns.ToList();
                                    if (currentSheet.SheetName == "Open Locations")
                                    {
                                        additionalColumns.Add(new AdditionalColumns { ColumnName = "DIRECT_HIRE", ColumnValue = "0" });
                                    }


                                    List<int> writablecolumns = new List<int>();
                                    try
                                    {
                                        string additionaHeaders = string.Empty;
                                        string additionaHeadersValues = string.Empty;

                                        int daysOpenIndexColumn = -1;
                                        for (int rowIndex = currentSheet.HeaderRow; rowIndex <= currentExcelSheet.LastRowNum; rowIndex++)
                                        {
                                            try
                                            {
                                                List<string> arr = new List<string>();
                                                int lastNotEmptyCellIndex = 0;
                                                string headerColumnRange = string.Empty;
                                                if (rowIndex == currentSheet.HeaderRow)
                                                {

                                                    if (currentExcelSheet.GetRow(rowIndex) != null)
                                                    {
                                                        AdditionalColumns sourceSheetColumn = additionalColumns.LastOrDefault(c => c.ColumnName == "SOURCE_SHEET");
                                                        sourceSheetColumn.ColumnValue = currentExcelSheet.SheetName;
                                                        AdditionalColumns cellRangeColumn = additionalColumns.LastOrDefault(c => c.ColumnName == "CELL_RANGE");
                                                        int fieldCount = currentSheet.FixedColumn ? currentSheet.ColumnCount : (currentExcelSheet.GetRow(rowIndex).LastCellNum);
                                                        fieldCount = Math.Min(fieldCount, 500);
                                                        for (int headerCellIndex = 0; headerCellIndex < fieldCount; headerCellIndex++)
                                                        {
                                                            string cellText = string.Empty;
                                                            try
                                                            {
                                                                if (currentExcelSheet.GetRow(rowIndex).GetCell(headerCellIndex) != null)
                                                                {

                                                                    var currentcell = currentExcelSheet.GetRow(rowIndex).GetCell(headerCellIndex);
                                                                    if (currentcell.CellType == CellType.Numeric && currentcell.DateCellValue != DateTime.MinValue)
                                                                    {
                                                                        cellText = currentcell.DateCellValue.ToString("dd'/'MM'/'yyyy");
                                                                    }
                                                                    else if (currentcell.CellType == CellType.Formula && currentcell.CachedFormulaResultType == CellType.Numeric && currentcell.DateCellValue != DateTime.MinValue)
                                                                    {
                                                                        cellText = currentcell.DateCellValue.ToString("dd'/'MM'/'yyyy");
                                                                    }
                                                                    else
                                                                    {
                                                                        cellText = currentExcelSheet.GetRow(rowIndex).GetCell(currentcell.ColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString();
                                                                    }

                                                                }

                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                log.LogError(ex.Message);
                                                            }

                                                            cellText = cellText.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
                                                            cellText = Regex.Replace(cellText, @"\s+", " ");

                                                            if (!string.IsNullOrWhiteSpace(cellText))
                                                            {
                                                                writablecolumns.Add(headerCellIndex);
                                                                if (cellText.Contains("|") || cellText.Contains("\"") || cellText.Contains("\f") || cellText.Contains("\b") || cellText.Contains("\t"))
                                                                {
                                                                    cellText = string.Format("\"{0}\"", cellText.Replace("\"", "\"\"").ToUpper().Trim());
                                                                }
                                                                else
                                                                {
                                                                    cellText = cellText.ToUpper().Trim();
                                                                }
                                                                if (arr.Count == 0)
                                                                {
                                                                    cellRangeColumn.ColumnValue = $"{GetExcelColumnName(headerCellIndex + 1)}{currentSheet.HeaderRow + 1}";
                                                                }
                                                                lastNotEmptyCellIndex = headerCellIndex;
                                                                arr.Add(cellText);
                                                                if (currentSheet.SheetName == "Open Locations")
                                                                {
                                                                    if (currentSheet.SheetName == "Open Locations" && cellText.ToUpper() == "DAYS OPEN")
                                                                    {
                                                                        daysOpenIndexColumn = headerCellIndex;
                                                                    }
                                                                }

                                                            }

                                                        }


                                                        cellRangeColumn.ColumnValue = $"{cellRangeColumn.ColumnValue}:{GetExcelColumnName(lastNotEmptyCellIndex + 1)}{currentSheet.HeaderRow + 1}";
                                                        additionaHeaders = string.Join("|", additionalColumns.Select(c => c.ColumnName));

                                                    }
                                                }
                                                else if (rowIndex >= currentSheet.HeaderRow + 1)
                                                {
                                                    if (currentExcelSheet.GetRow(rowIndex) != null)
                                                    {
                                                        AdditionalColumns directHireColumn = additionalColumns.LastOrDefault(c => c.ColumnName == "DIRECT_HIRE");
                                                        if (directHireColumn != null)
                                                        {
                                                            directHireColumn.ColumnValue = "0";
                                                        }
                                                        for (int writablecolumnIndex = 0; writablecolumnIndex < writablecolumns.Count; writablecolumnIndex++)
                                                        {
                                                            int dataCellIndex = writablecolumns[writablecolumnIndex];

                                                            string cellText = string.Empty;
                                                            try
                                                            {
                                                                if (currentExcelSheet.GetRow(rowIndex).GetCell(dataCellIndex) != null)
                                                                {
                                                                    var currentcell = currentExcelSheet.GetRow(rowIndex).GetCell(dataCellIndex);

                                                                    cellText = GetFormattedValue(dataFormatter, formulaEvaluator, currentcell);
                                                                     
                                                                    if (currentcell.CellType == CellType.Numeric && currentcell.DateCellValue != DateTime.MinValue)
                                                                    {

                                                                        cellText = currentcell.DateCellValue.ToString("dd'/'MM'/'yyyy");
                                                                    }




                                                                    if (directHireColumn != null && daysOpenIndexColumn > -1 && daysOpenIndexColumn == dataCellIndex && currentExcelSheet.GetRow(rowIndex).GetCell(dataCellIndex).CellStyle.FillPattern == FillPattern.SolidForeground)
                                                                    {
                                                                        var scolor = ((NPOI.XSSF.UserModel.XSSFColor)currentExcelSheet.GetRow(rowIndex).GetCell(dataCellIndex).CellStyle.FillForegroundColorColor);

                                                                        Color color = Color.FromArgb(scolor.ARGB[0], scolor.ARGB[1], scolor.ARGB[2], scolor.ARGB[3]);
                                                                        if (color == Color.FromArgb(255, 255, 229, 255))
                                                                        {
                                                                            directHireColumn.ColumnValue = "1";
                                                                        }

                                                                    }

                                                                }

                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                log.LogError(ex.Message);

                                                            }

                                                            if (cellText.Contains("|") || cellText.Contains("\"") || cellText.Contains("\n") || cellText.Contains("\r") || cellText.Contains("\f") || cellText.Contains("\b") || cellText.Contains("\t"))
                                                            {
                                                                cellText = string.Format("\"{0}\"", cellText.Replace("\"", "\"\""));
                                                            }
                                                            cellText = cellText.Trim();

                                                            arr.Add(cellText);
                                                        }
                                                        additionaHeadersValues = string.Join("|", additionalColumns.Select(c => c.ColumnValue));
                                                    }

                                                }

                                                //if (arr.Any(a => a.Replace( "\\\"","").Replace("\\n", "").Replace("\\r", "").Replace("\\f", "").Replace("\\b", "").Replace( "\\t", "").Trim() != ""))
                                                if (arr.Any(a => a.Replace("\n", "").Replace("\r", "").Replace("\f", "").Replace("\b", "").Replace("\t", "").Trim() != ""))
                                                {
                                                    if (additionalColumns.Count > 0)
                                                    {
                                                        if (rowIndex == currentSheet.HeaderRow)
                                                        {

                                                            csvContent += additionaHeaders + "|" + string.Join("|", arr) + "\n";
                                                            currentSheet.TableColumns = csvContent;

                                                        }
                                                        else
                                                        {
                                                            csvContent += additionaHeadersValues + "|" + string.Join("|", arr) + "\n";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        csvContent += string.Join("|", arr) + "\n";
                                                    }

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.LogError(ex.Message);
                                                ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                                obj.ExcelFileName = name;
                                                obj.ExcelSheetName = currentExcelSheet.SheetName;
                                                obj.CSVName = currentSheet.CsvFileName;
                                                obj.ExcelRowNumber = rowIndex.ToString();
                                                obj.ErrorMessage = "An errror occured during processing excel rows.";
                                                obj.ExceptionMessage = ex.Message;
                                                obj.DateCreated = DateTime.Now;
                                                SaveErrorLogsToTable(obj, log);
                                            }


                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        log.LogError(ex.Message);
                                        ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                        obj.ExcelFileName = name;
                                        obj.ExcelSheetName = currentExcelSheet.SheetName;
                                        obj.CSVName = currentSheet.CsvFileName;
                                        obj.ExcelRowNumber = string.Empty;
                                        obj.ErrorMessage = "An errror occured before processing excel rows.";
                                        obj.ExceptionMessage = ex.Message;
                                        obj.DateCreated = DateTime.Now;
                                        SaveErrorLogsToTable(obj, log);
                                    }

                                    if (!string.IsNullOrWhiteSpace(csvContent))
                                    {
                                        BlobAttribute blob = new BlobAttribute(csvFilePath, FileAccess.Write);
                                        using (Stream destination = binder.Bind<Stream>(blob))
                                        {
                                            StreamWriter csv = new StreamWriter(destination, Encoding.UTF8);
                                            csv.Write(csvContent);
                                            csv.Close();
                                        }
                                    }
                                    else
                                    {
                                        log.LogError("No header or data are found in csv file");
                                        ExcelToCSVErrorLog obj = new ExcelToCSVErrorLog();
                                        obj.ExcelFileName = name;
                                        obj.ExcelSheetName = currentExcelSheet.SheetName;
                                        obj.CSVName = currentSheet.CsvFileName;
                                        obj.ExcelRowNumber = string.Empty;
                                        obj.ErrorMessage = "No header or data are found in csv file";
                                        obj.ExceptionMessage = "No header or data are found in csv file";
                                        obj.DateCreated = DateTime.Now;
                                        SaveErrorLogsToTable(obj, log);
                                    }

                                }
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }


        }


        //
        // Get formatted value as string from the specified cell
        //
        private static string GetFormattedValue(DataFormatter dataFormatter, IFormulaEvaluator formulaEvaluator, ICell cell)
        {

            string returnValue = string.Empty;
            if (cell != null)
            {
                try
                {
                    // Get evaluated and formatted cell value
                    returnValue = dataFormatter.FormatCellValue(cell, formulaEvaluator);
                }
                catch
                {
                    // When failed in evaluating the formula, use stored values instead...
                    // and set cell value for reference from formulae in other cells...
                    if (cell.CellType == CellType.Formula)
                    {
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.String:
                                returnValue = cell.StringCellValue;
                                cell.SetCellValue(cell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                returnValue = dataFormatter.FormatRawCellContents(cell.NumericCellValue, 0, cell.CellStyle.GetDataFormatString());
                                cell.SetCellValue(cell.NumericCellValue);
                                break;
                            case CellType.Boolean:
                                returnValue = cell.BooleanCellValue.ToString();
                                cell.SetCellValue(cell.BooleanCellValue);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return (returnValue ?? string.Empty).Trim();
        }

        //
        // Get unformatted value as string from the specified cell
        //
        private static string GetUnformattedValue(DataFormatter dataFormatter, IFormulaEvaluator formulaEvaluator, ICell cell)
        {
            string returnValue = string.Empty;
            if (cell != null)
            {
                try
                {
                    // Get evaluated cell value
                    returnValue = (cell.CellType == CellType.Numeric ||
               (cell.CellType == CellType.Formula &&
               cell.CachedFormulaResultType == CellType.Numeric)) ?
                   formulaEvaluator.EvaluateInCell(cell).NumericCellValue.ToString() :
                   dataFormatter.FormatCellValue(cell, formulaEvaluator);
                }
                catch
                {
                    // When failed in evaluating the formula, use stored values instead...
                    // and set cell value for reference from formulae in other cells...
                    if (cell.CellType == CellType.Formula)
                    {
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.String:
                                returnValue = cell.StringCellValue;
                                cell.SetCellValue(cell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                returnValue = cell.NumericCellValue.ToString();
                                cell.SetCellValue(cell.NumericCellValue);
                                break;
                            case CellType.Boolean:
                                returnValue = cell.BooleanCellValue.ToString();
                                cell.SetCellValue(cell.BooleanCellValue);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return (returnValue ?? string.Empty).Trim();
        }
    }
}
