using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
//using System.Threading;
using ExcelDataReader;
using ExcelNumberFormat;
using Microsoft.Azure.Storage;
using Microsoft.Azure.Storage.Blob;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace MRNAFExportExcelToCSV
{

    class ExcelSheets
    {

        public string SheetName { get; set; }
        public string CsvFileName { get; set; }
    }


    public static class ExcelToCsv
    {
        [FunctionName("ExcelToCsv")]
        public static void Run([BlobTrigger("%ContainerName%/Excels/{name}", Connection = "AzureWebJobsStorage")] Stream excelFileInput, Binder binder, string name, ILogger log, ExecutionContext context)
        {
            log.LogInformation($"C# Blob trigger function executed at: {DateTime.Now}");

            var config = new ConfigurationBuilder()
           .SetBasePath(context.FunctionAppDirectory)
           .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            string defaultContainerName = config.GetConnectionStringOrSetting("ContainerName");

            List<ExcelSheets> excelSheets = new List<ExcelSheets>();
            List<AdditionalColumns> additionalColumns = new List<AdditionalColumns>();

            string folderPath = string.Empty;

            bool onlyFixedColumns = false;
            int columnCount = -1;

            int headerRow = 0;
            if (name.Contains("/"))
            {
                return;
            }
            else if (!name.EndsWith(".xls") && !name.EndsWith(".xlsx") && !name.EndsWith(".xlsm"))
            {
                log.LogInformation($"unable to covnert file: {name} into csv");
                return;
            }
            else
            {
                if (name.ToLower().StartsWith("vendor governance tracker"))
                {
                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "Govn_Tracker", CsvFileName = "Vendor_Governance.csv" }
                         };

                    folderPath = defaultContainerName + "/Vendors/{0}";
                }
                else if (name.ToLower().StartsWith("vendor scorecard tracker"))
                {
                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "Scorecard_Tracker", CsvFileName = "Vendor_Scorecard.csv" }
                         };

                    folderPath = defaultContainerName + "/Vendors/{0}";
                }
                else if (name.ToLower().StartsWith("reference_tables"))
                {
                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "STATIC_DECODES", CsvFileName = "Reference_Codes.csv" }
                         };

                    folderPath = defaultContainerName + "/Reference/{0}";
                }
                else if (name.ToLower().StartsWith("nurse list"))
                {

                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "Nurse List", CsvFileName = "Nurse_List_.csv" }
                         };

                    folderPath = defaultContainerName + "/Nursing/{0}";
                }
                else if (name.ToLower().StartsWith("vap tracker"))
                {
                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "VAP Table", CsvFileName = "Nurse_VAP_Table_.csv" }
                         };

                    folderPath = defaultContainerName + "/Nursing/{0}";
                }
                else if (name.ToLower().StartsWith("nob sessions"))
                {
                    var namearray = name.Split(" ");
                    if (namearray.Length < 3)
                    {
                        log.LogInformation($"unable to covnert file: {name} into csv");
                        return;
                    }
                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){
                        new ExcelSheets() { SheetName = "Sessions", CsvFileName = $"Nurse_Sessions_{namearray[2]}.csv" },
                                                            new ExcelSheets() { SheetName = "Training", CsvFileName = $"Nurse_Training_{namearray[2]}.csv" },
                                                            new ExcelSheets() { SheetName = "Nurse Docs", CsvFileName = $"Nurse_Docs_{namearray[2]}.csv" }
                         };
                    additionalColumns.Add(new AdditionalColumns() { ColumnName = "REGION", ColumnValue = namearray[2] });
                    folderPath = defaultContainerName + "/Nursing/{0}";
                }
                else if (name.ToUpper().StartsWith("INN"))
                {
                    var namearray = name.Split(" ");
                    if (namearray.Length < 2)
                    {
                        log.LogInformation($"unable to covnert file: {name} into csv");
                        return;
                    }
                    else
                    {
                        string countryCode = namearray[0].Substring(3);
                        headerRow = 0;
                        excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "Referral Tracker", CsvFileName = $"INN_PT_{countryCode}_RT1.csv" },
                                                            new ExcelSheets() { SheetName = "Patient Nurse List", CsvFileName = $"INN_PT_{countryCode}_PNL.csv" },
                                                            new ExcelSheets() { SheetName = "Visit Scheduler", CsvFileName = $"INN_PT_{countryCode}_VS1.csv" },
                                                            new ExcelSheets() { SheetName = "Nurse Database", CsvFileName = $"INN_PT_{countryCode}_ND1.csv" },
                                                            new ExcelSheets() { SheetName = "All Projects Information", CsvFileName = $"INN_PT_{countryCode}_API.csv" },
                                                            new ExcelSheets() { SheetName = "SNS Referral Tracker", CsvFileName = $"INN_PT_{countryCode}_SRT.csv" },
                                                            new ExcelSheets() { SheetName = "Site Nurse List", CsvFileName = $"INN_PT_{countryCode}_SNL.csv" },
                                                            new ExcelSheets() { SheetName = "SNS Visit Scheduler", CsvFileName = $"INN_PT_{countryCode}_SVS.csv" }
                         };
                        folderPath = defaultContainerName + "/INN/{0}";
                    }
                }
                else if (name.ToLower().Contains("_project issue log_"))
                {
                    var namearray = name.Split("_");
                    if (namearray.Length < 2)
                    {
                        log.LogInformation($"unable to covnert file: {name} into csv");
                        return;
                    }
                    else
                    {
                        string projectNumber = namearray[0];
                        headerRow = 1;
                        excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "Issues and Deviations", CsvFileName = $"IssueLog_Issues_and_Deviations_{projectNumber}.csv" },
                                                            new ExcelSheets() { SheetName = "Clinical Incidents", CsvFileName = $"IssueLog_Clinical_Incidents_{projectNumber}.csv" },

                         };
                        additionalColumns.Add(new AdditionalColumns() { ColumnName = "PROJECT NUMBER", ColumnValue = namearray[0] });
                        folderPath = defaultContainerName + "/PIL/{0}";
                    }
                }
                else if (name.ToLower().StartsWith("old hts opportunities"))
                {
                    headerRow = 0;
                    excelSheets = new List<ExcelSheets>(){new ExcelSheets() { SheetName = "Sheet3", CsvFileName = $"Historic_Projects.csv" }

                         };
                    onlyFixedColumns = true;
                    columnCount = 2;
                    folderPath = defaultContainerName + "/Other/{0}";
                }
                else
                {
                    var namearray = name.Split("_");
                    if (namearray.Length < 3 && !(Array.IndexOf(namearray, "Project Plan Tracker") > 1))
                    {
                        log.LogInformation($"unable to covnert file: {name} into csv");
                        return;
                    }
                    else
                    {
                        headerRow = 2;
                        additionalColumns.Add(new AdditionalColumns() { ColumnName = "PT_CODE", ColumnValue = namearray[0] });
                        excelSheets = new List<ExcelSheets>(){
                            new ExcelSheets() { SheetName = "SIV Tracker", CsvFileName =$"HTS_PT_{namearray[0]}_{namearray[1]}_ST.csv" },
                                                            new ExcelSheets() { SheetName = "Patient List", CsvFileName = $"HTS_PT_{namearray[0]}_{namearray[1]}_PL.csv" },
                                                            new ExcelSheets() { SheetName = "Nurse List", CsvFileName = $"HTS_PT_{namearray[0]}_{namearray[1]}_NL.csv" },
                                                            new ExcelSheets() { SheetName = "Visit Tracker", CsvFileName = $"HTS_PT_{namearray[0]}_{namearray[1]}_VT.csv" },
                                                            new ExcelSheets() { SheetName = "MRN Internal DCF Tracker", CsvFileName = $"HTS_PT_{namearray[0]}_{namearray[1]}_DCF.csv"}
                        };
                        folderPath = defaultContainerName + "/PT/{0}";
                    }
                }
                additionalColumns.Add(new AdditionalColumns() { ColumnName = "SOURCE_SPREAD_SHEET", ColumnValue = name });
                additionalColumns.Add(new AdditionalColumns() { ColumnName = "TIMEDATE_SNAPSHOT", ColumnValue = DateTime.Now.ToString("dd'/'MM'/'yyyy HH:mm:ss") });
            }


            log.LogInformation($"Do your processing on the excelFileInput file here.");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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
                    }
                }

                do
                {
                    try
                    {

                        if (!excelSheets.Any(e => e.SheetName == reader.Name.Trim()))
                        {
                            continue;
                        }


                        var csvFile = excelSheets.FirstOrDefault(e => e.SheetName == reader.Name.Trim()).CsvFileName;
                        string csvFilePath = string.Format(folderPath, csvFile);
                        var csvContent = string.Empty;
                        BlobAttribute blob = new BlobAttribute(csvFilePath, FileAccess.Write);

                        using (Stream destination = binder.Bind<Stream>(
                       blob))
                        {
                            int rowIndex = 0;
                            List<int> writablecolumns = new List<int>();
                            try
                            {
                                string additionaHeaders = string.Join("|", additionalColumns.Select(c => c.ColumnName));
                                string additionaHeadersValues = string.Join("|", additionalColumns.Select(c => c.ColumnValue));
                                while (reader.Read())
                                {

                                    List<string> arr = new List<string>();
                                    if (rowIndex == headerRow)
                                    {
                                        int fieldCount = onlyFixedColumns ? columnCount : reader.FieldCount;
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
                                                arr.Add(cellText);
                                            }
                                        }
                                    }
                                    else if (rowIndex >= headerRow + 1)
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
                                            if (rowIndex == headerRow)
                                            {

                                                csvContent += additionaHeaders + "|" + string.Join("|", arr) + "\n";
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

                                    rowIndex++;

                                }
                            }
                            catch (Exception ex)
                            {
                                log.LogError(ex.Message);
                            }
                            StreamWriter csv = new StreamWriter(destination, Encoding.UTF8);
                            csv.Write(csvContent);
                            csv.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        log.LogError(ex.Message);
                    }
                } while (reader.NextResult());

            }

            var archiveFolder = config.GetConnectionStringOrSetting("ArchiveFolder");
            string azurestorageconnectionString = config.GetConnectionStringOrSetting("AzureWebJobsStorage");
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(azurestorageconnectionString);
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference(defaultContainerName);
            var blockBlob = container.GetBlockBlobReference($"Excels/{name}");

            var destBlob = container.GetBlockBlobReference($"Excels/Archive/{name}"); // ==> Copy source blob to destination container


            destBlob.StartCopy(blockBlob);
            //remove source blob after copy is done.            

            blockBlob.DeleteIfExists();// ==> Delete blob
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
                        formatString == "	m/d/yy h:mm;@"
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
    }
}
