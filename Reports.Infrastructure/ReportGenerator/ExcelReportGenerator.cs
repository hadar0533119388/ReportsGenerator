using ClosedXML.Excel;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using Reports.Infrastructure.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.ReportGenerator
{
    public class ExcelReportGenerator : IReportGenerator
    {
        public Enums.ReportType Type => ReportType.Excel;
        private readonly IReportRepositoryAdoNet repositoryAdoNet;
        private readonly IReportRepositoryDapper repositoryDapper;
        private readonly ILogger logger;
        private readonly string libreOffice;


        public ExcelReportGenerator(IReportRepositoryAdoNet repositoryAdoNet, IReportRepositoryDapper repositoryDapper, ILogger logger, string libreOffice)
        {
            this.repositoryAdoNet = repositoryAdoNet;
            this.repositoryDapper = repositoryDapper;
            this.logger = logger;
            this.libreOffice = libreOffice;
        }

        public async Task<byte[]> ExecuteAsync(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {

                Manifest manifest = await repositoryDapper.GetManifestByManifestIDAsync(request.ManifestID);

                //For SP parameters
                request.Parameters["ManifestID"] = request.ManifestID;

                DataSet dataSet = repositoryAdoNet.GetData(request, reportDtl);

                if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    byte[] excel = GenerateExcel(dataSet, request, reportDtl, manifest);

                    if (request.IsPrint)
                    {
                        PrintExcelDocument(excel, request.PrinterName);
                    }

                    return excel;
                }
                else
                {
                    logger.WriteLog($"No data was returned for the report.");
                    return null;
                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Execute Async: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Execute Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private void PrintExcelDocument(byte[] excel, string printerName)
        {
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            try
            {
                // Saving the Excel array to a temporary file
                File.WriteAllBytes(tempFilePath, excel);

                // Check the printer status
                PrintQueueStatus status = GetPrinterStatus(printerName);

                if (status.HasFlag(PrintQueueStatus.PaperProblem) ||
                status.HasFlag(PrintQueueStatus.Offline) ||
                status.HasFlag(PrintQueueStatus.Error))
                {
                    throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]);
                }

                // Preparing to run LibreOffice for printing
                var printProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = libreOffice,
                        Arguments = $"--headless --nologo --norestore --pt \"{printerName}\" \"{tempFilePath}\"",
                        RedirectStandardError = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                // Running the process (LibreOffice print)
                using (Process process = new Process { StartInfo = printProcess.StartInfo })
                {
                    process.Start();

                    string errorOutput = process.StandardError.ReadToEnd();
                    string standardOutput = process.StandardOutput.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        logger.WriteLog($"Error to Print Excel Document. ExitCode: {process.ExitCode}");
                        if (!string.IsNullOrEmpty(errorOutput))
                        {
                            logger.WriteLog($"Error details: {errorOutput}");
                        }
                        throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]);
                    }
                    else
                    {
                        logger.WriteLog($"Print Excel Document completed successfully.");
                    }
                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Print Excel Document: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Print Excel Document: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, $"{ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]} : {ex.Message}");
            }
            finally
            {
                // Deleting the temporary file
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }

        private PrintQueueStatus GetPrinterStatus(string printerName)
        {
            using (LocalPrintServer printServer = new LocalPrintServer())
            {
                foreach (PrintQueue queue in printServer.GetPrintQueues())
                {
                    if (queue.Name.Equals(printerName, StringComparison.OrdinalIgnoreCase))
                    {
                        queue.Refresh();
                        return queue.QueueStatus;
                    }
                }
            }
            throw new CustomException((int)ErrorMessages.ErrorCodes.UnknownPrinter, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.UnknownPrinter]);
        }


        private byte[] GenerateExcel(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                if (Enum.TryParse(reportDtl.Template, true, out GenerateExcel generateExcel))
                {
                    switch (generateExcel)
                    {
                        case Enums.GenerateExcel.GenerateSUMentries9Report:
                            return GenerateSUMentries9Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateInvBckReport:
                            return GenerateInvBckReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMvalindex3Report:
                            return GenerateSUMvalindex3Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMdeliveryGush8Report:
                            return GenerateSUMdeliveryGush8Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMdeliveryLines8Report:
                            return GenerateSUMdeliveryLines8Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateCarsInShowroomsReport:
                            return GenerateCarsInShowroomsReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMqntIndex1Report:
                            return GenerateSUMqntIndex1Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateDTLentries9Report:
                            return GenerateDTLentries9Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateZeroInventory11Report:
                            return GenerateZeroInventory11Report(dataSet, request, reportDtl, manifest);

                        default:
                            break;
                    }
                }

                return null;
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Generate Excel: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate Excel: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        //Bold headings and unbold data
        private void ApplyTableStyleBoldHeadings(IXLTable tableRange)
        {
            tableRange.Theme = XLTableTheme.None;
            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            tableRange.Style.Font.FontColor = XLColor.Black;
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.None;

            var headerRow = tableRange.Range(1, 1, 1, tableRange.ColumnCount());
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            foreach (var cell in headerRow.Cells())
            {
                cell.Style.Alignment.WrapText = true;
                cell.Value = cell.Value.ToString().Replace(" ", "\n");
            }

        }

        //Bold data and unbold headings
        private void ApplyTableStyleBoldData(IXLTable tableRange)
        {
            tableRange.Theme = XLTableTheme.None;
            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            tableRange.Style.Font.FontColor = XLColor.Black;
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.None;


            var headerRow = tableRange.Range(1, 1, 1, tableRange.ColumnCount());
            headerRow.Style.Font.Bold = false;
            headerRow.Style.Font.Underline = XLFontUnderlineValues.Single;


            var dataRange = tableRange.DataRange;
            dataRange.Style.Font.Bold = true;
        }

        private void AddHeader2(IXLWorksheet worksheet, ReportRequest request, ReportDtl reportDtl, Manifest manifest, int row, int numberOfColumns)
        {
            string manifestDetails = $"{manifest.GroupName}, בונדד {manifest.ManifestID} ({manifest.TermName})";


            string currentDateAndUser = $"{DateTime.Now.ToString("dd/MM/yyyy HH:mm")}, {request.User}";

            worksheet.Cell(row, 1).Value = manifestDetails;
            worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, 4)).Merge();


            worksheet.Cell(row, 5).Value = reportDtl.ReportName;
            worksheet.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell(row, 5).Style.Font.Bold = true;
            worksheet.Range(worksheet.Cell(row, 5), worksheet.Cell(row, numberOfColumns - 4)).Merge();


            worksheet.Cell(row, numberOfColumns - 3).Value = currentDateAndUser;
            worksheet.Cell(row, numberOfColumns - 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Range(worksheet.Cell(row, numberOfColumns - 3), worksheet.Cell(row, numberOfColumns)).Merge();

            row++;
            worksheet.Range(row, 1, row, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

            worksheet.Columns().AdjustToContents();
        }

        private void ApplyNumberFormatToSheet(IXLWorksheet worksheet)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.DataType == XLDataType.Number || !string.IsNullOrEmpty(cell.FormulaA1))
                {
                    cell.Style.NumberFormat.Format = "#,##0";
                }
            }
        }

        private void PrintSettings(IXLWorksheet worksheet)
        {
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

            worksheet.RightToLeft = true;

            worksheet.Style.Font.FontSize = 10;

            worksheet.PageSetup.FitToPages(1, 0);

            double marginSize = 0.5;
            worksheet.PageSetup.Margins.Left = marginSize;
            worksheet.PageSetup.Margins.Right = marginSize;
            worksheet.PageSetup.Margins.Top = marginSize;
            worksheet.PageSetup.Margins.Bottom = marginSize;

            worksheet.PageSetup.CenterHorizontally = true;
        }

        private void RemoveColumnsByName(DataTable table, List<string> columnsToRemove)
        {
            foreach (var columnName in columnsToRemove)
            {
                if (table.Columns.Contains(columnName))
                {
                    table.Columns.Remove(columnName);
                }
            }
        }

        private void ApplyColumnFormula(IXLWorksheet worksheet, DataTable dataTable, int startCalc, int endCalc, List<string> columnNames, string formulaType)
        {
            foreach (var columnName in columnNames)
            {
                if (!dataTable.Columns.Contains(columnName))
                    continue;

                int columnIndex = dataTable.Columns[columnName].Ordinal + 1;
                string columnLetter = XLHelper.GetColumnLetterFromNumber(columnIndex);
                string formula = $"{formulaType}({columnLetter}{startCalc}:{columnLetter}{endCalc - 1})";


                worksheet.Cell(endCalc, columnIndex).FormulaA1 = formula;
                worksheet.Cell(endCalc, columnIndex).Style.Font.Bold = true;
            }
        }
       
        private byte[] GenerateSUMentries9Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[2].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentID") && request.Parameters["CustomsAgentID"] != null)
                    {
                        string customsAgentID = request.Parameters["CustomsAgentID"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentID))
                        {
                            string customsAgent = dataSet.Tables[3].Rows[0]["CustomsAgent"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"סוכן: {customsAgent} ({customsAgentID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תיק סוכן: {customsAgentFile}");
                        }
                    }

                    if (request.Parameters.ContainsKey("ConsignmentID") && request.Parameters["ConsignmentID"] != null)
                    {
                        string consignmentID = request.Parameters["ConsignmentID"]?.ToString();
                        if (!string.IsNullOrEmpty(consignmentID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"למזהה מטען: {consignmentID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("DeliverySiteNumber") && request.Parameters["DeliverySiteNumber"] != null)
                    {
                        string deliverySiteNumber = request.Parameters["DeliverySiteNumber"]?.ToString();
                        if (!string.IsNullOrEmpty(deliverySiteNumber))
                        {
                            string name = dataSet.Tables[4].Rows[0]["Name"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"אתר מקור: {name}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות", "ערך בש\"ח", "20”", "40”", "נפח" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    currentRow += 3;

                    worksheet.Cell(currentRow, 9).Value = "ריכוז כניסות לפי מטבע";
                    worksheet.Cell(currentRow, 9).Style.Font.Underline = XLFontUnderlineValues.Single;
                    worksheet.Cell(currentRow, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 9).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(currentRow, 9), worksheet.Cell(currentRow, 11)).Merge();

                    currentRow++;

                    var table2 = worksheet.Cell(currentRow, 9).InsertTable(dataSet.Tables[1]);
                    ApplyTableStyleBoldData(table2);

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMentries9 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateInvBckReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    if (request.IsPrint)
                    {
                        var columnsToDelete = new List<string> { "תיק סוכן", "כמות מוצהרת", "הערה" };
                        RemoveColumnsByName(dataSet.Tables[0], columnsToDelete);
                    }

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledCAID") && request.Parameters["BilledCAID"] != null)
                    {
                        string billedCAID = request.Parameters["BilledCAID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedCAID))
                        {
                            string billedCA = dataSet.Tables[2].Rows[0]["BilledCA"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"לסוכן: {billedCA} ({billedCAID})");
                        }
                    }


                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תיק: {customsAgentFile}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("GushState") && request.Parameters["GushState"] != null)
                    {
                        if (int.TryParse(request.Parameters["GushState"].ToString(), out int gushState))
                        {
                            if (filter.Length > 0) filter.Append(", ");

                            if (gushState == 0)
                            {
                                filter.Append("גושים פתוחים בלבד");
                            }
                            else if (gushState == 1)
                            {
                                filter.Append("גושים סגורים בלבד");
                            }
                        }
                    }

                    if (request.Parameters.ContainsKey("Location") && request.Parameters["Location"] != null)
                    {
                        string location = request.Parameters["Location"]?.ToString();
                        if (!string.IsNullOrEmpty(location))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"איתור: {location}");
                        }
                    }


                    if (request.Parameters.ContainsKey("FromInv") && request.Parameters["FromInv"] != null && request.Parameters.ContainsKey("ToInv") && request.Parameters["ToInv"] != null)
                    {
                        string fromInv = Convert.ToInt32(request.Parameters["FromInv"]).ToString("N0");
                        string toInv = Convert.ToInt32(request.Parameters["ToInv"]).ToString("N0");
                        if (!string.IsNullOrEmpty(fromInv) && !string.IsNullOrEmpty(toInv))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מיתרה {fromInv} עד יתרה {toInv}");
                        }
                    }

                    if (request.Parameters.ContainsKey("Pcustoms") && request.Parameters["Pcustoms"] != null)
                    {
                        string pcustoms = request.Parameters["Pcustoms"]?.ToString();

                        if (filter.Length > 0) filter.Append(", ");

                        if (pcustoms == "VN")
                        {
                            filter.Append("לרכבים בלבד");
                        }
                        else if (pcustoms == "PP")
                        {
                            filter.Append("ללא רכבים בלבד");
                        }

                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות מוצהרת", "טרם התקבל", "יתרה", "כמות משוחררת", "כמות ברשות מוסמכת" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();
                    
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate InvBck Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMvalindex3Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("ConsignmentID") && request.Parameters["ConsignmentID"] != null)
                    {
                        string consignmentID = request.Parameters["ConsignmentID"]?.ToString();
                        if (!string.IsNullOrEmpty(consignmentID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"הצהרת אחסנה: {consignmentID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }
                    if (request.Parameters.ContainsKey("GushState") && request.Parameters["GushState"] != null)
                    {
                        if (int.TryParse(request.Parameters["GushState"].ToString(), out int gushState))
                        {
                            if (filter.Length > 0) filter.Append(", ");

                            if (gushState == 0)
                            {
                                filter.Append("גושים פתוחים בלבד");
                            }
                            else if (gushState == 1)
                            {
                                filter.Append("גושים סגורים בלבד");
                            }
                        }
                    }
                    if (request.Parameters.ContainsKey("ZeroInventory") && request.Parameters["ZeroInventory"] != null)
                    {
                        string zeroInventory = request.Parameters["ZeroInventory"]?.ToString();
                        if (!string.IsNullOrEmpty(zeroInventory) && zeroInventory == "Y")
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append("כולל יתרות אפס");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות בפתיחה", "ערך בפתיחה", "נפח בפתיחה", "שטח בפתיחה", "משקל בפתיחה", "יתרת כמות", "יתרת ערך", "יתרת נפח", "יתרת שטח", "יתרת משקל" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);
                    
                    worksheet.Columns().AdjustToContents();

                    worksheet.Column(5).Width = 15; 
                    worksheet.Column(6).Width = 15;

                    foreach (var row in worksheet.Rows())
                    {
                        
                        row.Cell(5).Style.Alignment.WrapText = true;
                        row.Cell(6).Style.Alignment.WrapText = true;                       
                    }

                    
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMvalindex3 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMdeliveryGush8Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate) 
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות" };
                    var columnsToCount = new List<string> { "מונה" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();
                    
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMdeliveryGush8 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMdeliveryLines8Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            string code = dataSet.Tables[1].Rows[0]["Code"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({code})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("LineID") && request.Parameters["LineID"] != null)
                    {
                        string lineID = request.Parameters["LineID"]?.ToString();
                        if (!string.IsNullOrEmpty(lineID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"למזהה: {lineID}");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות שנמסרה מהמזהה" };
                    var columnsToCount = new List<string> { "תעודת מסירה" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();
                                        
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMdeliveryLines8 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateCarsInShowroomsReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                    worksheet.RightToLeft = true;

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 4);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns + 4)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);
                    
                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();
                    worksheet.Column(1).Width += 5;
                    
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate CarsInShowrooms Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMqntIndex1Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }                    

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }
                    
                    if (request.Parameters.ContainsKey("GushState") && request.Parameters["GushState"] != null)
                    {
                        if (int.TryParse(request.Parameters["GushState"].ToString(), out int gushState))
                        {
                            if (filter.Length > 0) filter.Append(", ");

                            if (gushState == 0)
                            {
                                filter.Append("גושים פתוחים בלבד");
                            }
                            else if (gushState == 1)
                            {
                                filter.Append("גושים סגורים בלבד");
                            }
                        }
                    }

                    if (request.Parameters.ContainsKey("FromInv") && request.Parameters["FromInv"] != null && request.Parameters.ContainsKey("ToInv") && request.Parameters["ToInv"] != null)
                    {
                        string fromInv = Convert.ToInt32(request.Parameters["FromInv"]).ToString("N0");
                        string toInv = Convert.ToInt32(request.Parameters["ToInv"]).ToString("N0");
                        if (!string.IsNullOrEmpty(fromInv) && !string.IsNullOrEmpty(toInv))
                        {                            
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מיתרה {fromInv} עד יתרה {toInv}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "יתרת כמות" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    worksheet.Column(4).Width = 25;
                    worksheet.Column(5).Width = 25;
                    worksheet.Column(6).Width = 25;
                    worksheet.Column(13).Width = 25;

                    foreach (var row in worksheet.Rows())
                    {
                        row.Cell(4).Style.Alignment.WrapText = true;
                        row.Cell(5).Style.Alignment.WrapText = true;
                        row.Cell(6).Style.Alignment.WrapText = true;
                        row.Cell(13).Style.Alignment.WrapText = true;
                    }
                    

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMqntIndex1 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateDTLentries9Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentID") && request.Parameters["CustomsAgentID"] != null)
                    {
                        string customsAgentID = request.Parameters["CustomsAgentID"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentID))
                        {
                            string customsAgent = dataSet.Tables[2].Rows[0]["CustomsAgent"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"סוכן: {customsAgent} ({customsAgentID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תיק סוכן: {customsAgentFile}");
                        }
                    }

                    if (request.Parameters.ContainsKey("ConsignmentID") && request.Parameters["ConsignmentID"] != null)
                    {
                        string consignmentID = request.Parameters["ConsignmentID"]?.ToString();
                        if (!string.IsNullOrEmpty(consignmentID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מטען: {consignmentID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("LineID") && request.Parameters["LineID"] != null)
                    {
                        string lineID = request.Parameters["LineID"]?.ToString();
                        if (!string.IsNullOrEmpty(lineID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מזהה: {lineID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("DeliverySiteNumber") && request.Parameters["DeliverySiteNumber"] != null)
                    {
                        string deliverySiteNumber = request.Parameters["DeliverySiteNumber"]?.ToString();
                        if (!string.IsNullOrEmpty(deliverySiteNumber))
                        {
                            string name = dataSet.Tables[3].Rows[0]["Name"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"אתר מקור: {name}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 1).Style.Alignment.WrapText = true;
                    worksheet.Row(currentRow).Height = 25;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate DTLentries9 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }


        private byte[] GenerateZeroInventory11Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 3);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ZeroInvDate") && request.Parameters["ZeroInvDate"] != null)
                    {
                        string zeroInvDate = request.Parameters["ZeroInvDate"]?.ToString();
                        if (!string.IsNullOrEmpty(zeroInvDate)
                            && DateTime.TryParseExact(zeroInvDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedZeroInvDate))
                        {
                            filter.Append($"גושים שיתרתם התאפסה החל מ: {parsedZeroInvDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }
             
                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("AlreadyClosed") && request.Parameters["AlreadyClosed"] != null)
                    {
                        string alreadyClosed = request.Parameters["AlreadyClosed"]?.ToString();
                        if (!string.IsNullOrEmpty(alreadyClosed) && alreadyClosed == "Y")
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append("כולל גושים סגורים");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns + 3)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;                   

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    worksheet.Column(5).Width = 10;
                    worksheet.Column(6).Width = 10;
                    

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate ZeroInventory11 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
    }
}
