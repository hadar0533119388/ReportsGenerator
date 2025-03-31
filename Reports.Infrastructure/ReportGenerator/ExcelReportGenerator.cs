using ClosedXML.Excel;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using Reports.Infrastructure.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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

        public ExcelReportGenerator(IReportRepositoryAdoNet repositoryAdoNet, IReportRepositoryDapper repositoryDapper, ILogger logger)
        {
            this.repositoryAdoNet = repositoryAdoNet;
            this.repositoryDapper = repositoryDapper;
            this.logger = logger;
        }

        public async Task<byte[]> ExecuteAsync(ReportRequest request)
        {
            try
            {
                //Get global data from ReportsDtl table
                var reportDtl = await repositoryDapper.GetReportsDtlByReportIDAsync(request.ReportID);

                if (reportDtl == null)
                    throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);

                Manifest manifest = await repositoryDapper.GetManifestByManifestIDAsync(request.ManifestID);

                //For SP parameters
                request.Parameters["ManifestID"] = request.ManifestID;

                DataSet dataSet = repositoryAdoNet.GetData(request, reportDtl);

                if (dataSet != null && dataSet.Tables.Count > 0)
                {                    
                    return GenerateExcel(dataSet, request, reportDtl, manifest);
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
                            return GenerateInvBckReport(dataSet, reportDtl);
                        case Enums.GenerateExcel.GenerateSUMvalindex3Report:
                            return GenerateSUMvalindex3Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMdeliveryGush8Report:
                            return GenerateSUMdeliveryGush8Report(dataSet, request, reportDtl, manifest);

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

        private byte[] GenerateSUMentries9Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
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

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate))
                        {
                            filter.Append($"לתקופה: {toDate} - {fromDate}");
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
                            filter.Append($"ללקוח: {customsAgent} ({customsAgentID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"לתיק סוכן: {customsAgentFile}");
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
                    worksheet.Cell(currentRow, 1).FormulaA1 = $"COUNTA(A{startCalc}:A{currentRow - 1})";
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 6).FormulaA1 = $"SUM(F{startCalc}:F{currentRow - 1})";
                    worksheet.Cell(currentRow, 6).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 9).FormulaA1 = $"SUM(I{startCalc}:I{currentRow - 1})";
                    worksheet.Cell(currentRow, 9).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 11).FormulaA1 = $"SUM(K{startCalc}:K{currentRow - 1})";
                    worksheet.Cell(currentRow, 11).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 12).FormulaA1 = $"SUM(L{startCalc}:L{currentRow - 1})";
                    worksheet.Cell(currentRow, 12).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 13).FormulaA1 = $"SUM(M{startCalc}:M{currentRow - 1})";
                    worksheet.Cell(currentRow, 13).Style.Font.Bold = true;



                    currentRow += 3;

                    worksheet.Cell(currentRow, 7).Value = "ריכוז כניסות לפי מטבע";
                    worksheet.Cell(currentRow, 7).Style.Font.Underline = XLFontUnderlineValues.Single;
                    worksheet.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 7).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(currentRow, 7), worksheet.Cell(currentRow, 9)).Merge();

                    currentRow++;

                    var table2 = worksheet.Cell(currentRow, 7).InsertTable(dataSet.Tables[1]);
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

        private byte[] GenerateInvBckReport(DataSet dataSet, ReportDtl reportDtl)
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


                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 2;
                    worksheet.Cell(currentRow, 1).FormulaA1 = $"COUNTA(A{startCalc}:A{currentRow - 1})";
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 10).FormulaA1 = $"SUM(J{startCalc}:J{currentRow - 1})";
                    worksheet.Cell(currentRow, 10).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 12).FormulaA1 = $"SUM(L{startCalc}:L{currentRow - 1})";
                    worksheet.Cell(currentRow, 12).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 13).FormulaA1 = $"SUM(M{startCalc}:M{currentRow - 1})";
                    worksheet.Cell(currentRow, 13).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 14).FormulaA1 = $"SUM(N{startCalc}:N{currentRow - 1})";
                    worksheet.Cell(currentRow, 14).Style.Font.Bold = true;


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

                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                    worksheet.RightToLeft = true;

                    worksheet.Style.Font.FontSize = 10;
                   
                    worksheet.PageSetup.FitToPages(1, 0);
                    worksheet.PageSetup.PagesWide = 1;
                    worksheet.PageSetup.PagesTall = 0;
 
                    double marginSize = 0.5; 
                    worksheet.PageSetup.Margins.Left = marginSize;
                    worksheet.PageSetup.Margins.Right = marginSize;
                    worksheet.PageSetup.Margins.Top = marginSize;
                    worksheet.PageSetup.Margins.Bottom = marginSize;
                    
                    worksheet.PageSetup.CenterHorizontally = true;

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate))
                        {
                            filter.Append($"נכון לתאריך: {validityDate}");
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
                    worksheet.Cell(currentRow, 1).FormulaA1 = $"COUNTA(A{startCalc}:A{currentRow - 1})";
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 7).FormulaA1 = $"SUM(G{startCalc}:G{currentRow - 1})";
                    worksheet.Cell(currentRow, 7).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 8).FormulaA1 = $"SUM(H{startCalc}:H{currentRow - 1})";
                    worksheet.Cell(currentRow, 8).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 9).FormulaA1 = $"SUM(I{startCalc}:I{currentRow - 1})";
                    worksheet.Cell(currentRow, 9).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 10).FormulaA1 = $"SUM(J{startCalc}:J{currentRow - 1})";
                    worksheet.Cell(currentRow, 10).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 11).FormulaA1 = $"SUM(K{startCalc}:K{currentRow - 1})";
                    worksheet.Cell(currentRow, 11).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 12).FormulaA1 = $"SUM(L{startCalc}:L{currentRow - 1})";
                    worksheet.Cell(currentRow, 12).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 13).FormulaA1 = $"SUM(M{startCalc}:M{currentRow - 1})";
                    worksheet.Cell(currentRow, 13).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 14).FormulaA1 = $"SUM(N{startCalc}:N{currentRow - 1})";
                    worksheet.Cell(currentRow, 14).Style.Font.Bold = true;

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

                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                    worksheet.RightToLeft = true;

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate))
                        {
                            filter.Append($"לתקופה: {toDate} - {fromDate}");
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
                    worksheet.Cell(currentRow, 1).FormulaA1 = $"COUNTA(A{startCalc}:A{currentRow - 1})";
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, 9).FormulaA1 = $"SUM(I{startCalc}:I{currentRow - 1})";
                    worksheet.Cell(currentRow, 9).Style.Font.Bold = true;

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
    }
}
