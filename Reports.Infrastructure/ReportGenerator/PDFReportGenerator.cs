using DinkToPdf;
using PuppeteerSharp;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using Reports.Infrastructure.Repositories;
using Stubble.Core.Builders;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.ReportGenerator
{
    public class PDFReportGenerator : IReportGenerator
    {
        public Enums.ReportType Type => Enums.ReportType.PDF;
        private readonly IReportRepositoryDapper repositoryDapper;
        private readonly ILogger logger;
        private readonly string sumatraPDF;
        private readonly string chrome;


        public PDFReportGenerator(IReportRepositoryDapper repositoryDapper, ILogger logger, string sumatraPDF, string chrome)
        {
            this.repositoryDapper = repositoryDapper;
            this.logger = logger;
            this.sumatraPDF = sumatraPDF;
            this.chrome = chrome;
        }

        public async Task<byte[]> ExecuteAsync(ReportRequest request)
        {
            try
            {
                //Get global data from ReportsDtl table
                var reportDtl = await repositoryDapper.GetReportsDtlByReportIDAsync(request.ReportID);

                //Get data for a specific report
                var dataReport = await repositoryDapper.GetDataAsync(request, reportDtl);

                //Convert data to dictionary
                Dictionary<string, object> data = FlattenObject(dataReport);


                if (data != null && data.Any())
                {
                    string htmlContent = GenerateHtml(data, reportDtl.Template);
                    byte[] pdf = await HtmlToPdfDocument(htmlContent);

                    if (request.IsPrint)
                    {
                        PrintPdfDocument(pdf, request.PrinterName);
                    }

                    return pdf;
                }
                else
                {
                    logger.WriteLog($"No data was returned for the report.");
                    return null;
                }
                
            }
            catch(Exception ex)
            {
                logger.WriteLog($"Error to Execute Async: {ex.Message}");
                return null;
            }
        }


        private Dictionary<string, object> FlattenObject(object obj, string prefix = "")
        {
            try
            {
                if (obj == null) return null;

                var result = new Dictionary<string, object>();

                foreach (var prop in obj.GetType().GetProperties())
                {
                    var value = prop.GetValue(obj);

                    if (value == null)
                    {
                        result[prefix + prop.Name] = null;
                    }
                    else if (IsSimpleType(value.GetType()))
                    {
                        if (value is DateTime dt)
                        {
                            result[prefix + prop.Name] = dt.ToString("dd/MM/yyyy");
                        }
                        else
                        {
                            result[prefix + prop.Name] = value;
                        }
                    }
                    else if (value is IEnumerable list && value.GetType() != typeof(string))
                    {
                        // If the field is a list (including lists of objects)
                        var listData = new List<Dictionary<string, object>>();
                        foreach (var item in list)
                        {
                            listData.Add(FlattenObject(item, prefix + prop.Name + "_"));
                        }
                        result[prefix + prop.Name] = listData;
                    }
                    else
                    {
                        // If it's another internal object (not a list) - flatten the object inside
                        var flattenedObject = FlattenObject(value, prefix + prop.Name + "_");
                        foreach (var subProp in flattenedObject)
                        {
                            result[subProp.Key] = subProp.Value;
                        }
                    }
                }

                return result;
            }
            catch(Exception ex)
            {
                logger.WriteLog($"Error to Flatten Object to dictionary: {ex.Message}");
                return null;
            }
        }

        private bool IsSimpleType(Type type)
        {
            return type.IsPrimitive || type.IsValueType || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime);
        }

        private string GenerateHtml(Dictionary<string, object> data, string templateReport)
        {
            try
            {
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "Templates", templateReport);
                string template = File.ReadAllText(templatePath);
                
                var stubble = new StubbleBuilder().Build();
                string html = stubble.Render(template, data);

                // שמירת קובץ HTML עם הנתונים
                //File.WriteAllText("output.html", html);

                logger.WriteLog($"Generate HTML completed successfully.");

                return html;
            }
            catch(Exception ex)
            {
                logger.WriteLog($"Error to generate HTML: {ex.Message}");
                return null;
            }
        }

        private  async Task<byte[]> HtmlToPdfDocument(string htmlContent)
        {
            try
            {
                var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                {
                    Headless = true,
                    ExecutablePath = chrome
                });
                var page = await browser.NewPageAsync();

                await page.SetContentAsync(htmlContent);

                var pdfBuffer = await page.PdfDataAsync();

                await browser.CloseAsync();

                logger.WriteLog("Generate PDF from HTML completed successfully.");
                
                return pdfBuffer;

            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to generate PDF from HTML: {ex.Message}");
                return null;
            }
        }


        private void PrintPdfDocument(byte[] pdf, string printerName)
        {
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");

            try
            {
                // Saving the array to a temporary file
                File.WriteAllBytes(tempFilePath, pdf);

                // Creating parameters for running SumatraPDF for printing
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = sumatraPDF,
                    Arguments = $"-print-to \"{printerName}\" {tempFilePath}",
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = false
                };


                using (Process process = new Process { StartInfo = psi })
                {
                    process.Start();

                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        logger.WriteLog($"Error to Print PDF Document. ExitCode: {process.ExitCode}");
                        if (!string.IsNullOrEmpty(error))
                        {
                            logger.WriteLog($"Error details: {error}");
                        }
                    }
                    else
                    {
                        logger.WriteLog($"Print PDF Document completed successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Print PDF Document: {ex.Message}");
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
    }
}
