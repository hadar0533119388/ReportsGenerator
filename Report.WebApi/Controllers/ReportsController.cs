using Report.WebApi.Filters;
using Reports.Core.Configuration;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using static Reports.Infrastructure.Models.Enums;

namespace Report.WebApi.Controllers
{
    [IPRestrictionFilter]
    [RoutePrefix("api/reports")]
    public class ReportsController : ApiController
    {
        [HttpPost]
        [Route("generate")]
        public async Task<IHttpActionResult> GenerateReport([FromBody] ReportRequest request)
        {
            try
            {
                if (!ModelState.IsValid || request == null)
                    throw new CustomException((int)ErrorMessages.ErrorCodes.InvalidInput, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.InvalidInput]);

                ValidateReportParameters(request);

                byte[] reportBytes = await ReportGeneratorFacade.GenerateReportAsync(request);

                if (reportBytes == null || reportBytes.Length == 0)
                    return NotFound();

                return ResponseMessage(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(reportBytes)
                    {
                        Headers = { ContentType = new MediaTypeHeaderValue("application/octet-stream") }
                    }
                });
            }
            catch (CustomException ex)
            {
                var httpStatus = ErrorMessages.StatusCodes.ContainsKey(ex.ErrorCode)
                                ? ErrorMessages.StatusCodes[ex.ErrorCode]
                                : HttpStatusCode.InternalServerError; 

                return Content(httpStatus, new ErrorResponse
                {
                    ErrorCode = ex.ErrorCode,
                    ErrorMessage = ex.Message
                });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }
        private void ValidateReportParameters(ReportRequest request)
        {
            var requiredParams = GetRequiredParametersForReport(request.ReportID);
            var missingParams = requiredParams.Where(p => !request.Parameters.ContainsKey(p)).ToList();

            if (missingParams.Any())
            {
                throw new ArgumentException($"Missing required parameters: {string.Join(", ", missingParams)}");
            }
        }

        private List<string> GetRequiredParametersForReport(string reportID)
        {
            var reportRequirements = new Dictionary<string, List<string>>
        {
            { ReportID.SUMentries9.ToString(), new List<string> { "FromDate", "ToDate" } }
        };

            return reportRequirements.TryGetValue(reportID, out var requiredParams) ? requiredParams : new List<string>();
        }
    }
}
