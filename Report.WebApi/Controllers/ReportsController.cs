using Report.WebApi.Filters;
using Reports.Core.Configuration;
using Reports.Infrastructure.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

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
            if (request == null)
                return BadRequest("Invalid request.");

            try
            {
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
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }
    }
}
