using System.Web.Http;
using System.Net;
using System.Net.Http;
using System.Web.Http.Filters;
using System.Web.Http.Results;
using System.Web.Http.Controllers;
using System.Web;

namespace Report.WebApi.Filters
{
    public class IPRestrictionFilter : ActionFilterAttribute
    {
        public override void OnActionExecuting(HttpActionContext context)
        {
            var remoteIp = HttpContext.Current.Request.UserHostAddress;
            if (remoteIp != "127.0.0.1" && remoteIp != "::1") // IPv4 и IPv6
            {
                context.Response = context.Request.CreateResponse(HttpStatusCode.Forbidden, "Access denied");
            }     
            base.OnActionExecuting(context);
        }
    }
}