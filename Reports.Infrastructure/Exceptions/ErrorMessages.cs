using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Exceptions
{
    public class ErrorMessages
    {
        public enum ErrorCodes
        {
            GlobalError = -1,
            InvalidInput = 401,
            DBAccessFailure = 402,
            LogAccessFailure = 403,
            NoDataFound = 404,
            UnknownPrinter = 405,
            FailedToPrint = 406

        }
        public static readonly Dictionary<int, string> Messages = new Dictionary<int, string>
    {
        { 401, "Invalid Input" },
        { 402, "DB access failure" },
        { 403, "Log access failure" },
        { 404, "No Data found" },
        { 405, "Unknown Printer" },
        { 406, "Failed to print" }
    };
        public static readonly Dictionary<int, HttpStatusCode> StatusCodes = new Dictionary<int, HttpStatusCode>
    {
        { -1, HttpStatusCode.InternalServerError },
        { 401, HttpStatusCode.BadRequest },
        { 402, HttpStatusCode.InternalServerError },
        { 403, HttpStatusCode.InternalServerError },
        { 404, HttpStatusCode.NotFound },
        { 405, HttpStatusCode.InternalServerError },
        { 406, HttpStatusCode.InternalServerError }
    };
    }
}
