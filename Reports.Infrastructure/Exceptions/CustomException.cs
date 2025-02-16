using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Exceptions
{
    public class CustomException : Exception
    {
        public int ErrorCode { get; }

        public CustomException(int errorCode, string message) : base(message)
        {
            ErrorCode = errorCode;
        }
    }
}
