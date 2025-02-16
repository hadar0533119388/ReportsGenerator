using Reports.Infrastructure.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Logger
{
    public class Logger : ILogger
    {
        private readonly string logDirectory;
        private string logFileName;
        public Logger(string logDirectory)
        {
            this.logDirectory = logDirectory;

            if (string.IsNullOrEmpty(logDirectory))
            {
                throw new Exception("LogPath is not found in the config file.");
            }

            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }
        }

        public void WriteLog(string message)
        {
            try
            {
                logFileName = Path.Combine(logDirectory, $"Reports_{DateTime.Now:yyyyMMdd}.log");
                using (StreamWriter writer = new StreamWriter(logFileName, true))
                {
                    writer.WriteLine($"{DateTime.Now:dd/MM/yyyy HH:mm:ss.fff} - {message}");
                }
            }
            catch (Exception ex)
            {
                throw new CustomException((int)ErrorMessages.ErrorCodes.LogAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.LogAccessFailure] } : {ex.Message}");
            }
        }
    }
}
