using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Models;
using Reports.Infrastructure.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
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

        public ExcelReportGenerator(IReportRepositoryAdoNet repositoryAdoNet)
        {
            this.repositoryAdoNet = repositoryAdoNet;
        }

        public async Task<byte[]> ExecuteAsync(ReportRequest request)
        {

            DataSet dataSet = await repositoryAdoNet.GetDataAsync(request.ReportID, request.Parameters);

            if (dataSet != null && dataSet.Tables.Count > 0)
            {
                DataTable dataTable = dataSet.Tables[0];
                object excelContent = GenerateExcel(dataTable); //consider to use NPOI
                byte[] excel = ExcelToBytes(excelContent);
            }
            else
            {
                //write to log
            }
            return null;
        }

        private object GenerateExcel(DataTable dataTable)
        {
            return null;
        }

        private byte[] ExcelToBytes(object excelContent)
        {
            return null;
        }
    }
}
