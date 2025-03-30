using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.Repositories
{
    public class ReportRepositoryAdoNet : IReportRepositoryAdoNet
    {
        private readonly string connectionString;
        private readonly ILogger logger;


        public ReportRepositoryAdoNet(string connectionString, ILogger logger)
        {
            this.connectionString = connectionString;
            this.logger = logger;
        }


        public Dictionary<string, object> GetFilteredParameters(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                HashSet<string> parametersToRemove = new HashSet<string>();

                if (Enum.TryParse(reportDtl.FunctionName, true, out StoredProcedure storedProcedure))
                {
                    switch (storedProcedure)
                    {
                        //case StoredProcedure.GetDataForSUMentries9Report:
                        //    parametersToRemove.Add("ManifestID");
                        //    break;

                        default:
                            break;
                    }
                }

                return request.Parameters.Where(p => !parametersToRemove.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value);
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Filtered Parameters For {reportDtl.ReportID} Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
        public DataSet GetData(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                Dictionary<string, object> parameters = GetFilteredParameters(request, reportDtl);

                using (SqlConnection conn = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand(reportDtl.FunctionName, conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            cmd.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
                        }
                    }

                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataSet ds = new DataSet();
                        adapter.Fill(ds);
                        return ds;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For {reportDtl.ReportID} Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }
    }
}
