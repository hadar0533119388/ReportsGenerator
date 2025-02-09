using Dapper;
using Reports.Infrastructure.DTOs;
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
    public class ReportRepositoryDapper : IReportRepositoryDapper
    {
        private readonly string connectionString;
        private readonly ILogger logger;


        public ReportRepositoryDapper(string connectionString, ILogger logger)
        {
            this.connectionString = connectionString;
            this.logger = logger;
        }

        public async Task<dynamic> GetDataAsync(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                Manifest manifest = await GetManifestByManifestIDAsync(request.ManifestID);

                if (Enum.TryParse(reportDtl.FunctionName, true, out StoredProcedure storedProcedure))
                {
                    switch (storedProcedure)
                    {
                        case StoredProcedure.GetDataForR912470Report:
                            R912470ReportResponse R912470ReportResponse = await GetDataForR912470Report(request.Parameters, manifest);
                            return R912470ReportResponse;

                        default:
                            break;
                    }

                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data Async: {ex.Message}");
            }
            return null;
        }

        public async Task<Manifest> GetManifestByManifestIDAsync(string manifestID)
        {
            try
            {               
                using (var conn = new SqlConnection(connectionString))
                {
                    var parameters = new { ManifestID = manifestID };

                    return await conn.QueryFirstOrDefaultAsync<Manifest>(
                        StoredProcedure.GetManifestByManifestID.ToString(),
                        parameters,
                        commandType: CommandType.StoredProcedure);                    
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Manifest By ManifestID Async: {ex.Message}");
                return null;
            }
        }
        public async Task<ReportDtl> GetReportsDtlByReportIDAsync(string reportID)
        {
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    var parameters = new { ReportID = reportID };

                    return await conn.QueryFirstOrDefaultAsync<ReportDtl>(
                        StoredProcedure.GetReportsDtl.ToString(),
                        parameters,
                        commandType: CommandType.StoredProcedure);
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get ReportsDtl By ReportID Async: {ex.Message}");
                return null;
            }
        }
        public async Task<R912470ReportResponse> GetDataForR912470Report(Dictionary<string, object> parameters, Manifest manifest)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR912470Report.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {                        
                        var response = new R912470ReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            ItemsList = (await multi.ReadAsync<Item>()).ToList(),
                            Control1050List = (await multi.ReadAsync<Control1050>()).ToList(),
                            Manifest = manifest
                        };
                        
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R912470 Report: {ex.Message}");
                return null;
            }
        }

    }
}
