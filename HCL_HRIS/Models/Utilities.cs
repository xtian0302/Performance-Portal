using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace HCL_HRIS.Models
{
    public static class Utilities
    {
        public static SqlConnection getDPLConn()
        {
            HCL_HRISEntities db = new HCL_HRISEntities();
            return new SqlConnection(db.Database.Connection.ConnectionString);
        }
        public static bool IsValid(string _username, string _password)
        {
            HCL_HRISEntities db = new HCL_HRISEntities();
            using (var cn = new SqlConnection(db.Database.Connection.ConnectionString))
            {
                if (String.IsNullOrWhiteSpace(_username)) _username = "";
                if (String.IsNullOrWhiteSpace(_password)) _password = "";

                string _sql = @"SELECT [sap_id] FROM [dbo].[User] " +
                       @"WHERE [sap_id] = @u AND [password] = @p";
                var cmd = new SqlCommand(_sql, cn);
                cmd.Parameters
                    .Add(new SqlParameter("@u", SqlDbType.NVarChar))
                    .Value = _username;
                cmd.Parameters
                    .Add(new SqlParameter("@p", SqlDbType.NVarChar))
                    .Value = _password;
                cn.Open();
                var reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Dispose();
                    cmd.Dispose();
                    return true;
                }
                else
                {
                    reader.Dispose();
                    cmd.Dispose();
                    return false;
                }
            }
        }
    }
}