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
        public static SqlConnection getConn()
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

                string _sql = @"SELECT [sap_id] FROM [dbo].[users] " +
                       @"WHERE [sap_id] = @u AND [password] = @p";
                var cmd = new SqlCommand(_sql, cn);
                cmd.Parameters
                    .Add(new SqlParameter("@u", SqlDbType.Int))
                    .Value = Int32.Parse(_username);
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

        internal static bool IsValid(int sap_id, string password)
        {
            throw new NotImplementedException();
        }
    }
    public static class ExceptionHelper
    {
        private static Exception GetInnermostException(Exception ex)
        {
            while (ex.InnerException != null)
            {
                ex = ex.InnerException;
            }
            return ex;
        }

        public static bool IsUniqueConstraintViolation(Exception ex)
        {
            var innermost = GetInnermostException(ex);
            var sqlException = innermost as SqlException;

            return sqlException != null && sqlException.Class == 14 && (sqlException.Number == 2601 || sqlException.Number == 2627);
        }
    }
}