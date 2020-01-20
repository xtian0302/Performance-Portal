using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
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
    public static class Calculations
    {
        public static double getEQOverallProd(int aveprod, int complete, int otc)
        {
            return (aveprod * .3) + (complete * .3) + (otc * .4); 
        }
        public static double getOverallScoredProd(double aveprod, double complete, double otc)
        { 
            return (getEQScoreProd(aveprod) * .3) + (getEQScoreCmpltPrcnt(complete) * .3) + (getEQScoreOTC(otc) * .4);
        } 
        public static double getQAScoredProd(double bc, double euc, double cc)
        {
            return (getEQScoreBC(bc) * .3) + (getEQScoreEUC(euc) * .3) + (getEQScoreCC(cc) * .4);
        }
        public static double getCompScoredProd(double wpu, int lms)
        {
            return (getEQWpuScore(wpu) * .5) + (lms * .5) ;
        }
        public static int getEQScoreBC(double BcScore)  
        {
            double score = BcScore * 100;
            if (score == 0)
            {
                return 0;
            }
            else if (score == 100)
            {
                return 5;
            }
            else if (score >= 95)
            {
                return 4;
            }
            else if (score >= 90)
            {
                return 3;
            }
            else if (score >= 85)
            {
                return 2;
            }
            else if (score < 85)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQScoreEUC(double EucScore)
        {
            double score = EucScore * 100;
            if (score == 0)
            {
                return 0;
            }
            else if (score == 100)
            {
                return 5;
            }
            else if (score >= 98)
            {
                return 4;
            }
            else if (score >= 95)
            {
                return 3;
            }
            else if (score >= 90)
            {
                return 2;
            }
            else if (score < 90)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQScoreCC(double CcScore)
        {
            double score = CcScore * 100;
            if (score == 0)
            {
                return 0;
            }
            else if (score >= 99.5)
            {
                return 5;
            }
            else if (score < 99.5)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQScoreProd(double score)
        {
            if (score == 0)
            {
                return 0;   
            }
            else if (score > 30)
            {
                return 5;
            }
            else if (score >= 29)
            {
                return 4;
            }
            else if (score >= 26)
            {
                return 3;
            }
            else if (score >= 23)
            {
                return 2;
            }
            else if (score < 23)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQScoreCmpltPrcnt(double PcentScore)
        {
            double score = PcentScore * 100;
            if (score == 0)
            {
                return 0;
            }
            else if (score >71)
            {
                return 5;
            }
            else if (score >= 69)
            {
                return 4;
            }
            else if (score >= 65)
            {
                return 3;
            }
            else if (score >= 60)
            {
                return 2;
            }
            else if (score < 60)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQScoreOTC(double PcentScore)
        {
            double score = PcentScore * 100;
            if (score == 0)
            {
                return 0;
            }
            else if (score > 98)
            {
                return 5;
            }
            else if (score >= 97)
            {
                return 4;
            }
            else if (score >= 95)
            {
                return 3;
            } 
            else if (score < 95)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQWpuScore(double PcentScore)
        {
            double score = PcentScore * 100;
            //if (score == 0)
            //{
            //    return 0;
            //}
            //else 
            if (score == 100)
            {
                return 5;
            } 
            else if (score >= 75)
            {
                return 3;
            } 
            else if (score < 75)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getEQAbsScore(double PcentScore)
        {
            double score = PcentScore * 100;
            //if (score == 0)
            //{
            //    return 0;
            //}
            //else 
            if (score == 0)
            {
                return 5;
            }
            else if (score < 5)
            {
                return 4;
            }
            else if (score == 5)
            {
                return 3;
            }
            else if (score < 10)
            {
                return 2;
            }
            else if (score >= 10)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public static int getWeekOfMonth(DateTime date)
        {
            DateTime beginningOfMonth = new DateTime(date.Year, date.Month, 1);

            while (date.Date.AddDays(1).DayOfWeek != CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek)
                date = date.AddDays(1);

            return (int)Math.Truncate((double)date.Subtract(beginningOfMonth).TotalDays / 7f) + 1;
        }

    }
}