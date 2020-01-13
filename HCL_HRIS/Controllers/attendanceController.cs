using HCL_HRIS.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HCL_HRIS.Controllers
{
    public class attendanceController : Controller
    {
        // GET: attendance
        public ActionResult Index()
        {
            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_ProdPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<ProdPerDay> prodCollection = new List<ProdPerDay>();
             while (reader.Read())
            {
                ProdPerDay prod = new ProdPerDay();
                prod.prodCount = (int.Parse(reader["prodCount"].ToString()));
                prod.prodDate = (Convert.ToDateTime(reader["Date"])); 
                prodCollection.Add(prod);
            } 
            reader.Close();
            command.Dispose();

            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = command.ExecuteReader();

            List<MinsPerDay> minsCollection = new List<MinsPerDay>();
            while (reader.Read())
            {
                MinsPerDay mins = new MinsPerDay();
                mins.minsLogged = (int.Parse(reader["Minutes"].ToString()));
                mins.minsDate = (Convert.ToDateTime(reader["Date"]));
                if (reader.GetInt32(reader.GetOrdinal("TotalAudits")) != 0)  { 
                    try { 
                        mins.bcRate = string.Format("{0:0%}", 1 - (decimal.Parse(reader["BCFail"].ToString())/decimal.Parse(reader["TotalAudits"].ToString())));
                    }catch(Exception e) {
                        mins.bcRate = "100%";
                    }
                    try {
                        mins.eucRate = string.Format("{0:0%}", 1 - (decimal.Parse(reader["EUCFail"].ToString())/decimal.Parse(reader["TotalAudits"].ToString())));
                    }catch(Exception e) {
                        mins.eucRate = "100%";
                    }
                    try {
                        mins.ccRate = string.Format("{0:0%}", 1 - (decimal.Parse(reader["CCFail"].ToString())/decimal.Parse(reader["TotalAudits"].ToString())));
                    }catch(Exception e) {
                        mins.ccRate = "100%";
                    }
                }else{
                    mins.bcRate = "NA";
                    mins.eucRate = "NA";
                    mins.ccRate = "NA"; 
                }
                minsCollection.Add(mins);
            }
        
            reader.Close();
            command.Dispose();
            connection.Close();
            ViewBag.prodCollection = prodCollection;
            ViewBag.minsCollection = minsCollection;
            return View();
        }
    }
    public class ProdPerDay
    {
        public int prodCount { get; set; }
        public DateTime prodDate { get; set; } 
    }
    public class MinsPerDay
    {
        public int minsLogged { get; set; }
        public DateTime minsDate { get; set; }
        public string bcRate { get; set; }
        public string eucRate { get; set; }
        public string ccRate { get; set; }

    }
}