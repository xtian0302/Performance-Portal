using HCL_HRIS.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
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
                prod.minsLogged = 540;
                prodCollection.Add(prod);
            } 
            reader.Close();
            command.Dispose();

            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = command.ExecuteReader();

            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0 , hours2 = 0;
                try{ 
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }catch(Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24) {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24){
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)-24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)-24).ToString("D2")); 
                }else if((hours2 + 16) > 24) {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                } else {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2")); 
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = command.ExecuteReader();

            List<Absents> leaves = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["shift"].ToString();
                abs.shift = shift; 
                leaves.Add(abs);
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
                DateTime day = Convert.ToDateTime(reader["Date"]);
                mins.minsDate = day;
                DateTime login = Convert.ToDateTime(reader["Login"]); 
                string shift = reader["Shift"].ToString();
                mins.shift = shift;
                try { 
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }catch(Exception e){
                    Debug.WriteLine(e.Message);
                }
                int hours1 = 0 , hours2 = 0;
                try{ 
                    hours1 = int.Parse(mins.shift.Substring(0, 2));
                    hours2 = int.Parse(mins.shift.Substring(6, 2));
                }catch(Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24) {
                    mins.shift = mins.shift.Replace(hours1.ToString("D2"), "00");
                    mins.shift = mins.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24){
                    mins.shift = mins.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)-24).ToString("D2"));
                    mins.shift = mins.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)-24).ToString("D2")); 
                }else if((hours2 + 16) > 24) {
                    mins.shift = mins.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    mins.shift = mins.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                } else { 
                    mins.shift = mins.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    mins.shift = mins.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2")); 
                }
                if (mins.shift.Equals("OFF")){
                    mins.shift = "";
                }
                if (reader.GetInt32(reader.GetOrdinal("TotalAudits")) != 0)  { 
                    try { 
                        mins.bcRate = string.Format("{0:0%}", 1 - (decimal.Parse(reader["BCFail"].ToString())/decimal.Parse(reader["TotalAudits"].ToString())));
                    }catch(Exception e) {
                        Debug.Print(e.Message);
                        mins.bcRate = "100%";
                    }
                    try {
                        mins.eucRate = string.Format("{0:0%}", 1 - (decimal.Parse(reader["EUCFail"].ToString())/decimal.Parse(reader["TotalAudits"].ToString())));
                    }catch(Exception e) {
                        Debug.Print(e.Message);
                        mins.eucRate = "100%";
                    }
                    try {
                        mins.ccRate = string.Format("{0:0%}", 1 - (decimal.Parse(reader["CCFail"].ToString())/decimal.Parse(reader["TotalAudits"].ToString())));
                    }catch(Exception e) {
                        Debug.Print(e.Message);
                        mins.ccRate = "100%";
                    }
                }else{
                    mins.bcRate = "NA";
                    mins.eucRate = "NA";
                    mins.ccRate = "NA";
                }
                try { 
                    prodCollection.Where(x => x.prodDate == mins.minsDate).First().minsLogged = mins.minsLogged;
                }catch(Exception){
                    Debug.WriteLine("Same prod date not found");
                }
                minsCollection.Add(mins);
            }
        
            reader.Close();
            command.Dispose();
            connection.Close();
            ViewBag.prodCollection = prodCollection;
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents; 
            ViewBag.leaves = leaves;
            return View();
        }
    }
    public class ProdPerDay
    {
        public int prodCount { get; set; }
        public DateTime prodDate { get; set; }
        public int minsLogged { get; set; }
    }
    public class MinsPerDay
    {
        public int minsLogged { get; set; }
        public DateTime minsDate { get; set; }
        public string bcRate { get; set; }
        public string eucRate { get; set; }
        public string ccRate { get; set; }
        public string shift { get; set; }
        public DateTime loginTime { get; set; }
        public double minsLate { get; set; } 
    }

    public class Absents
    {
        public DateTime day { get; set; }
        public string shift { get; set; } 
    }

}