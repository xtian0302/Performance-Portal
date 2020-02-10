using ClosedXML.Excel;
using HCL_HRIS.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
namespace HCL_HRIS.Controllers
{
    public class HomeController : Controller
    { 
        private HCL_HRISEntities db = new HCL_HRISEntities();
        public async Task<ActionResult> Index()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated) {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            //check if user is agent. if team lead redirect to tl page
            if(usr.user_role.Equals("Team Leader")) {
                return RedirectToAction("TeamLead", "Home"); 
            }
            if (!usr.user_role.Equals("Administrator"))
            { 
                if (usr.sub_department.Equals("PPMC") || usr.sub_department.Equals("PPMC IB/BPM")){
                    return RedirectToAction("PPMCIndex", "Home");
                } else if (usr.sub_department.Equals("PPMC IB L2")){
                    return RedirectToAction("PPMCL2Index", "Home"); 
                } else if (usr.sub_department.Equals("Kaiser Closet")){
                    return RedirectToAction("KaiserClosetIndex", "Home"); 
                } else if (usr.sub_department.Equals("Kaiser SMC Resupply")){
                    return RedirectToAction("KaiserSMCIndex", "Home");
                } 
                else if (usr.sub_department.Equals("Kaiser BU/ AH")|| usr.sub_department.Equals("Kaiser BU/AH") || usr.sub_department.Equals("Kaiser Pickup"))
                {
                    return RedirectToAction("KaiserOthersIndex", "Home");
                }
                else if (usr.sub_department.Equals("PPMC BPM"))
                {
                    return RedirectToAction("PPMCBPMIndex", "Home");
                }

            }
            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn(); 
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0; 
            SqlCommand command = new SqlCommand("Select top 1 * from top5 order by date desc", connection);
            connection.Open();
            SqlDataReader reader =await command.ExecuteReaderAsync();
            if(reader.HasRows){
                while (reader.Read()){
                    ViewBag.top1sap = reader["top1_sap"];
                    ViewBag.top1name = reader["top1_name"];
                    ViewBag.top2sap = reader["top2_sap"];
                    ViewBag.top2name = reader["top2_name"];
                    ViewBag.top3sap = reader["top3_sap"];
                    ViewBag.top3name = reader["top3_name"];
                    ViewBag.top4sap = reader["top4_sap"];
                    ViewBag.top4name = reader["top4_name"];
                    ViewBag.top5sap = reader["top5_sap"];
                    ViewBag.top5name = reader["top5_name"];
                }
            }else{
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap =2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            }
            ViewBag.Top5 = await Services.DataAccess.query_Top5Async(connection);
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select rank, (Select Count(*) from rankings where track = 'Sleep EQ') as count from (Select RANK() OVER(ORDER BY score DESC) as rank,sap_id as sapno from rankings where track = 'Sleep EQ') tb where sapno = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name); 
            reader =await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"]; 
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            } 

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name; 
            reader =await command.ExecuteReaderAsync();  
                while (reader.Read()){
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks"))) {
                    ViewBag.wpu = 0.0;
                } else { 
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }  

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;  
            reader =await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            {   aveprod = tryGetData(reader, "AveProd");
                if (Double.Parse(reader["Completes"].ToString()) == 0 || Double.Parse(reader["Concludes"].ToString()) == 0){
                    cmplt = 0;
                } else {
                    cmplt = Double.Parse(reader["Completes"].ToString()) / Double.Parse(reader["Concludes"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLA")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLA"))){
                    otc = Double.Parse(reader["WithinSLA"].ToString()) / (Double.Parse(reader["WithinSLA"].ToString()) + Double.Parse(reader["NotWithinSLA"].ToString()));
                } else {
                    otc = 0;
                }
                try { 
                ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }catch(Exception e) {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString()); 
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose(); 

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0){
               ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0-((double)bcErrorCurrMos/(double)totalAuditCurrMos), 1.0-((double)eucErrorCurrMos/(double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            } else {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(0, 0, 0);
            }
            ViewBag.ProdScore = HCL_HRIS.Models.Calculations.getOverallScoredProd(aveprod,cmplt,otc);
            ViewBag.OverallScore = string.Format("{0:0.##}", (ViewBag.ProdScore * 0.45) + (ViewBag.QAScore * 0.3) + (ViewBag.lms*0.05) + (ViewBag.WpuScore * 0.05) + (ViewBag.AbsScore * 0.15));
            ViewBag.ProdScore = string.Format("{0:0.#}", ViewBag.ProdScore);
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);
           
            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name; 
            reader =await command.ExecuteReaderAsync(); 
            List<MinsPerDay> minsCollection = new List<MinsPerDay>();
            while (reader.Read()){
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
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            } 
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader =await command.ExecuteReaderAsync(); 
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name; 
            reader =await command.ExecuteReaderAsync(); 
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
            connection.Close(); 
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> PPMCIndex()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated)
            {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            //check if user is agent. if team lead redirect to tl page
            if (usr.user_role.Equals("Team Leader"))
            {
                return RedirectToAction("TeamLead", "Home");
            }

            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            SqlCommand command = new SqlCommand("Select (Select name from users  where sap_id = sapno) as name, sapno from (Select top 5 sap_id as sapno,rank from ppmcl1_overall order by rank)tb", connection);
            connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            int i = 1;
            if (reader.HasRows)  {
                while (reader.Read()) {
                    if(i == 1) {
                        ViewBag.top1sap = reader["sapno"];
                        ViewBag.top1name = reader["name"];
                    }else if(i == 2) { 
                        ViewBag.top2sap = reader["sapno"];
                        ViewBag.top2name = reader["name"];
                    }else if(i == 3) { 
                        ViewBag.top3sap = reader["sapno"];
                        ViewBag.top3name = reader["name"];
                    }else if(i == 4) { 
                        ViewBag.top4sap = reader["sapno"];
                        ViewBag.top4name = reader["name"];
                    }else if(i == 5) { 
                        ViewBag.top5sap = reader["sapno"];
                        ViewBag.top5name = reader["name"];
                    } 
                    i++;
                }
            }
            else
            {
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap = 2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            } 
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select ave_calls_handled_score, aht_score, cash_col_score, eom_score, rank as rank, (select max(rank) from ppmcl1_overall) as count from ppmcl1_overall where sap_id = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {  
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"]; 
                    ViewBag.OverallScore = reader["eom_score"];
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_calls_handled_score")) * .4) + (reader.GetInt32(reader.GetOrdinal("aht_score")) * .3) + (reader.GetInt32(reader.GetOrdinal("cash_col_score")) * .3));

                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            { 
                try
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }
                catch (Exception e)
                {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose();

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0)
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1, 1, 1);
            }
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);

            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
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
                try
                {
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = await command.ExecuteReaderAsync();
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0, hours2 = 0;
                try
                {
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16) - 24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours2 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2"));
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            connection.Close();
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> TeamLead()
        {
            SqlConnection connection = Utilities.getConn();
            SqlCommand command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();
            List<MinsPerDay> minsCollection = new List<MinsPerDay>();
            while (reader.Read())
            {
                MinsPerDay mins = new MinsPerDay();
                mins.minsLogged = (int.Parse(reader["Minutes"].ToString()));
                mins.minsDate = (Convert.ToDateTime(reader["Date"]));
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose(); 

            command = new SqlCommand("get_TeamRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.scoreRank = reader["scoreRank"];
                ViewBag.prodRank = reader["prodRank"];
                ViewBag.qualityRank = reader["qualityRank"];
                ViewBag.behaviorRank = reader["behaviorRank"];
                ViewBag.complianceRank = reader["complianceRank"];
                ViewBag.total = reader["total"];
            }
            reader.Close();
            command.Dispose();
            connection.Close();

            command = new SqlCommand("Select top 1 * from top5 order by date desc", connection);    
            connection.Open();
            reader = command.ExecuteReader();
            if(reader.HasRows){
                while (reader.Read()){
                    ViewBag.top1sap = reader["top1_sap"];
                    ViewBag.top1name = reader["top1_name"];
                    ViewBag.top2sap = reader["top2_sap"];
                    ViewBag.top2name = reader["top2_name"];
                    ViewBag.top3sap = reader["top3_sap"];
                    ViewBag.top3name = reader["top3_name"];
                    ViewBag.top4sap = reader["top4_sap"];
                    ViewBag.top4name = reader["top4_name"];
                    ViewBag.top5sap = reader["top5_sap"];
                    ViewBag.top5name = reader["top5_name"];
                }
            }else{
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap =2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            } 
            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows){
                while (reader.Read()){ 
                ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString()))); 
                }
            }else{
                ViewBag.lms = 0; 
            } 
            command = new SqlCommand("Select top 1 * from group_scores where leader_sap = @sap_id order by group_scores_id desc", connection);
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows){
                while (reader.Read()){
                    ViewBag.group_score = reader["group_score"];
                    ViewBag.group_prod = reader["group_prod"];
                    ViewBag.group_quality = reader["group_quality"];
                    ViewBag.group_behavior = reader["group_behavior"];
                    ViewBag.group_compliance = reader["group_compliance"];  
                }
            }else
            {
                ViewBag.group_score = 0;
                ViewBag.group_prod = 0;
                ViewBag.group_quality = 0;
                ViewBag.group_behavior = 0;
                ViewBag.group_compliance = 0;
            }
            connection.Close();
            ViewBag.minsCollection = minsCollection;

            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.manager = usr.group.track.user.name; 
            ViewBag.leader_mail = leader.nt_login + "@hcl.com"; 
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> KaiserClosetIndex()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated)
            {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr; 

            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            SqlCommand command = new SqlCommand("Select (Select name from users  where sap_id = sapno) as name, sapno from (Select top 5 sap_id as sapno,rank from kaiser_closet_overall order by rank)tb", connection);
            connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            int i = 1;
            if (reader.HasRows)  {
                while (reader.Read()) {
                    if(i == 1) {
                        ViewBag.top1sap = reader["sapno"];
                        ViewBag.top1name = reader["name"];
                    }else if(i == 2) { 
                        ViewBag.top2sap = reader["sapno"];
                        ViewBag.top2name = reader["name"];
                    }else if(i == 3) { 
                        ViewBag.top3sap = reader["sapno"];
                        ViewBag.top3name = reader["name"];
                    }else if(i == 4) { 
                        ViewBag.top4sap = reader["sapno"];
                        ViewBag.top4name = reader["name"];
                    }else if(i == 5) { 
                        ViewBag.top5sap = reader["sapno"];
                        ViewBag.top5name = reader["name"];
                    } 
                    i++;
                }
            }
            else
            {
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap = 2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            } 
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select ave_prod_score, otc_score, eom_score, rank as rank, (select max(rank) from kaiser_closet_overall) as count from kaiser_closet_overall where sap_id = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {  
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"]; 
                    ViewBag.OverallScore = reader["eom_score"];
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_prod_score")) * .4) + (reader.GetInt32(reader.GetOrdinal("otc_score")) * .6));

                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            { 
                try
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }
                catch (Exception e)
                {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose();

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0)
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1, 1, 1);
            }
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);

            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
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
                try
                {
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = await command.ExecuteReaderAsync();
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0, hours2 = 0;
                try
                {
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16) - 24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours2 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2"));
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            connection.Close();
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        public async Task<ActionResult> KaiserSMCIndex()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated)
            {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            SqlCommand command = new SqlCommand("Select (Select name from users  where sap_id = sapno) as name, sapno from (Select top 5 sap_id as sapno,rank from kaiser_smc_overall order by rank)tb", connection);
            connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            int i = 1;
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    if (i == 1)
                    {
                        ViewBag.top1sap = reader["sapno"];
                        ViewBag.top1name = reader["name"];
                    }
                    else if (i == 2)
                    {
                        ViewBag.top2sap = reader["sapno"];
                        ViewBag.top2name = reader["name"];
                    }
                    else if (i == 3)
                    {
                        ViewBag.top3sap = reader["sapno"];
                        ViewBag.top3name = reader["name"];
                    }
                    else if (i == 4)
                    {
                        ViewBag.top4sap = reader["sapno"];
                        ViewBag.top4name = reader["name"];
                    }
                    else if (i == 5)
                    {
                        ViewBag.top5sap = reader["sapno"];
                        ViewBag.top5name = reader["name"];
                    }
                    i++;
                }
            }
            else
            {
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap = 2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            }
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select ave_prod_score, eom_score, rank as rank, (select max(rank) from kaiser_smc_overall) as count from kaiser_smc_overall where sap_id = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"];
                    ViewBag.OverallScore = reader["eom_score"];
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_prod_score"))) );

                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            {
                try
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }
                catch (Exception e)
                {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose();

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0)
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1, 1, 1);
            }
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);

            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
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
                try
                {
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = await command.ExecuteReaderAsync();
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0, hours2 = 0;
                try
                {
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16) - 24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours2 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2"));
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            connection.Close();
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> KaiserOthersIndex()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated)
            {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            SqlCommand command = new SqlCommand("Select (Select name from users  where sap_id = sapno) as name, sapno from (Select top 5 sap_id as sapno,rank from kaiser_others_overall order by rank)tb", connection);
            connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            int i = 1;
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    if (i == 1)
                    {
                        ViewBag.top1sap = reader["sapno"];
                        ViewBag.top1name = reader["name"];
                    }
                    else if (i == 2)
                    {
                        ViewBag.top2sap = reader["sapno"];
                        ViewBag.top2name = reader["name"];
                    }
                    else if (i == 3)
                    {
                        ViewBag.top3sap = reader["sapno"];
                        ViewBag.top3name = reader["name"];
                    }
                    else if (i == 4)
                    {
                        ViewBag.top4sap = reader["sapno"];
                        ViewBag.top4name = reader["name"];
                    }
                    else if (i == 5)
                    {
                        ViewBag.top5sap = reader["sapno"];
                        ViewBag.top5name = reader["name"];
                    }
                    i++;
                }
            }
            else
            {
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap = 2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            }
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select ave_prod_score, eom_score, rank as rank, (select max(rank) from kaiser_others_overall) as count from kaiser_others_overall where sap_id = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"];
                    ViewBag.OverallScore = reader["eom_score"];
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_prod_score"))));

                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            {
                try
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }
                catch (Exception e)
                {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose();

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0)
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1, 1, 1);
            }
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);

            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
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
                try
                {
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = await command.ExecuteReaderAsync();
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0, hours2 = 0;
                try
                {
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16) - 24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours2 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2"));
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            connection.Close();
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> PPMCBPMIndex()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated)
            {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            SqlCommand command = new SqlCommand("Select (Select name from users  where sap_id = sapno) as name, sapno from (Select top 5 sap_id as sapno,rank from ppmc_bpm_overall order by rank)tb", connection);
            connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            int i = 1;
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    if (i == 1)
                    {
                        ViewBag.top1sap = reader["sapno"];
                        ViewBag.top1name = reader["name"];
                    }
                    else if (i == 2)
                    {
                        ViewBag.top2sap = reader["sapno"];
                        ViewBag.top2name = reader["name"];
                    }
                    else if (i == 3)
                    {
                        ViewBag.top3sap = reader["sapno"];
                        ViewBag.top3name = reader["name"];
                    }
                    else if (i == 4)
                    {
                        ViewBag.top4sap = reader["sapno"];
                        ViewBag.top4name = reader["name"];
                    }
                    else if (i == 5)
                    {
                        ViewBag.top5sap = reader["sapno"];
                        ViewBag.top5name = reader["name"];
                    }
                    i++;
                }
            }
            else
            {
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap = 2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            }
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select bpm_score,otc_score, eom_score, rank as rank, (select max(rank) from ppmc_bpm_overall) as count from ppmc_bpm_overall where sap_id = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"];
                    ViewBag.OverallScore = reader["eom_score"];
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("bpm_score"))*.5)+ (reader.GetInt32(reader.GetOrdinal("otc_score")) * .5) );

                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            {
                try
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }
                catch (Exception e)
                {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose();

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0)
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1, 1, 1);
            }
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);

            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
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
                try
                {
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = await command.ExecuteReaderAsync();
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0, hours2 = 0;
                try
                {
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16) - 24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours2 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2"));
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            connection.Close();
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        public async Task<ActionResult> PPMCL2Index()
        {

            //Check if user is logged in else return to login page
            if (!Request.IsAuthenticated)
            {
                return RedirectToAction("Login", "Users");
            }
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            //check if user is agent. if team lead redirect to tl page
            if (usr.user_role.Equals("Team Leader"))
            {
                return RedirectToAction("TeamLead", "Home");
            }

            //queries start here
            //Get top 5 agents of track
            SqlConnection connection = Utilities.getConn();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            SqlCommand command = new SqlCommand("Select (Select name from users  where sap_id = sapno) as name, sapno from (Select top 5 sap_id as sapno,rank from ppmcl2_overall order by rank)tb", connection);
            connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            int i = 1;
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    if (i == 1)
                    {
                        ViewBag.top1sap = reader["sapno"];
                        ViewBag.top1name = reader["name"];
                    }
                    else if (i == 2)
                    {
                        ViewBag.top2sap = reader["sapno"];
                        ViewBag.top2name = reader["name"];
                    }
                    else if (i == 3)
                    {
                        ViewBag.top3sap = reader["sapno"];
                        ViewBag.top3name = reader["name"];
                    }
                    else if (i == 4)
                    {
                        ViewBag.top4sap = reader["sapno"];
                        ViewBag.top4name = reader["name"];
                    }
                    else if (i == 5)
                    {
                        ViewBag.top5sap = reader["sapno"];
                        ViewBag.top5name = reader["name"];
                    }
                    i++;
                }
            }
            else
            {
                ViewBag.top1sap = 1;
                ViewBag.top1name = "Agent1";
                ViewBag.top2sap = 2;
                ViewBag.top2name = "Agent2";
                ViewBag.top3sap = 3;
                ViewBag.top3name = "Agent3";
                ViewBag.top4sap = 4;
                ViewBag.top4name = "Agent4";
                ViewBag.top5sap = 5;
                ViewBag.top5name = "Agent5";
            }
            //Get ranking of this agent against other agents
            command = new SqlCommand("Select ave_calls_handled_score, aht_score, cash_col_score, eom_score, rank as rank, (select max(rank) from ppmcl2_overall) as count from ppmcl2_overall where sap_id = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.myRank = reader["rank"];
                    ViewBag.outOf = reader["count"];
                    ViewBag.OverallScore = reader["eom_score"];
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_calls_handled_score")) * .4) + (reader.GetInt32(reader.GetOrdinal("aht_score")) * .3) + (reader.GetInt32(reader.GetOrdinal("cash_col_score")) * .3));

                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }

            //get WPU Scores
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                ViewBag.WpuScore = HCL_HRIS.Models.Calculations.getEQWpuScore(ViewBag.wpu);
            }

            //Get Prod, Quality, Absenteeism and LMS Scores
            command = new SqlCommand("get_Prodfast", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            {
                try
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                }
                catch (Exception e)
                {
                    ViewBag.lms = 0;
                }
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.AbsScore = HCL_HRIS.Models.Calculations.getEQAbsScore(ViewBag.absCurr);
            }
            reader.Close();
            command.Dispose();

            //Calculate Scores for Viewing
            if (totalAuditCurrMos != 0)
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1, 1, 1);
            }
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);

            //get Attendance Calendar Data
            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
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
                try
                {
                    mins.minsLate = login.Subtract(day.AddHours(double.Parse(shift.Substring(0, 2))).AddMinutes(double.Parse(shift.Substring(3, 2)))).TotalMinutes;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(" String in Question :" + shift);
                }
                minsCollection.Add(mins);
            }
            reader.Close();
            command.Dispose();

            //Get Leaves for attendance calendar data
            command = new SqlCommand("get_Leaves", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            reader = await command.ExecuteReaderAsync();
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

            //Get Absents for attendance calendar data
            command = new SqlCommand("get_Absents", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            reader = await command.ExecuteReaderAsync();
            List<Absents> absents = new List<Absents>();
            while (reader.Read())
            {
                Absents abs = new Absents();
                abs.day = reader.GetDateTime(reader.GetOrdinal("day"));
                string shift = reader["Shift"].ToString();
                abs.shift = shift;
                int hours1 = 0, hours2 = 0;
                try
                {
                    hours1 = int.Parse(abs.shift.Substring(0, 2));
                    hours2 = int.Parse(abs.shift.Substring(6, 2));
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                if ((hours1 + 16) == 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), "00");
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours1 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16) - 24).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else if ((hours2 + 16) > 24)
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16) - 24).ToString("D2"));
                }
                else
                {
                    abs.shift = abs.shift.Replace(hours1.ToString("D2"), ((hours1 + 16)).ToString("D2"));
                    abs.shift = abs.shift.Replace(hours2.ToString("D2"), ((hours2 + 16)).ToString("D2"));
                }
                absents.Add(abs);
            }
            reader.Close();
            command.Dispose();
            connection.Close();
            //Close connection
            //End of queries

            //Return to View Calendar Collection
            ViewBag.minsCollection = minsCollection;
            ViewBag.absents = absents;
            ViewBag.leaves = leaves;

            //Return to View User information
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_id = leader.user_id;
            ViewBag.leader_name = leader.name.Trim();
            ViewBag.leader_mail = leader.nt_login + "@hcl.com";
            ViewBag.manager = usr.group.track.user.name;
            ViewBag.manager_mail = usr.group.track.user.nt_login + "@hcl.com";
            ViewBag.user_role = usr.user_role;

            //Return to view list of announcements
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
         
        public ActionResult Import()
        {
            return View();
        }

        public async Task<ActionResult> OldIndex()
        {
            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();
            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();
            command = new SqlCommand("get_Prod", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                //  Previous Month Productivity
                ViewBag.AveProdPrev = int.Parse(reader["AveProdPrev"].ToString());
                if (Double.Parse(reader["CompletesPrev"].ToString()) == 0 || Double.Parse(reader["ConcludesPrev"].ToString()) == 0)
                {
                    ViewBag.CompletePercentPrev = 0;
                }
                else
                {
                    ViewBag.CompletePercentPrev = Double.Parse(reader["CompletesPrev"].ToString()) / Double.Parse(reader["ConcludesPrev"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAPrev")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAPrev")))
                {
                    ViewBag.OTCPrev = Double.Parse(reader["WithinSLAPrev"].ToString()) / (Double.Parse(reader["WithinSLAPrev"].ToString()) + Double.Parse(reader["NotWithinSLAPrev"].ToString()));
                }
                else
                {
                    ViewBag.OTCPrev = 0;
                }
                //  Current Month Productivity
                ViewBag.AveProd = int.Parse(reader["AveProd"].ToString());
                if (Double.Parse(reader["Completes"].ToString()) == 0 || Double.Parse(reader["Concludes"].ToString()) == 0)
                {
                    ViewBag.CompletePercent = 0;
                }
                else
                {
                    ViewBag.CompletePercent = Double.Parse(reader["Completes"].ToString()) / Double.Parse(reader["Concludes"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLA")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLA")))
                {
                    ViewBag.OTC = Double.Parse(reader["WithinSLA"].ToString()) / (Double.Parse(reader["WithinSLA"].ToString()) + Double.Parse(reader["NotWithinSLA"].ToString()));
                }
                else
                {
                    ViewBag.OTC = 0;
                }
                //  Week1 Productivity
                ViewBag.AveProdWk1 = int.Parse(reader["AveProdWk1"].ToString());
                if (Double.Parse(reader["CompletesWk1"].ToString()) == 0 || Double.Parse(reader["ConcludesWk1"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk1 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk1 = Double.Parse(reader["CompletesWk1"].ToString()) / Double.Parse(reader["ConcludesWk1"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk1")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk1")))
                {
                    ViewBag.OTCWk1 = Double.Parse(reader["WithinSLAWk1"].ToString()) / (Double.Parse(reader["WithinSLAWk1"].ToString()) + Double.Parse(reader["NotWithinSLAWk1"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk1 = 0;
                }

                //  Week2 Productivity
                ViewBag.AveProdWk2 = int.Parse(reader["AveProdWk2"].ToString());
                if (Double.Parse(reader["CompletesWk2"].ToString()) == 0 || Double.Parse(reader["ConcludesWk2"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk2 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk2 = Double.Parse(reader["CompletesWk2"].ToString()) / Double.Parse(reader["ConcludesWk2"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk2")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk2")))
                {
                    ViewBag.OTCWk2 = Double.Parse(reader["WithinSLAWk2"].ToString()) / (Double.Parse(reader["WithinSLAWk2"].ToString()) + Double.Parse(reader["NotWithinSLAWk2"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk2 = 0;
                }

                //  Week3 Productivity
                ViewBag.AveProdWk3 = int.Parse(reader["AveProdWk3"].ToString());
                if (Double.Parse(reader["CompletesWk3"].ToString()) == 0 || Double.Parse(reader["ConcludesWk3"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk3 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk3 = Double.Parse(reader["CompletesWk3"].ToString()) / Double.Parse(reader["ConcludesWk3"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk3")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk3")))
                {
                    ViewBag.OTCWk3 = Double.Parse(reader["WithinSLAWk3"].ToString()) / (Double.Parse(reader["WithinSLAWk3"].ToString()) + Double.Parse(reader["NotWithinSLAWk3"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk3 = 0;
                }

                //  Week4 Productivity
                ViewBag.AveProdWk4 = int.Parse(reader["AveProdWk4"].ToString());
                if (Double.Parse(reader["CompletesWk4"].ToString()) == 0 || Double.Parse(reader["ConcludesWk4"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk4 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk4 = Double.Parse(reader["CompletesWk4"].ToString()) / Double.Parse(reader["ConcludesWk4"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk4")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk4")))
                {
                    ViewBag.OTCWk4 = Double.Parse(reader["WithinSLAWk4"].ToString()) / (Double.Parse(reader["WithinSLAWk4"].ToString()) + Double.Parse(reader["NotWithinSLAWk4"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk4 = 0;
                }

                //  Week5 Productivity
                ViewBag.AveProdWk5 = int.Parse(reader["AveProdWk5"].ToString());
                if (Double.Parse(reader["CompletesWk5"].ToString()) == 0 || Double.Parse(reader["ConcludesWk5"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk5 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk5 = Double.Parse(reader["CompletesWk5"].ToString()) / Double.Parse(reader["ConcludesWk5"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk5")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk5")))
                {
                    ViewBag.OTCWk5 = Double.Parse(reader["WithinSLAWk5"].ToString()) / (Double.Parse(reader["WithinSLAWk5"].ToString()) + Double.Parse(reader["NotWithinSLAWk5"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk5 = 0;
                }
            }
            connection.Close();
            command = new SqlCommand("get_ProdRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.AveProdPrevRank = int.Parse(reader["AveProdPrevRank"].ToString());
                ViewBag.CompletePercentPrevRank = int.Parse(reader["CompletePercentPrev"].ToString());
                ViewBag.OTCPrevRank = int.Parse(reader["OTCPrev"].ToString());
                ViewBag.AveProdRank = int.Parse(reader["AveProdRank"].ToString());
                ViewBag.CompletePercentRank = int.Parse(reader["CompletePercent"].ToString());
                ViewBag.OTCRank = int.Parse(reader["OTC"].ToString());
                ViewBag.AveProdWk1Rank = int.Parse(reader["AveProdWk1Rank"].ToString());
                ViewBag.CompletePercentWk1Rank = int.Parse(reader["CompletePercentWk1"].ToString());
                ViewBag.OTCWk1Rank = int.Parse(reader["OTCWk1"].ToString());
                ViewBag.AveProdWk2Rank = int.Parse(reader["AveProdWk2Rank"].ToString());
                ViewBag.CompletePercentWk2Rank = int.Parse(reader["CompletePercentWk2"].ToString());
                ViewBag.OTCWk2Rank = int.Parse(reader["OTCWk2"].ToString());
                ViewBag.AveProdWk3Rank = int.Parse(reader["AveProdWk3Rank"].ToString());
                ViewBag.CompletePercentWk3Rank = int.Parse(reader["CompletePercentWk3"].ToString());
                ViewBag.OTCWk3Rank = int.Parse(reader["OTCWk3"].ToString());
                ViewBag.AveProdWk4Rank = int.Parse(reader["AveProdWk4Rank"].ToString());
                ViewBag.CompletePercentWk4Rank = int.Parse(reader["CompletePercentWk4"].ToString());
                ViewBag.OTCWk4Rank = int.Parse(reader["OTCWk4"].ToString());
                ViewBag.AveProdWk5Rank = int.Parse(reader["AveProdWk5Rank"].ToString());
                ViewBag.CompletePercentWk5Rank = int.Parse(reader["CompletePercentWk5"].ToString());
                ViewBag.OTCWk5Rank = int.Parse(reader["OTCWk5"].ToString());
            }
            connection.Close();

            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public ActionResult Chat()
        {
            int sap_id = Int32.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            List<user> agentsList = db.users.Where(x => x.group.user.user_id == usr.user_id).ToList<user>();
            ViewBag.agentsList = agentsList;
            ViewBag.user = usr;
            return View();
        }
        public async Task<ActionResult> Details()
        {  
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            if (!usr.user_role.Equals("Administrator"))
            {
                if (usr.sub_department.Equals("PPMC") || usr.sub_department.Equals("PPMC IB/BPM"))
                {
                    return RedirectToAction("PPMCDetails", "Home");
                } else if (usr.sub_department.Equals("PPMC IB L2")){
                    return RedirectToAction("PPMCL2Details", "Home"); 
                } else if (usr.sub_department.Equals("Kaiser Closet")){
                    return RedirectToAction("KaiserClosetDetails", "Home"); 
                } else if (usr.sub_department.Equals("Kaiser SMC Resupply")) {
                    return RedirectToAction("KaiserSMCDetails", "Home");
                }
                else if (usr.sub_department.Equals("Kaiser BU/ AH") || usr.sub_department.Equals("Kaiser BU/AH") || usr.sub_department.Equals("Kaiser Pickup"))
                {
                    return RedirectToAction("KaiserOthersDetails", "Home");
                }
                else if (usr.sub_department.Equals("PPMC BPM"))
                {
                    return RedirectToAction("PPMCBPMDetails", "Home");
                }
            }
            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();
            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();

            command = new SqlCommand("get_Prod", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                //  Previous Month Productivity
                ViewBag.AveProdPrev = tryGetData(reader, "AveProdPrev");
                if (Double.Parse(reader["CompletesPrev"].ToString()) == 0 || Double.Parse(reader["ConcludesPrev"].ToString()) == 0)
                {
                    ViewBag.CompletePercentPrev = 0;
                }
                else
                {
                    ViewBag.CompletePercentPrev = Double.Parse(reader["CompletesPrev"].ToString()) / Double.Parse(reader["ConcludesPrev"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAPrev")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAPrev")))
                {
                    ViewBag.OTCPrev = Double.Parse(reader["WithinSLAPrev"].ToString()) / (Double.Parse(reader["WithinSLAPrev"].ToString()) + Double.Parse(reader["NotWithinSLAPrev"].ToString()));
                }
                else
                {
                    ViewBag.OTCPrev = 0;
                }
                //  Current Month Preaderroductivity
                ViewBag.AveProd = tryGetData(reader,"AveProd");
                if (Double.Parse(reader["Completes"].ToString()) == 0 || Double.Parse(reader["Concludes"].ToString()) == 0)
                {
                    ViewBag.CompletePercent = 0;
                }
                else
                {
                    ViewBag.CompletePercent = Double.Parse(reader["Completes"].ToString()) / Double.Parse(reader["Concludes"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLA")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLA")))
                {
                    ViewBag.OTC = Double.Parse(reader["WithinSLA"].ToString()) / (Double.Parse(reader["WithinSLA"].ToString()) + Double.Parse(reader["NotWithinSLA"].ToString()));
                }
                else
                {
                    ViewBag.OTC = 0;
                }
                //  Week1 Productivity
                ViewBag.AveProdWk1 = int.Parse(reader["AveProdWk1"].ToString());
                if (Double.Parse(reader["CompletesWk1"].ToString()) == 0 || Double.Parse(reader["ConcludesWk1"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk1 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk1 = Double.Parse(reader["CompletesWk1"].ToString()) / Double.Parse(reader["ConcludesWk1"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk1")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk1")))
                {
                    ViewBag.OTCWk1 = Double.Parse(reader["WithinSLAWk1"].ToString()) / (Double.Parse(reader["WithinSLAWk1"].ToString()) + Double.Parse(reader["NotWithinSLAWk1"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk1 = 0;
                }

                //  Week2 Productivity
                ViewBag.AveProdWk2 = int.Parse(reader["AveProdWk2"].ToString());
                if (Double.Parse(reader["CompletesWk2"].ToString()) == 0 || Double.Parse(reader["ConcludesWk2"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk2 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk2 = Double.Parse(reader["CompletesWk2"].ToString()) / Double.Parse(reader["ConcludesWk2"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk2")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk2")))
                {
                    ViewBag.OTCWk2 = Double.Parse(reader["WithinSLAWk2"].ToString()) / (Double.Parse(reader["WithinSLAWk2"].ToString()) + Double.Parse(reader["NotWithinSLAWk2"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk2 = 0;
                }

                //  Week3 Productivity
                ViewBag.AveProdWk3 = int.Parse(reader["AveProdWk3"].ToString());
                if (Double.Parse(reader["CompletesWk3"].ToString()) == 0 || Double.Parse(reader["ConcludesWk3"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk3 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk3 = Double.Parse(reader["CompletesWk3"].ToString()) / Double.Parse(reader["ConcludesWk3"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk3")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk3")))
                {
                    ViewBag.OTCWk3 = Double.Parse(reader["WithinSLAWk3"].ToString()) / (Double.Parse(reader["WithinSLAWk3"].ToString()) + Double.Parse(reader["NotWithinSLAWk3"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk3 = 0;
                }

                //  Week4 Productivity
                ViewBag.AveProdWk4 = int.Parse(reader["AveProdWk4"].ToString());
                if (Double.Parse(reader["CompletesWk4"].ToString()) == 0 || Double.Parse(reader["ConcludesWk4"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk4 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk4 = Double.Parse(reader["CompletesWk4"].ToString()) / Double.Parse(reader["ConcludesWk4"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk4")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk4")))
                {
                    ViewBag.OTCWk4 = Double.Parse(reader["WithinSLAWk4"].ToString()) / (Double.Parse(reader["WithinSLAWk4"].ToString()) + Double.Parse(reader["NotWithinSLAWk4"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk4 = 0;
                }

                //  Week5 Productivity
                ViewBag.AveProdWk5 = int.Parse(reader["AveProdWk5"].ToString());
                if (Double.Parse(reader["CompletesWk5"].ToString()) == 0 || Double.Parse(reader["ConcludesWk5"].ToString()) == 0)
                {
                    ViewBag.CompletePercentWk5 = 0;
                }
                else
                {
                    ViewBag.CompletePercentWk5 = Double.Parse(reader["CompletesWk5"].ToString()) / Double.Parse(reader["ConcludesWk5"].ToString());
                }
                if (!reader.IsDBNull(reader.GetOrdinal("WithinSLAWk5")) || !reader.IsDBNull(reader.GetOrdinal("NotWithinSLAWk5")))
                {
                    ViewBag.OTCWk5 = Double.Parse(reader["WithinSLAWk5"].ToString()) / (Double.Parse(reader["WithinSLAWk5"].ToString()) + Double.Parse(reader["NotWithinSLAWk5"].ToString()));
                }
                else
                {
                    ViewBag.OTCWk5 = 0;
                }
            }
            connection.Close();
            command = new SqlCommand("get_ProdRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            command.CommandTimeout = 0;
            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.AveProdPrevRank = int.Parse(reader["AveProdPrevRank"].ToString());
                ViewBag.CompletePercentPrevRank = int.Parse(reader["CompletePercentPrev"].ToString());
                ViewBag.OTCPrevRank = int.Parse(reader["OTCPrev"].ToString());
                ViewBag.AveProdRank = int.Parse(reader["AveProdRank"].ToString());
                ViewBag.CompletePercentRank = int.Parse(reader["CompletePercent"].ToString());
                ViewBag.OTCRank = int.Parse(reader["OTC"].ToString());
                ViewBag.AveProdWk1Rank = int.Parse(reader["AveProdWk1Rank"].ToString());
                ViewBag.CompletePercentWk1Rank = int.Parse(reader["CompletePercentWk1"].ToString());
                ViewBag.OTCWk1Rank = int.Parse(reader["OTCWk1"].ToString());
                ViewBag.AveProdWk2Rank = int.Parse(reader["AveProdWk2Rank"].ToString());
                ViewBag.CompletePercentWk2Rank = int.Parse(reader["CompletePercentWk2"].ToString());
                ViewBag.OTCWk2Rank = int.Parse(reader["OTCWk2"].ToString());
                ViewBag.AveProdWk3Rank = int.Parse(reader["AveProdWk3Rank"].ToString());
                ViewBag.CompletePercentWk3Rank = int.Parse(reader["CompletePercentWk3"].ToString());
                ViewBag.OTCWk3Rank = int.Parse(reader["OTCWk3"].ToString());
                ViewBag.AveProdWk4Rank = int.Parse(reader["AveProdWk4Rank"].ToString());
                ViewBag.CompletePercentWk4Rank = int.Parse(reader["CompletePercentWk4"].ToString());
                ViewBag.OTCWk4Rank = int.Parse(reader["OTCWk4"].ToString());
                ViewBag.AveProdWk5Rank = int.Parse(reader["AveProdWk5Rank"].ToString());
                ViewBag.CompletePercentWk5Rank = int.Parse(reader["CompletePercentWk5"].ToString());
                ViewBag.OTCWk5Rank = int.Parse(reader["OTCWk5"].ToString());
            }
            connection.Close();
            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader(); 
            if (reader.HasRows) {
                while (reader.Read())
                { 
                        ViewBag.lms = double.Parse(string.Format("{0:0.#}",decimal.Parse(reader["LmsScore"].ToString()))); 
                        ViewBag.lmspercent = reader["LmsPercent"];
                }
            } else {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();  
                while (reader.Read()){
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks"))) {
                    ViewBag.wpu = 0.0;
                } else { 
                    ViewBag.wpu = reader["monthmarks"];
                } 
                if (reader.IsDBNull(reader.GetOrdinal("week1marks"))) {
                    ViewBag.wpuwk1 = 0.0;
                } else { 
                    ViewBag.wpuwk1 = reader["week1marks"];
                } 
                if (reader.IsDBNull(reader.GetOrdinal("week2marks"))) {
                    ViewBag.wpuwk2 = 0.0;
                } else { 
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks"))) {
                    ViewBag.wpuwk3 = 0.0;
                } else { 
                    ViewBag.wpuwk3 = reader["week3marks"];
                } 
                if (reader.IsDBNull(reader.GetOrdinal("week4marks"))) {
                    ViewBag.wpuwk4 = 0.0;
                } else { 
                    ViewBag.wpuwk4 = reader["week4marks"];
                } 
                if (reader.IsDBNull(reader.GetOrdinal("week5marks"))) {
                    ViewBag.wpuwk5 = 0.0;
                } else { 
                    ViewBag.wpuwk5 = reader["week5marks"];
                } 
            } 
            connection.Close();
             
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        public async Task<ActionResult> PPMCL2Details()
        {
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();

            command = new SqlCommand("Select ave_calls_handled,ave_calls_handled_score,aht, aht_score, cash_col, cash_col_score, eom_score, rank as rank, (select max(rank) from ppmcl2_overall) as count from ppmcl2_overall where sap_id = @1", connection);
            connection.Open();
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_calls_handled_score")) * .4) + (reader.GetInt32(reader.GetOrdinal("aht_score")) * .3) + (reader.GetInt32(reader.GetOrdinal("cash_col_score")) * .3));
                    ViewBag.AveCallsScore = reader.GetInt32(reader.GetOrdinal("ave_calls_handled_score"));
                    ViewBag.AveCalls = reader.GetDecimal(reader.GetOrdinal("ave_calls_handled"));
                    ViewBag.AHTScore = reader.GetInt32(reader.GetOrdinal("aht_score"));
                    ViewBag.AHT = reader.GetDecimal(reader.GetOrdinal("aht"));
                    ViewBag.CashColScore = reader.GetInt32(reader.GetOrdinal("cash_col_score"));
                    ViewBag.CashCol = string.Format("{0:P0}", reader.GetDecimal(reader.GetOrdinal("cash_col")));
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }
            connection.Close();

            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();
            
            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                    ViewBag.lmspercent = reader["LmsPercent"];
                }
            }
            else
            {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week1marks")))
                {
                    ViewBag.wpuwk1 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk1 = reader["week1marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week2marks")))
                {
                    ViewBag.wpuwk2 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks")))
                {
                    ViewBag.wpuwk3 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk3 = reader["week3marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week4marks")))
                {
                    ViewBag.wpuwk4 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk4 = reader["week4marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week5marks")))
                {
                    ViewBag.wpuwk5 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk5 = reader["week5marks"];
                }
            }
            connection.Close();

            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> KaiserClosetDetails()
        {
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();

            command = new SqlCommand("Select ave_prod,ave_prod_score,otc, otc_score, eom_score, rank as rank, (select max(rank) from kaiser_closet_overall) as count from kaiser_closet_overall where sap_id = @1", connection);
            connection.Open();
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_prod_score")) * .4) + (reader.GetInt32(reader.GetOrdinal("otc_score")) * .6));
                    ViewBag.AveProdScore = reader.GetInt32(reader.GetOrdinal("ave_prod_score"));
                    ViewBag.AveProd = reader.GetDecimal(reader.GetOrdinal("ave_prod")); 
                    ViewBag.OtcScore = reader.GetInt32(reader.GetOrdinal("otc_score"));
                    ViewBag.Otc = string.Format("{0:P0}", reader.GetDecimal(reader.GetOrdinal("otc")));
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }
            connection.Close();

            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();

            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                    ViewBag.lmspercent = reader["LmsPercent"];
                }
            }
            else
            {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week1marks")))
                {
                    ViewBag.wpuwk1 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk1 = reader["week1marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week2marks")))
                {
                    ViewBag.wpuwk2 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks")))
                {
                    ViewBag.wpuwk3 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk3 = reader["week3marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week4marks")))
                {
                    ViewBag.wpuwk4 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk4 = reader["week4marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week5marks")))
                {
                    ViewBag.wpuwk5 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk5 = reader["week5marks"];
                }
            }
            connection.Close();

            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }

        public async Task<ActionResult> PPMCBPMDetails()
        {
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();

            command = new SqlCommand("Select bpm, bpm_score,otc, otc_score, eom_score, rank as rank, (select max(rank) from ppmc_bpm_overall) as count from ppmc_bpm_overall where sap_id = @1", connection);
            connection.Open();
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("bpm_score")) * .5) + (reader.GetInt32(reader.GetOrdinal("otc_score")) * .5));
                    ViewBag.BPMScore = reader.GetInt32(reader.GetOrdinal("bpm_score"));
                    ViewBag.BPMProd = reader.GetDecimal(reader.GetOrdinal("bpm"));
                    ViewBag.OtcScore = reader.GetInt32(reader.GetOrdinal("otc_score"));
                    ViewBag.Otc = string.Format("{0:P0}", reader.GetDecimal(reader.GetOrdinal("otc")));
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }
            connection.Close();

            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();

            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                    ViewBag.lmspercent = reader["LmsPercent"];
                }
            }
            else
            {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week1marks")))
                {
                    ViewBag.wpuwk1 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk1 = reader["week1marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week2marks")))
                {
                    ViewBag.wpuwk2 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks")))
                {
                    ViewBag.wpuwk3 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk3 = reader["week3marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week4marks")))
                {
                    ViewBag.wpuwk4 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk4 = reader["week4marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week5marks")))
                {
                    ViewBag.wpuwk5 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk5 = reader["week5marks"];
                }
            }
            connection.Close();

            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        public async Task<ActionResult> KaiserSMCDetails()
        {
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();

            command = new SqlCommand("Select ave_prod,ave_prod_score, eom_score, rank as rank, (select max(rank) from kaiser_smc_overall) as count from kaiser_smc_overall where sap_id = @1", connection);
            connection.Open();
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_prod_score"))));
                    ViewBag.AveProdScore = reader.GetInt32(reader.GetOrdinal("ave_prod_score"));
                    ViewBag.AveProd = reader.GetDecimal(reader.GetOrdinal("ave_prod"));  
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }
            connection.Close();

            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();

            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                    ViewBag.lmspercent = reader["LmsPercent"];
                }
            }
            else
            {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week1marks")))
                {
                    ViewBag.wpuwk1 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk1 = reader["week1marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week2marks")))
                {
                    ViewBag.wpuwk2 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks")))
                {
                    ViewBag.wpuwk3 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk3 = reader["week3marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week4marks")))
                {
                    ViewBag.wpuwk4 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk4 = reader["week4marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week5marks")))
                {
                    ViewBag.wpuwk5 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk5 = reader["week5marks"];
                }
            }
            connection.Close();

            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        public async Task<ActionResult> KaiserOthersDetails()
        {
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();

            command = new SqlCommand("Select ave_prod,ave_prod_score, eom_score, rank as rank, (select max(rank) from kaiser_others_overall) as count from kaiser_others_overall where sap_id = @1", connection);
            connection.Open();
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_prod_score"))));
                    ViewBag.AveProdScore = reader.GetInt32(reader.GetOrdinal("ave_prod_score"));
                    ViewBag.AveProd = reader.GetDecimal(reader.GetOrdinal("ave_prod"));  
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }
            connection.Close();

            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();

            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                    ViewBag.lmspercent = reader["LmsPercent"];
                }
            }
            else
            {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week1marks")))
                {
                    ViewBag.wpuwk1 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk1 = reader["week1marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week2marks")))
                {
                    ViewBag.wpuwk2 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks")))
                {
                    ViewBag.wpuwk3 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk3 = reader["week3marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week4marks")))
                {
                    ViewBag.wpuwk4 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk4 = reader["week4marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week5marks")))
                {
                    ViewBag.wpuwk5 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk5 = reader["week5marks"];
                }
            }
            connection.Close();

            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        public async Task<ActionResult> PPMCDetails()
        {
            // get user identity for queries
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;

            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double bcErrorPrevMos = 0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
            while (reader.Read())
            {
                bcErrorPrevMos = int.Parse(reader["bcErrorPrevMos"].ToString());
                eucErrorPrevMos = int.Parse(reader["eucErrorPrevMos"].ToString());
                ccErrorPrevMos = int.Parse(reader["ccErrorPrevMos"].ToString());
                totalAuditPrevMos = int.Parse(reader["totalAuditPrevMos"].ToString());
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString());
                bcErrorWk1 = int.Parse(reader["bcErrorWk1"].ToString());
                eucErrorWk1 = int.Parse(reader["eucErrorWk1"].ToString());
                ccErrorWk1 = int.Parse(reader["ccErrorWk1"].ToString());
                totalAuditWk1 = int.Parse(reader["totalAuditWk1"].ToString());
                bcErrorWk2 = int.Parse(reader["bcErrorWk2"].ToString());
                eucErrorWk2 = int.Parse(reader["eucErrorWk2"].ToString());
                ccErrorWk2 = int.Parse(reader["ccErrorWk2"].ToString());
                totalAuditWk2 = int.Parse(reader["totalAuditWk2"].ToString());
                bcErrorWk3 = int.Parse(reader["bcErrorWk3"].ToString());
                eucErrorWk3 = int.Parse(reader["eucErrorWk3"].ToString());
                ccErrorWk3 = int.Parse(reader["ccErrorWk3"].ToString());
                totalAuditWk3 = int.Parse(reader["totalAuditWk3"].ToString());
                bcErrorWk4 = int.Parse(reader["bcErrorWk4"].ToString());
                eucErrorWk4 = int.Parse(reader["eucErrorWk4"].ToString());
                ccErrorWk4 = int.Parse(reader["ccErrorWk4"].ToString());
                totalAuditWk4 = int.Parse(reader["totalAuditWk4"].ToString());
                bcErrorWk5 = int.Parse(reader["bcErrorWk5"].ToString());
                eucErrorWk5 = int.Parse(reader["eucErrorWk5"].ToString());
                ccErrorWk5 = int.Parse(reader["ccErrorWk5"].ToString());
                totalAuditWk5 = int.Parse(reader["totalAuditWk5"].ToString());
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1}
            if (totalAuditPrevMos != 0)
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{0:0.00}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{0:0.00}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else
            {
                ViewBag.BCPrev = string.Format("{0:0.00}", 0);
                ViewBag.EUCPrev = string.Format("{0:0.00}", 0);
                ViewBag.CCPrev = string.Format("{0:0.00}", 0);
            }
            if (totalAuditCurrMos != 0)
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{0:0.00}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{0:0.00}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else
            {
                ViewBag.BCCurr = string.Format("{0:0.00}", 0);
                ViewBag.EUCCurr = string.Format("{0:0.00}", 0);
                ViewBag.CCCurr = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk1 != 0)
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else
            {
                ViewBag.BCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk1 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk1 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk2 != 0)
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else
            {
                ViewBag.BCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk2 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk2 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk3 != 0)
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else
            {
                ViewBag.BCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk3 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk3 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk4 != 0)
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else
            {
                ViewBag.BCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk4 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk4 = string.Format("{0:0.00}", 0);
            }
            if (totalAuditWk5 != 0)
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{0:0.00}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else
            {
                ViewBag.BCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.EUCWk5 = string.Format("{0:0.00}", 0);
                ViewBag.CCWk5 = string.Format("{0:0.00}", 0);
            }

            connection.Close();

            command = new SqlCommand("Select ave_calls_handled,ave_calls_handled_score,aht, aht_score, cash_col, cash_col_score, eom_score, rank as rank, (select max(rank) from ppmcl1_overall) as count from ppmcl1_overall where sap_id = @1", connection);
            connection.Open();
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = Int32.Parse(User.Identity.Name);
            reader = await command.ExecuteReaderAsync();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.ProdScore = string.Format("{0:0.#}", (reader.GetInt32(reader.GetOrdinal("ave_calls_handled_score")) * .4) + (reader.GetInt32(reader.GetOrdinal("aht_score")) * .3) + (reader.GetInt32(reader.GetOrdinal("cash_col_score")) * .3));
                    ViewBag.AveCallsScore = reader.GetInt32(reader.GetOrdinal("ave_calls_handled_score"));
                    ViewBag.AveCalls = reader.GetDecimal(reader.GetOrdinal("ave_calls_handled"));
                    ViewBag.AHTScore = reader.GetInt32(reader.GetOrdinal("aht_score"));
                    ViewBag.AHT = reader.GetDecimal(reader.GetOrdinal("aht"));
                    ViewBag.CashColScore = reader.GetInt32(reader.GetOrdinal("cash_col_score"));
                    ViewBag.CashCol = string.Format("{0:P0}", reader.GetDecimal(reader.GetOrdinal("cash_col")));
                }
            }
            else
            {
                ViewBag.myRank = 0;
                ViewBag.outOf = 0;
            }
            connection.Close();

            command = new SqlCommand("get_QualityRanks", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sapid", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.bcPrevRank = int.Parse(reader["bc_prev_rank"].ToString());
                ViewBag.bcCurrRank = int.Parse(reader["bc_curr_rank"].ToString());
                ViewBag.bcWk1Rank = int.Parse(reader["bc_wk1_rank"].ToString());
                ViewBag.bcWk2Rank = int.Parse(reader["bc_wk2_rank"].ToString());
                ViewBag.bcWk3Rank = int.Parse(reader["bc_wk3_rank"].ToString());
                ViewBag.bcWk4Rank = int.Parse(reader["bc_wk4_rank"].ToString());
                ViewBag.bcWk5Rank = int.Parse(reader["bc_wk5_rank"].ToString());
                ViewBag.eucPrevRank = int.Parse(reader["euc_prev_rank"].ToString());
                ViewBag.eucCurrRank = int.Parse(reader["euc_curr_rank"].ToString());
                ViewBag.eucWk1Rank = int.Parse(reader["euc_wk1_rank"].ToString());
                ViewBag.eucWk2Rank = int.Parse(reader["euc_wk2_rank"].ToString());
                ViewBag.eucWk3Rank = int.Parse(reader["euc_wk3_rank"].ToString());
                ViewBag.eucWk4Rank = int.Parse(reader["euc_wk4_rank"].ToString());
                ViewBag.eucWk5Rank = int.Parse(reader["euc_wk5_rank"].ToString());
                ViewBag.ccPrevRank = int.Parse(reader["cc_prev_rank"].ToString());
                ViewBag.ccCurrRank = int.Parse(reader["cc_curr_rank"].ToString());
                ViewBag.ccWk1Rank = int.Parse(reader["cc_wk1_rank"].ToString());
                ViewBag.ccWk2Rank = int.Parse(reader["cc_wk2_rank"].ToString());
                ViewBag.ccWk3Rank = int.Parse(reader["cc_wk3_rank"].ToString());
                ViewBag.ccWk4Rank = int.Parse(reader["cc_wk4_rank"].ToString());
                ViewBag.ccWk5Rank = int.Parse(reader["cc_wk5_rank"].ToString());
            }
            connection.Close();

            command = new SqlCommand("get_Absenteeism", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                ViewBag.absPrev = tryGetData(reader, "AbsenteeismPrev");
                ViewBag.absCurr = tryGetData(reader, "AbsenteeismCurr");
                ViewBag.absWk1 = tryGetData(reader, "AbsenteeismWk1");
                ViewBag.absWk2 = tryGetData(reader, "AbsenteeismWk2");
                ViewBag.absWk3 = tryGetData(reader, "AbsenteeismWk3");
                ViewBag.absWk4 = tryGetData(reader, "AbsenteeismWk4");
                ViewBag.absWk5 = tryGetData(reader, "AbsenteeismWk5");
            }
            connection.Close();
            
            command = new SqlCommand("get_Comp", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
                    ViewBag.lmspercent = reader["LmsPercent"];
                }
            }
            else
            {
                ViewBag.lms = 0;
                ViewBag.lmspercent = "0%";
            }
            connection.Close();
            command = new SqlCommand("get_WPU", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;

            connection.Open();
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                ViewBag.wpuprev = reader["prevmonth"];
                if (reader.IsDBNull(reader.GetOrdinal("monthmarks")))
                {
                    ViewBag.wpu = 0.0;
                }
                else
                {
                    ViewBag.wpu = reader["monthmarks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week1marks")))
                {
                    ViewBag.wpuwk1 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk1 = reader["week1marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week2marks")))
                {
                    ViewBag.wpuwk2 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk2 = reader["week2marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week3marks")))
                {
                    ViewBag.wpuwk3 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk3 = reader["week3marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week4marks")))
                {
                    ViewBag.wpuwk4 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk4 = reader["week4marks"];
                }
                if (reader.IsDBNull(reader.GetOrdinal("week5marks")))
                {
                    ViewBag.wpuwk5 = 0.0;
                }
                else
                {
                    ViewBag.wpuwk5 = reader["week5marks"];
                }
            }
            connection.Close();

            ViewBag.name = usr.name.Trim();
            ViewBag.user = usr;
            user leader = db.users.Where(x => x.user_id == usr.group.group_leader).First();
            ViewBag.leader_sap = leader.sap_id;
            ViewBag.leader_name = leader.name.Trim();
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        }
        [HttpPost]
        public async Task<JsonResult> EqUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].eq_prod " + //42 columns excl. ID
                        "([queue_name], [Process_Instance_ID], [Work_Item_ID], [Intake_CPU], [BU_ID], [Cust_ID], [Requeued_Queue_Name], [Task_Create_Date_Time], [Process_Type], [Task_Status], [Task_Resolution], [Idle_Seconds], [Handling_Seconds], [Duration_Seconds], [Touch_Count], [Last_Associate_ID], [Last_Associate_Name], [Last_Associcate_Manager], [Last_Associate_Site], [First_Assignment_Timestamp], [Last_Assignment_Timestamp], [Last_Assignment_Date], [Task_Create_Date], [Task_Resolved_Timestamp], [Last_Follow_Up_Code], [Last_Follow_Up_Code_Timestamp], [NEB], [PAP], [O2_RNTL], [NI_VENT], [NPWT], [PAP_SUPPLY], [OTHER], [BED], [Date], [SAP_ID], [CITRIX_ID], [Name], [Supervisor], [Wave], [Track_Roster_Wise], [Track_Queue_Wise], [Division])"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11, @12, @13, @14, @15, @16, @17, @18, @19, @20, @21, @22, @23, @24, @25, @26, @27, @28, @29, @30, @31, @32, @33, @34, @35, @36, @37, @38, @39, @40, @41, @42, @43)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-")|| row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.NVarChar))
                                        .Value = row.Cell(1).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                        .Value = row.Cell(2).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                        .Value = row.Cell(3).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.NVarChar))
                                        .Value = row.Cell(4).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.NVarChar))
                                        .Value = row.Cell(5).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.NVarChar))
                                        .Value = row.Cell(6).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.NVarChar))
                                        .Value = row.Cell(7).Value.ToString();
                                    if (row.Cell(8).Value.ToString().Equals(""))
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.DateTime)).Value = DBNull.Value;
                                    }
                                    else
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.DateTime)).Value = row.Cell(8).GetDateTime();
                                    }
                                    // DateTime.FromOADate(row.Cell(7).GetDouble() - For Excel DOubles as Dates  
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.NVarChar))
                                        .Value = row.Cell(9).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@10", SqlDbType.NVarChar))
                                        .Value = row.Cell(10).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@11", SqlDbType.NVarChar))
                                        .Value = row.Cell(11).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@12", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(12).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@13", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(13).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@14", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(14).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@15", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(15).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@16", SqlDbType.NVarChar))
                                        .Value = row.Cell(16).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@17", SqlDbType.NVarChar))
                                        .Value = row.Cell(17).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@18", SqlDbType.NVarChar))
                                        .Value = row.Cell(18).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@19", SqlDbType.NVarChar))
                                        .Value = row.Cell(19).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@20", SqlDbType.DateTime))
                                            .Value = row.Cell(20).GetDateTime(); 
                                    cmd.Parameters.Add(new SqlParameter("@21", SqlDbType.DateTime))
                                        .Value = row.Cell(21).GetDateTime(); 
                                    cmd.Parameters.Add(new SqlParameter("@22", SqlDbType.DateTime))
                                        .Value = row.Cell(22).GetDateTime(); 
                                    cmd.Parameters.Add(new SqlParameter("@23", SqlDbType.DateTime))
                                        .Value = row.Cell(23).GetDateTime(); 
                                    if(row.Cell(24).GetDateTime() <= DateTime.MinValue) {
                                        cmd.Parameters.Add(new SqlParameter("@24", SqlDbType.DateTime))
                                        .Value = DBNull.Value;
                                    }
                                    else if(row.Cell(24).GetDateTime() >= DateTime.MaxValue) {
                                        cmd.Parameters.Add(new SqlParameter("@24", SqlDbType.DateTime))
                                        .Value = DBNull.Value;
                                    }
                                    else { 
                                    cmd.Parameters.Add(new SqlParameter("@24", SqlDbType.DateTime))
                                        .Value = row.Cell(24).GetDateTime();
                                    } 
                                    cmd.Parameters.Add(new SqlParameter("@25", SqlDbType.NVarChar))
                                        .Value = row.Cell(25).Value.ToString();
                                    if (row.Cell(26).Value.ToString().Equals("")) {
                                        cmd.Parameters.Add(new SqlParameter("@26", SqlDbType.DateTime)).Value = DBNull.Value;
                                    }
                                    else {
                                        cmd.Parameters.Add(new SqlParameter("@26", SqlDbType.DateTime)).Value = row.Cell(26).GetDateTime();
                                    } 
                                    cmd.Parameters.Add(new SqlParameter("@27", SqlDbType.NVarChar))
                                        .Value = row.Cell(27).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@28", SqlDbType.NVarChar))
                                        .Value = row.Cell(28).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@29", SqlDbType.NVarChar))
                                        .Value = row.Cell(29).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@30", SqlDbType.NVarChar))
                                        .Value = row.Cell(30).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@31", SqlDbType.NVarChar))
                                        .Value = row.Cell(31).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@32", SqlDbType.NVarChar))
                                        .Value = row.Cell(32).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@33", SqlDbType.NVarChar))
                                        .Value = row.Cell(33).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@34", SqlDbType.NVarChar))
                                        .Value = row.Cell(34).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@35", SqlDbType.DateTime))
                                        .Value = row.Cell(35).Value; 
                                    cmd.Parameters.Add(new SqlParameter("@36", SqlDbType.Int))
                                        .Value = row.Cell(36).Value;
                                    cmd.Parameters.Add(new SqlParameter("@37", SqlDbType.NVarChar))
                                        .Value = row.Cell(37).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@38", SqlDbType.NVarChar))
                                        .Value = row.Cell(38).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@39", SqlDbType.NVarChar))
                                        .Value = row.Cell(39).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@40", SqlDbType.NVarChar))
                                        .Value = row.Cell(40).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@41", SqlDbType.NVarChar))
                                        .Value = row.Cell(41).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@42", SqlDbType.NVarChar))
                                        .Value = row.Cell(42).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@43", SqlDbType.NVarChar))
                                        .Value = row.Cell(43).Value.ToString();

                                    insertCount++;
                                    cn.Open();
                                    try
                                    { 
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}"; 
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> UsersUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("PROJECT");
                        } catch {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].users " + //42 columns excl. ID
                        "([sap_id], [name], [password], [user_role], [designation], [division], [status], [sub_department], [phase], [band], [tenurity], [hcl_hire_date], [abay_start_date], [cms_id], [citrix], [nt_login], [finesse_extension], [finesse_names], [finesse_enterprise_names], [badge_id], [birth_date], [address], [contact_number], [nda], [nho/policies_sign_off], [bgv], [versant], [typing], [aptitude], [group_policy])"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11, @12, @13, @14, @15, @16, @17, @18, @19, @20, @21, @22, @23, @24, @25, @26, @27, @28, @29, @30)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                        .Value = row.Cell(2).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                        .Value = "pass";
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.NVarChar))
                                        .Value = "Agent";
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.NVarChar))
                                        .Value = row.Cell(11).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.NVarChar))
                                        .Value = row.Cell(12).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.NVarChar))
                                        .Value = row.Cell(13).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.NVarChar))
                                        .Value = row.Cell(14).Value.ToString();
                                    //if (row.Cell(8).Value.ToString().Equals(""))
                                    //{
                                    //    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.DateTime)).Value = DBNull.Value;
                                    //}
                                    //else
                                    //{
                                    //    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.DateTime)).Value = row.Cell(8).GetDateTime();
                                    //}
                                    // DateTime.FromOADate(row.Cell(7).GetDouble() - For Excel DOubles as Dates  
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.NVarChar))
                                        .Value = row.Cell(15).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@10", SqlDbType.NVarChar))
                                        .Value = row.Cell(16).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@11", SqlDbType.NVarChar))
                                        .Value = row.Cell(17).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@12", SqlDbType.DateTime))
                                        .Value = row.Cell(18).GetDateTime();
                                    if (row.Cell(19).GetString() == "")
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@13", SqlDbType.DateTime))
                                            .Value = DBNull.Value; 
                                    }
                                    else { 
                                    cmd.Parameters.Add(new SqlParameter("@13", SqlDbType.DateTime))
                                        .Value = row.Cell(19).GetDateTime();
                                    }
                                    cmd.Parameters.Add(new SqlParameter("@14", SqlDbType.NVarChar))
                                        .Value = row.Cell(20).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@15", SqlDbType.NVarChar))
                                        .Value = row.Cell(21).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@16", SqlDbType.NVarChar))
                                        .Value = row.Cell(22).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@17", SqlDbType.NVarChar))
                                        .Value = row.Cell(23).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@18", SqlDbType.NVarChar))
                                        .Value = row.Cell(24).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@19", SqlDbType.NVarChar))
                                        .Value = row.Cell(25).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@20", SqlDbType.NVarChar))
                                        .Value = row.Cell(26).Value.ToString();

                                    if (row.Cell(27).CachedValue == null)
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@21", SqlDbType.NVarChar))
                                        .Value = "";
                                    }
                                    else
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@21", SqlDbType.NVarChar))
                                        .Value = row.Cell(27).GetString();
                                    }
                                    if (row.Cell(28).CachedValue == null ) {
                                        cmd.Parameters.Add(new SqlParameter("@22", SqlDbType.NVarChar))
                                        .Value = "";
                                    }
                                    else
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@22", SqlDbType.NVarChar))
                                        .Value = row.Cell(28).GetString();
                                    }
                                    if (row.Cell(29).CachedValue == null ) {
                                        cmd.Parameters.Add(new SqlParameter("@23", SqlDbType.NVarChar))
                                        .Value = "";
                                    }
                                    else
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@23", SqlDbType.NVarChar))
                                        .Value = row.Cell(29).GetString();
                                    }
                                    cmd.Parameters.Add(new SqlParameter("@24", SqlDbType.NVarChar))
                                        .Value = row.Cell(31).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@25", SqlDbType.NVarChar))
                                        .Value = row.Cell(32).Value.ToString();
                                    if (row.Cell(33).CachedValue == null ) {
                                        cmd.Parameters.Add(new SqlParameter("@26", SqlDbType.NVarChar))
                                        .Value = "";
                                    }
                                    else
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@26", SqlDbType.NVarChar))
                                        .Value = row.Cell(33).GetString();
                                    }
                                    cmd.Parameters.Add(new SqlParameter("@27", SqlDbType.NVarChar))
                                        .Value = row.Cell(34).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@28", SqlDbType.NVarChar))
                                        .Value = row.Cell(35).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@29", SqlDbType.NVarChar))
                                        .Value = row.Cell(36).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@30", SqlDbType.NVarChar))
                                        .Value = row.Cell(37).Value.ToString(); 

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> TeamsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].[group] " + //42 columns excl. ID
                        "([group_name], [group_leader])"
                        + " VALUES (@1, @2)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.NChar))
                                        .Value = row.Cell(2).Value.ToString();

                                    int sap_id = Int32.Parse(row.Cell(1).Value.ToString());
                                    user usr;
                                    try {
                                        usr = db.users.Where(x => x.sap_id == sap_id).First();
                                    }
                                    catch
                                    {
                                        usr = db.users.Where(x => x.user_id == 1).First();
                                    }
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Int))
                                        .Value = usr.user_id;  

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> LOBsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].track " + //42 columns excl. ID
                        "([track_name], [track_manager])"
                        + " VALUES (@1, @2)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.NVarChar))
                                        .Value = row.Cell(2).Value.ToString();

                                    int sap_id = Int32.Parse(row.Cell(1).Value.ToString());
                                    user usr;
                                    try
                                    {
                                        usr = db.users.Where(x => x.sap_id == sap_id).First();
                                    }
                                    catch
                                    {
                                        usr = db.users.Where(x => x.user_id == 1).First();
                                    }

                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Int))
                                        .Value = usr.user_id;

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> LOBsUpdateUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row

                        foreach (var row in WorkSheet.RowsUsed())
                        {
                            //do something here
                            if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                            {
                                Debug.WriteLine("- Empty row encountered");
                            }
                            else
                            {
                                string _sql = string.Format("UPDATE [dbo].[group] SET" + //42 columns excl. ID
                                "[track_id] = @2" +
                                " WHERE [group_leader] = @1");
                                using (SqlConnection cn = Utilities.getConn())
                                {
                                    var cmd = new SqlCommand(_sql, cn); 

                                    int sap_id = Int32.Parse(row.Cell(1).Value.ToString());
                                    user usr;
                                    try
                                    {
                                        usr = db.users.Where(x => x.sap_id == sap_id).First();
                                    }
                                    catch
                                    {
                                        usr = db.users.Where(x => x.user_id == 1).First();
                                    }

                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = usr.user_id;
                                    int currId = usr.user_id;
                                    try { 
                                        sap_id = Int32.Parse(row.Cell(2).Value.ToString());
                                    }
                                    catch
                                    {
                                        sap_id = 1;
                                    }
                                    try
                                    {
                                        usr = db.users.Where(x => x.sap_id == sap_id).First();
                                    }
                                    catch
                                    {
                                        usr = db.users.Where(x => x.user_id == 1).First();
                                    }
                                    track trk;
                                    try
                                    {
                                        trk = db.tracks.Where(x => x.track_manager == usr.user_id).First();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            trk = db.tracks.Where(x => x.track_manager == currId).First();
                                        }
                                        catch
                                        {
                                            trk = db.tracks.Where(x => x.track_manager == 1).First();
                                        }
                                    } 
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Int))
                                        .Value = trk.track_id;
                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                        }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> TeamsUpdateUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row

                        foreach (var row in WorkSheet.RowsUsed())
                        {
                            //do something here
                            if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                            {
                                Debug.WriteLine("- Empty row encountered");
                            }
                            else
                            {
                                string _sql = string.Format("UPDATE [dbo].[users] SET " + //42 columns excl. ID
                                "[group_id] = @2" +
                                " WHERE [user_id] = @1");
                                using (SqlConnection cn = Utilities.getConn())
                                {
                                    var cmd = new SqlCommand(_sql, cn);

                                    int sap_id = Int32.Parse(row.Cell(1).Value.ToString());
                                    user usr;
                                    try
                                    {
                                        usr = db.users.Where(x => x.sap_id == sap_id).First();
                                    }
                                    catch
                                    {
                                        usr = db.users.Where(x => x.user_id == 1).First();
                                    }

                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = usr.user_id;
                                    int currId = usr.user_id;
                                    try
                                    {
                                        sap_id = Int32.Parse(row.Cell(2).Value.ToString());
                                    }
                                    catch
                                    {
                                        sap_id = 1;
                                    }
                                    try
                                    {
                                        usr = db.users.Where(x => x.sap_id == sap_id).First();
                                    }
                                    catch
                                    {
                                        usr = db.users.Where(x => x.user_id == 1).First();
                                    }
                                    group grp;
                                    try
                                    {
                                        grp = db.groups.Where(x => x.group_leader == usr.user_id).First();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            grp = db.groups.Where(x => x.group_leader == currId).First();
                                        }
                                        catch
                                        {
                                            grp = db.groups.Where(x => x.group_leader == 1).First();
                                        }
                                    }
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Int))
                                        .Value = grp.group_id;
                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                        }
                        String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                        Debug.WriteLine("{status:'Upload Complete'}");
                        file = null;
                        GC.Collect();
                        return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> FinessesUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("reason code");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].finesse " + //42 columns excl. ID
                        "([sap_id],[finesse_agent_team], [finesse_agent_name], [finesse_detail_event], [finesse_detail_event_datetime], [finesse_detail_reason_code], [finesse_detail_reason_code_name], [finesse_detail_duration])"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                               {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                        .Value = row.Cell(2).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                        .Value = row.Cell(3).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.NVarChar))
                                        .Value = row.Cell(4).Value.ToString();
                                    if (row.Cell(5).Value.ToString().Equals(""))
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.DateTime)).Value = DBNull.Value;
                                    }
                                    else
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.DateTime)).Value = row.Cell(5).GetDateTime();
                                    }
                                    cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.NVarChar))
                                        .Value = row.Cell(6).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.NVarChar))
                                        .Value = row.Cell(7).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(8).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> LmsesUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].lms " + //42 columns excl. ID
                        "([user_id],[training_title], [status], [registered_date], [due_date], [completed_date])"
                        + " VALUES (@1, @2, @3, @4, @5, @6)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString().Substring(0,8));
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                        .Value = row.Cell(6).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                        .Value = row.Cell(7).Value.ToString(); 
                                    if (row.Cell(8).Value.ToString().Equals("")){
                                        cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.DateTime)).Value = DBNull.Value;
                                    }else{
                                        cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.DateTime)).Value = row.Cell(8).GetDateTime();
                                    } 
                                    if (row.Cell(9).Value.ToString().Equals("")){
                                        cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Date)).Value = DBNull.Value;
                                    }else{
                                        cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Date)).Value = row.Cell(9).GetDateTime();
                                    } 
                                    if (row.Cell(10).Value.ToString().Equals("") || row.Cell(10).Value.ToString().Equals("N/A") || row.Cell(10).Value.ToString().Equals("LOA") || row.Cell(10).Value.ToString().Equals("ABSCONDING"))
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.DateTime)).Value = DBNull.Value;
                                    }else{
                                        cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.DateTime)).Value = row.Cell(10).GetDateTime();
                                    } 

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> WpusUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls") || file.FileName.EndsWith(".xlsm"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Raw Data");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].wpu " + //42 columns excl. ID
                        "([track], [week], [sap_id], [examinee_type], [project_name], [LOB], [location], [marks_obtained], [total_marks], [passing_percent], [result])"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.NVarChar))
                                        .Value = row.Cell(1).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                        .Value = row.Cell(2).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(3).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.NVarChar))
                                        .Value = row.Cell(5).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.NVarChar))
                                        .Value = row.Cell(6).Value.ToString(); 
                                    cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.NVarChar))
                                        .Value = row.Cell(7).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.NVarChar))
                                        .Value = row.Cell(8).Value.ToString();
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(11).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(12).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@10", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(13).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@11", SqlDbType.NVarChar))
                                        .Value = row.Cell(14).Value.ToString();
                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public async Task<JsonResult> SchedulesUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls") || file.FileName.EndsWith(".xlsm"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("Sheet1");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].schedule " + //42 columns excl. ID
                        "([sap_id], [day], [shift])"
                        + " VALUES (@1, @2, @3)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Date))
                                        .Value = row.Cell(3).GetDateTime();
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                        .Value = row.Cell(4).Value.ToString(); 
                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> KaiserClosetOverallsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("dump");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].kaiser_closet_overall " + //42 columns excl. ID
                        "([sap_id], [ave_prod], [ave_prod_score], [otc], [otc_score], [eom_score], [rank])"
                        + " VALUES (@1, @2, @3, @4, @5, @8, @9)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString()); 
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(7).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(8).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(9).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(10).Value.ToString()); 
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(23).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(24).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> PPMCBPMOverallsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("dump");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].ppmc_bpm_overall " + //42 columns excl. ID
                        "([sap_id], [bpm], [bpm_score], [otc], [otc_score], [eom_score], [rank])"
                        + " VALUES (@1, @2, @3, @4, @5, @8, @9)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(6).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(7).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(8).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(9).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(22).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(23).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> KaiserSMCOverallsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("dump");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].kaiser_smc_overall " + //42 columns excl. ID
                        "([sap_id], [ave_prod], [ave_prod_score], [eom_score], [rank])"
                        + " VALUES (@1, @2, @3, @8, @9)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(7).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(8).Value.ToString());  
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(21).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(22).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> KaiserOthersOverallsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("dump");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].kaiser_others_overall " + //42 columns excl. ID
                        "([sap_id], [ave_prod], [ave_prod_score], [eom_score], [rank])"
                        + " VALUES (@1, @2, @3, @8, @9)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(7).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(8).Value.ToString());  
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(21).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(22).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public async Task<JsonResult> PPMCL1OverallsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("dump");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].ppmcl1_overall " + //42 columns excl. ID
                        "([sap_id], [ave_calls_handled], [ave_calls_handled_score], [aht], [aht_score], [cash_col], [cash_col_score], [eom_score], [rank])"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString()); 
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(6).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(7).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(8).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(9).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(10).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(11).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(24).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(25).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public async Task<JsonResult> PPMCL2OverallsUpload()
        {
            int insertCount = 0;
            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Files/Uploads/"), fileName);
                file.SaveAs(path);
                if (file.ContentLength > 0)
                {
                    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook Workbook = new XLWorkbook();
                        try
                        {
                            Workbook = new XLWorkbook(file.InputStream);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Check your file. {ex.Message}");
                        }
                        IXLWorksheet WorkSheet = null;
                        try
                        {
                            WorkSheet = Workbook.Worksheet("dump");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].ppmcl2_overall " + //42 columns excl. ID
                        "([sap_id], [ave_calls_handled], [ave_calls_handled_score], [aht], [aht_score], [cash_col], [cash_col_score], [eom_score], [rank])"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                if (row.Cell(1).GetString().Equals("-") || row.Cell(1).GetString().Equals(""))
                                {
                                    Debug.WriteLine("- Empty row encountered");
                                }
                                else
                                {
                                    var cmd = new SqlCommand(_sql, cn);
                                    cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(1).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(6).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(7).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(8).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(9).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(10).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.Int))
                                        .Value = Decimal.Parse(row.Cell(11).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Decimal))
                                        .Value = Decimal.Parse(row.Cell(24).Value.ToString());
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Int))
                                        .Value = Int32.Parse(row.Cell(25).Value.ToString());

                                    insertCount++;
                                    cn.Open();
                                    try
                                    {
                                        await cmd.ExecuteNonQueryAsync();
                                        Debug.WriteLine("{status:'Line Inserted " + insertCount + "'}");
                                        cmd.Dispose();
                                        cn.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        cn.Close();
                                        String str = "{message:'" + e.Message + "'}";
                                        Debug.WriteLine(str);
                                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                                    }
                                }
                                Workbook.Dispose();
                            }
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
                            file = null;
                            GC.Collect();
                            return Json(JObject.Parse(str2).ToString(), JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        String str = "{message:'Only.xlsx and .xls files are allowed'}";
                        Debug.WriteLine(str);
                        return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    String str = "{message:'Not a valid file'}";
                    Debug.WriteLine(str);
                    return Json(JObject.Parse(str).ToString(), JsonRequestBehavior.AllowGet);
                }
            }

            String line = "{message:'No Valid File Found'}";
            Debug.WriteLine(line);
            return Json(JObject.Parse(line).ToString(), JsonRequestBehavior.AllowGet);
        }
        public ActionResult DownloadEQProdTemplate()
        {
            string filename = "EqProd.xlsx";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/Files/Excel_Templates/" + filename;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
        public ActionResult DownloadLOBTemplate()
        {
            string filename = "LOB_Template.xlsx";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/Files/Excel_Templates/" + filename;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
        public ActionResult DownloadLOBUpdateTemplate()
        {
            string filename = "LOB_Update_Template.xlsx";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/Files/Excel_Templates/" + filename;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
        public ActionResult DownloadTeamTemplate()
        {
            string filename = "Team_Template.xlsx";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/Files/Excel_Templates/" + filename;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
        public ActionResult DownloadTeamUpdateTemplate()
        {
            string filename = "Team_Update_Template.xlsx";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/Files/Excel_Templates/" + filename;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
        public ActionResult DownloadUsersTemplate()
        {
            string filename = "Users_Template.xlsx";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/Files/Excel_Templates/" + filename;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
        private double tryGetData(SqlDataReader reader, string columnName){ 
            if (reader.IsDBNull(reader.GetOrdinal(columnName))){
               return 0;
            } else {
                return double.Parse(reader[columnName].ToString());
            }
        }
    }
}