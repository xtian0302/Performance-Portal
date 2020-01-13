﻿using ClosedXML.Excel;
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
            SqlConnection connection = Utilities.getConn();

            SqlCommand command = new SqlCommand("get_Errors", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name; 
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            List<string> BCList = new List<string>();
            double eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0;
            while (reader.Read()){
                bcErrorCurrMos = int.Parse(reader["bcErrorCurrMos"].ToString());
                eucErrorCurrMos = int.Parse(reader["eucErrorCurrMos"].ToString());
                ccErrorCurrMos = int.Parse(reader["ccErrorCurrMos"].ToString());
                totalAuditCurrMos = int.Parse(reader["totalAuditCurrMos"].ToString()); 
            }
            reader.Close();
            command.Dispose();
            //------------------------ BC EUC CC ViewBag ----------------------------00:P1} 
            if (totalAuditCurrMos != 0){
               ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(1.0-((double)bcErrorCurrMos/(double)totalAuditCurrMos), 1.0-((double)eucErrorCurrMos/(double)totalAuditCurrMos), 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            } else {
                ViewBag.QAScore = HCL_HRIS.Models.Calculations.getQAScoredProd(0, 0, 0);
            } 

            connection.Close(); 

            command = new SqlCommand("get_Prod", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name; 
            connection.Open();
            reader = command.ExecuteReader();
            double aveprod = 0.0, cmplt = 0.0, otc = 0.0;
            while (reader.Read())
            { //  Current Month Preaderroductivity
                    aveprod = tryGetData(reader, "AveProd");
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
            }
            ViewBag.ProdScore = HCL_HRIS.Models.Calculations.getOverallScoredProd(aveprod,cmplt,otc);
            ViewBag.OverallScore = string.Format("{0:0.##}", (ViewBag.ProdScore * 0.45) + (ViewBag.QAScore * 0.3));
            ViewBag.ProdScore = string.Format("{0:0.#}", ViewBag.ProdScore);
            ViewBag.QAScore = string.Format("{0:0.#}", ViewBag.QAScore);
            reader.Close();
            command.Dispose();
            connection.Close();

            command = new SqlCommand("get_MinsPerDay", connection);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = User.Identity.Name;
            connection.Open();
            reader = command.ExecuteReader(); 
            List<MinsPerDay> minsCollection = new List<MinsPerDay>();
            while (reader.Read()){
                MinsPerDay mins = new MinsPerDay();
                mins.minsLogged = (int.Parse(reader["Minutes"].ToString()));
                mins.minsDate = (Convert.ToDateTime(reader["Date"])); 
                minsCollection.Add(mins);
            } 
            reader.Close();
            command.Dispose();
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
            ViewBag.user_role = usr.user_role;
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

        public async Task<ActionResult> Details()
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

        [HttpPost]
        public JsonResult EqUpload()
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
                                        cmd.ExecuteNonQuery();
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
        public JsonResult UsersUpload()
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
                                        cmd.ExecuteNonQuery();
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
        public JsonResult TeamsUpload()
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
                                        cmd.ExecuteNonQuery();
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
        public JsonResult LOBsUpload()
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
                                        cmd.ExecuteNonQuery();
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
        public JsonResult LOBsUpdateUpload()
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
                                        cmd.ExecuteNonQuery();
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
        public JsonResult TeamsUpdateUpload()
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
                                        cmd.ExecuteNonQuery();
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
        public JsonResult FinessesUpload()
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
                                        cmd.ExecuteNonQuery();
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
        private int tryGetData(SqlDataReader reader, string columnName){ 
            if (reader.IsDBNull(reader.GetOrdinal(columnName))){
               return 0;
            } else {
               return int.Parse(reader[columnName].ToString());
            }
        }
    }
}