using HCL_HRIS.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Diagnostics;
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
            double bcErrorPrevMos=0, eucErrorPrevMos = 0, ccErrorPrevMos = 0, totalAuditPrevMos = 0, eucErrorCurrMos = 0, ccErrorCurrMos = 0, bcErrorCurrMos = 0, totalAuditCurrMos = 0, bcErrorWk1 = 0, eucErrorWk1 = 0, ccErrorWk1 = 0, totalAuditWk1 = 0, bcErrorWk2 = 0, eucErrorWk2 = 0, ccErrorWk2 = 0, totalAuditWk2 = 0, eucErrorWk3 = 0, ccErrorWk3 = 0, bcErrorWk3 = 0, totalAuditWk3 = 0, bcErrorWk4 = 0, eucErrorWk4 = 0, ccErrorWk4 = 0, totalAuditWk4 = 0, bcErrorWk5 = 0, eucErrorWk5 = 0, ccErrorWk5 = 0, totalAuditWk5 = 0;
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
            //------------------------ BC EUC CC ViewBag ----------------------------
            if (totalAuditPrevMos != 0) { 
                ViewBag.BCPrev = string.Format("{00:P1}",1.0 - ((double)bcErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.EUCPrev = string.Format("{00:P1}", 1.0 - ((double)eucErrorPrevMos / (double)totalAuditPrevMos));
                ViewBag.CCPrev = string.Format("{00:P1}", 1.0 - ((double)ccErrorPrevMos / (double)totalAuditPrevMos));
            }
            else { 
                ViewBag.BCPrev = string.Format("{00:P1}", 0);
                ViewBag.EUCPrev = string.Format("{00:P1}", 0);
                ViewBag.CCPrev = string.Format("{00:P1}", 0);
            }
            if (totalAuditCurrMos != 0) { 
                ViewBag.BCCurr = string.Format("{00:P1}", 1.0 - ((double)bcErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.EUCCurr = string.Format("{00:P1}", 1.0 - ((double)eucErrorCurrMos / (double)totalAuditCurrMos));
                ViewBag.CCCurr = string.Format("{00:P1}", 1.0 - ((double)ccErrorCurrMos / (double)totalAuditCurrMos));
            }
            else { 
                ViewBag.BCCurr = string.Format("{00:P1}", 0);
                ViewBag.EUCCurr = string.Format("{00:P1}", 0);
                ViewBag.CCCurr = string.Format("{00:P1}", 0);
            }
            if (totalAuditWk1 != 0) { 
                ViewBag.BCWk1 = string.Format("{00:P1}", 1.0 - ((double)bcErrorWk1 / (double)totalAuditWk1));
                ViewBag.EUCWk1 = string.Format("{00:P1}", 1.0 - ((double)eucErrorWk1 / (double)totalAuditWk1));
                ViewBag.CCWk1 = string.Format("{00:P1}", 1.0 - ((double)ccErrorWk1 / (double)totalAuditWk1));
            }
            else { 
                ViewBag.BCWk1 = string.Format("{00:P1}", 0);
                ViewBag.EUCWk1 = string.Format("{00:P1}", 0);
                ViewBag.CCWk1 = string.Format("{00:P1}", 0);
            }
            if (totalAuditWk2 != 0) { 
                ViewBag.BCWk2 = string.Format("{00:P1}", 1.0 - ((double)bcErrorWk2 / (double)totalAuditWk2));
                ViewBag.EUCWk2 = string.Format("{00:P1}", 1.0 - ((double)eucErrorWk2 / (double)totalAuditWk2));
                ViewBag.CCWk2 = string.Format("{00:P1}", 1.0 - ((double)ccErrorWk2 / (double)totalAuditWk2));
            }
            else { 
                ViewBag.BCWk2 = string.Format("{00:P1}", 0);
                ViewBag.EUCWk2 = string.Format("{00:P1}", 0);
                ViewBag.CCWk2 = string.Format("{00:P1}", 0);
            }
            if (totalAuditWk3 != 0) { 
                ViewBag.BCWk3 = string.Format("{00:P1}", 1.0 - ((double)bcErrorWk3 / (double)totalAuditWk3));
                ViewBag.EUCWk3 = string.Format("{00:P1}", 1.0 - ((double)eucErrorWk3 / (double)totalAuditWk3));
                ViewBag.CCWk3 = string.Format("{00:P1}", 1.0 - ((double)ccErrorWk3 / (double)totalAuditWk3));
            }
            else { 
                ViewBag.BCWk3 = string.Format("{00:P1}", 0);
                ViewBag.EUCWk3 = string.Format("{00:P1}", 0);
                ViewBag.CCWk3 = string.Format("{00:P1}", 0);
            }
            if (totalAuditWk4 != 0) { 
                ViewBag.BCWk4 = string.Format("{00:P1}", 1.0 - ((double)bcErrorWk4 / (double)totalAuditWk4));
                ViewBag.EUCWk4 = string.Format("{00:P1}", 1.0 - ((double)eucErrorWk4 / (double)totalAuditWk4));
                ViewBag.CCWk4 = string.Format("{00:P1}", 1.0 - ((double)ccErrorWk4 / (double)totalAuditWk4));
            }
            else { 
                ViewBag.BCWk4 = string.Format("{00:P1}", 0);
                ViewBag.EUCWk4 = string.Format("{00:P1}", 0);
                ViewBag.CCWk4 = string.Format("{00:P1}", 0);
            }
            if (totalAuditWk5 != 0) { 
                ViewBag.BCWk5 = string.Format("{00:P1}", 1.0 - ((double)bcErrorWk5 / (double)totalAuditWk5));
                ViewBag.EUCWk5 = string.Format("{00:P1}", 1.0 - ((double)eucErrorWk5 / (double)totalAuditWk5));
                ViewBag.CCWk5 = string.Format("{00:P1}", 1.0 - ((double)ccErrorWk5 / (double)totalAuditWk5));
            }
            else { 
                ViewBag.BCWk5 = string.Format("{00:P1}", 0);
                ViewBag.EUCWk5 = string.Format("{00:P1}", 0);
                ViewBag.CCWk5 = string.Format("{00:P1}", 0);
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
            return View(await db.announcements.OrderBy(x => x.announcement_id).Take(3).ToListAsync());
        } 

    }
}