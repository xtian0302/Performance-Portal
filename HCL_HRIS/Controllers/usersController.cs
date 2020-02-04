using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HCL_HRIS.Models;
using System.Web.Configuration;
using System.Web.Security;
using System.Diagnostics;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;

namespace HCL_HRIS.Controllers
{
    public class usersController : Controller
    {
        private HCL_HRISEntities db = new HCL_HRISEntities();

        // GET: users
        public async Task<ActionResult> Index()
        {
            int id = int.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == id).First();
            if (usr.user_role == "Administrator")
            {
                return View(await db.users.ToListAsync());
            }
            else
            {
                return new HttpNotFoundResult("You are not Allowed Access to this Page");
            }
        }

        // GET: users/Team
        public async Task<ActionResult> Team()
        {
            int id = int.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == id).First();
            if (usr.user_role == "Team Leader")
            {
                return View(await db.users.Where(x => x.group.user.sap_id == id).ToListAsync());
            }
            else
            {
                return new HttpNotFoundResult("You are not Allowed Access to this Page");
            }
        }

        public async Task<ActionResult> View(int? id)
        { 
            user usr = db.users.Where(x => x.user_id == id).First();
            int sap_id = usr.sap_id;
            SqlConnection connection = Utilities.getConn(); 
            //queries start here
            //Get top 5 agents of track 
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

            //Get ranking of this agent against other agents
            command = new SqlCommand("Select rank, (Select Count(*) from rankings where track = 'Sleep EQ') as count from (Select RANK() OVER(ORDER BY score DESC) as rank,sap_id as sapno from rankings where track = 'Sleep EQ') tb where sapno = @1", connection);
            command.Parameters
                .Add(new SqlParameter("@1", SqlDbType.Int))
                .Value = sap_id; 
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
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = sap_id; 
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
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = sap_id;  
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
                ViewBag.lms = double.Parse(string.Format("{0:0.#}", decimal.Parse(reader["LmsScore"].ToString())));
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
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = sap_id; 
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
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = sap_id;

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
            command.Parameters.Add("@sap_id", SqlDbType.VarChar).Value = sap_id; 
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
        // GET: users/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            user user = await db.users.FindAsync(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            return View(user);
        }

        // GET: users/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: users/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "user_id,sap_id,name,password,group_id,user_role")] user user)
        {
            if (ModelState.IsValid)
            {
                try { 
                db.users.Add(user);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
                }
                catch (DbUpdateException ex)
                {
                    if (ExceptionHelper.IsUniqueConstraintViolation(ex))
                    {
                        ModelState.AddModelError("sap_id", $"The SAP Number '{user.sap_id}' is already in use, please enter a different SAP Number.");
                        return View(nameof(Create), user);
                    }
                }
            }

            return View(user);
        }

        // GET: users/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            user user = await db.users.FindAsync(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            return View(user);
        }

        // POST: users/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "user_id,sap_id,name,password,group_id,user_role")] user user)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    db.Entry(user).State = EntityState.Modified;
                    await db.SaveChangesAsync();
                    return RedirectToAction("Index");
                }
                catch (DbUpdateException ex)
                {
                    if (ExceptionHelper.IsUniqueConstraintViolation(ex))
                    {
                        ModelState.AddModelError("sap_id", $"The SAP Number '{user.sap_id}' is already in use, please enter a different SAP Number.");
                        return View(nameof(Edit), user);
                    }
                }
            }
            return View(user);
        }

        // GET: users/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            user user = await db.users.FindAsync(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            return View(user);
        }

        // POST: users/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            user user = await db.users.FindAsync(id);
            db.users.Remove(user);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
        public ActionResult Login()
        {
            return View();
        }

        // Login And LogOut POST
        [HttpPost]
        public ActionResult Login(Models.user user)
        {
            if (ModelState.IsValid)
            {
                if (Utilities.IsValid(user.sap_id.ToString(), user.password))
                {
                    FormsAuthenticationConfiguration formsAuthentication = new FormsAuthenticationConfiguration();
                    FormsAuthentication.SetAuthCookie(user.sap_id.ToString(), true);   
                    return RedirectToAction("Index", "Home");
                }
                else
                {
                    ModelState.AddModelError("", "SAP Number and/or Password is wrong!");
                }
            }
            return View(user);
        }
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Login", "Users");
        }

        // GET: users/GetGroups
        [HttpGet]
        public JsonResult GetGroups()
        {
            var objCustomerlist = db.groups.ToList();
            var subCategoryToReturn = objCustomerlist.Select(S => new {
                group_id = S.group_id,
                group_name = S.group_name,
                group_leader = S.group_leader
            });
            return Json(subCategoryToReturn, JsonRequestBehavior.AllowGet);
        }

        // GET: users/GetGroupChats
        [HttpGet]
        public IQueryable<chat> GetGroupChats()
        {
            int? sap_id = Int32.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == sap_id).First();
            int? group_id = usr.group_id;
            return db.chats.Where(x => x.group_id == group_id);
        }

        // GET: users/GetMemChats
        [HttpGet]
        public IQueryable<chat> GetMemChats(int sap_id)
        {
            int? my_sap = Int32.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == my_sap).First();
            user leader = db.users.Where(x => x.sap_id == sap_id).First();
            int? sup_id = leader.user_id, sup_sap_id = leader.sap_id;
            return db.chats.Where(x => (x.chat_from == usr.user_id && x.chat_to == sup_sap_id) || ((x.chat_from == sup_id && x.chat_to == usr.sap_id)));
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
