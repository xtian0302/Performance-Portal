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

    }
}
