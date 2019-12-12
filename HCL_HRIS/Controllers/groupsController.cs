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
using System.Diagnostics;

namespace HCL_HRIS.Controllers
{
    public class groupsController : Controller
    {
        private HCL_HRISEntities db = new HCL_HRISEntities();

        // GET: groups
        public async Task<ActionResult> Index()
        {  
            return View(await db.groups.ToListAsync());
        }

        // GET: groups/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            group group = await db.groups.FindAsync(id);
            if (group == null)
            {
                return HttpNotFound();
            }
            return View(group);
        }

        // GET: groups/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: groups/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "group_id,group_name,group_leader,track_id")] group group)
        {
            if (ModelState.IsValid)
            {
                try { 
                db.groups.Add(group);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
                }catch(Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }

            return View(group);
        }

        // GET: groups/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            group group = await db.groups.FindAsync(id);
            if (group == null)
            {
                return HttpNotFound();
            }
            return View(group);
        }

        // POST: groups/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "group_id,group_name,group_leader,track_id")] group group)
        {
            if (ModelState.IsValid)
            {
                db.Entry(group).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(group);
        }

        // GET: groups/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            group group = await db.groups.FindAsync(id);
            if (group == null)
            {
                return HttpNotFound();
            }
            return View(group);
        }

        // POST: groups/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            group group = await db.groups.FindAsync(id);
            db.groups.Remove(group);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        // GET: groups/GetUsers

        [HttpGet]
        public JsonResult GetUsers()
        {
            var objCustomerlist = db.users.ToList();
            var subCategoryToReturn = objCustomerlist.Select(S => new {
                user_id = S.user_id,
                name = S.name,
                sap_id = S.sap_id 
            });
            return Json(subCategoryToReturn, JsonRequestBehavior.AllowGet);
        }
        // GET: groups/GetUsersTracks

        [HttpGet]
        public JsonResult GetTracks()
        {
            var objCustomerlist = db.tracks.ToList();

            return Json(objCustomerlist, JsonRequestBehavior.AllowGet);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
