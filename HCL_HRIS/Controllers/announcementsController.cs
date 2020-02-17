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

namespace HCL_HRIS.Controllers
{
    public class announcementsController : Controller
    {
        private HCL_HRISEntities db = new HCL_HRISEntities();

        // GET: announcements
        public async Task<ActionResult> Index()
        {
            int sap_id = Int32.Parse(User.Identity.Name.Trim());
            user usr = db.users.Where(x => x.sap_id == sap_id).First(); 
            ViewBag.user = usr;
            return View(await db.announcements.ToListAsync());
        }

        // GET: announcements/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            announcement announcement = await db.announcements.FindAsync(id);
            if (announcement == null)
            {
                return HttpNotFound();
            }
            return View(announcement);
        }

        // GET: announcements/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: announcements/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "announcement_id,announcement_title,announcement_details,announcement_date")] announcement announcement)
        {
            if (ModelState.IsValid)
            {
                db.announcements.Add(announcement);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(announcement);
        }

        // GET: announcements/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            announcement announcement = await db.announcements.FindAsync(id);
            if (announcement == null)
            {
                return HttpNotFound();
            }
            return View(announcement);
        }

        // POST: announcements/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "announcement_id,announcement_title,announcement_details,announcement_date")] announcement announcement)
        {
            if (ModelState.IsValid)
            {
                db.Entry(announcement).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(announcement);
        }

        // GET: announcements/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            announcement announcement = await db.announcements.FindAsync(id);
            if (announcement == null)
            {
                return HttpNotFound();
            }
            return View(announcement);
        }

        // POST: announcements/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            announcement announcement = await db.announcements.FindAsync(id);
            db.announcements.Remove(announcement);
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
    }
}
