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
    public class tracksController : Controller
    {
        private HCL_HRISEntities db = new HCL_HRISEntities();

        // GET: tracks
        public async Task<ActionResult> Index()
        {
            int id = int.Parse(User.Identity.Name);
            user usr = db.users.Where(x => x.sap_id == id).First();
            if (usr.user_role == "Administrator" || usr.user_role == "Manager")
            {
                return View(await db.tracks.ToListAsync());
            }
            else
            {
                return new HttpNotFoundResult("You are not Allowed Access to this Page");
            }
        }

        // GET: tracks/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            track track = await db.tracks.FindAsync(id);
            if (track == null)
            {
                return HttpNotFound();
            }
            return View(track);
        }

        // GET: tracks/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: tracks/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "track_id,track_name,track_manager")] track track)
        {
            if (ModelState.IsValid)
            {
                db.tracks.Add(track);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(track);
        }

        // GET: tracks/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            track track = await db.tracks.FindAsync(id);
            if (track == null)
            {
                return HttpNotFound();
            }
            return View(track);
        }

        // POST: tracks/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "track_id,track_name,track_manager")] track track)
        {
            if (ModelState.IsValid)
            {
                db.Entry(track).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(track);
        }

        // GET: tracks/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            track track = await db.tracks.FindAsync(id);
            if (track == null)
            {
                return HttpNotFound();
            }
            return View(track);
        }

        // POST: tracks/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            track track = await db.tracks.FindAsync(id);
            db.tracks.Remove(track);
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
