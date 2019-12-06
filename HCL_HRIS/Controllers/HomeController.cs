using HCL_HRIS.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
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
            return View(await db.announcements.OrderBy(x=>x.announcement_id).Take(3).ToListAsync());
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}