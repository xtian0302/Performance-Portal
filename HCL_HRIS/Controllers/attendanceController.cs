using System;
using System.Collections.Generic;
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
            return View();
        }
    }
}