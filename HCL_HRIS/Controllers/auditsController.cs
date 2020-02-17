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
using System.IO;
using System.Diagnostics;
using ClosedXML.Excel;
using System.Data.SqlClient;
using Newtonsoft.Json.Linq;

namespace HCL_HRIS.Controllers
{
    public class auditsController : Controller
    {
        private HCL_HRISEntities db = new HCL_HRISEntities();

        // GET: audits
        public async Task<ActionResult> Index()
        {
            return View(await db.audits.OrderByDescending(x => x.audit_id).Take(1000).ToListAsync());
        }
        // GET: audits/Import
        public ActionResult Import()
        {
            return View();
        }
        // GET: audits/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            audit audit = await db.audits.FindAsync(id);
            if (audit == null)
            {
                return HttpNotFound();
            }
            return View(audit);
        }

        // GET: audits/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: audits/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "audit_id,case_id,audit_type,type_of_monitoring,auditor_sap,agent_sap,audit_date,transaction_date,bc,euc,cc,fatal")] audit audit)
        {
            if (ModelState.IsValid)
            {
                db.audits.Add(audit);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(audit);
        }

        // GET: audits/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            audit audit = await db.audits.FindAsync(id);
            if (audit == null)
            {
                return HttpNotFound();
            }
            return View(audit);
        }

        // POST: audits/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "audit_id,case_id,audit_type,type_of_monitoring,auditor_sap,agent_sap,audit_date,transaction_date,bc,euc,cc,fatal")] audit audit)
        {
            if (ModelState.IsValid)
            {
                db.Entry(audit).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(audit);
        }

        // GET: audits/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            audit audit = await db.audits.FindAsync(id);
            if (audit == null)
            {
                return HttpNotFound();
            }
            return View(audit);
        }

        // POST: audits/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            audit audit = await db.audits.FindAsync(id);
            db.audits.Remove(audit);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }
        public ActionResult DownloadFile()
        {
            string filename = "Audit_Template.xlsx";
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
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        //- Upload -

        [HttpPost]
        public JsonResult Upload()
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
                        string _sql = string.Format("INSERT INTO [dbo].audit (case_id,audit_type,type_of_monitoring,auditor_sap,agent_sap,audit_date,transaction_date,bc,euc,cc,fatal)"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11)");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                var cmd = new SqlCommand(_sql, cn);
                                cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.NVarChar))
                                    .Value = row.Cell(1).Value.ToString();
                                cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                    .Value = row.Cell(2).Value.ToString();
                                cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                    .Value = row.Cell(3).Value.ToString();
                                cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.Int))
                                    .Value = Int32.Parse( row.Cell(4).Value.ToString());
                                cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Int))
                                    .Value = Int32.Parse(row.Cell(5).Value.ToString());
                                cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.DateTime))
                                    .Value = row.Cell(6).GetDateTime();
                                cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.DateTime))
                                    .Value = row.Cell(7).GetDateTime();
                                // DateTime.FromOADate(row.Cell(7).GetDouble() - For Excel DOubles as Dates
                                //Pass = 0, Fail = 1
                                if (row.Cell(8).Value.ToString().ToUpper().Equals("PASS")) { 
                                cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Bit)).Value = false;
                                } else if(row.Cell(8).Value.ToString().ToUpper().Equals("FAIL")){
                                cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Bit)).Value = true; 
                                } else {
                                cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Bit)).Value = DBNull.Value; 
                                }
                                if (row.Cell(9).Value.ToString().ToUpper().Equals("PASS")) { 
                                cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Bit)).Value = false;
                                } else if(row.Cell(9).Value.ToString().ToUpper().Equals("FAIL")) {
                                cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Bit)).Value = true;
                                } else {
                                cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Bit)).Value = DBNull.Value; 
                                }
                                if (row.Cell(10).Value.ToString().ToUpper().Equals("PASS")) { 
                                cmd.Parameters.Add(new SqlParameter("@10", SqlDbType.Bit)).Value = false;
                                } else if(row.Cell(10).Value.ToString().ToUpper().Equals("FAIL")){
                                cmd.Parameters.Add(new SqlParameter("@10", SqlDbType.Bit)).Value = true; 
                                } else {
                                cmd.Parameters.Add(new SqlParameter("@10", SqlDbType.Bit)).Value = DBNull.Value; 
                                }
                                if (row.Cell(11).Value.ToString().ToUpper().Equals("PASS")) { 
                                cmd.Parameters.Add(new SqlParameter("@11", SqlDbType.Bit)).Value = false;
                                } else if(row.Cell(11).Value.ToString().ToUpper().Equals("FAIL")){
                                cmd.Parameters.Add(new SqlParameter("@11", SqlDbType.Bit)).Value = true; 
                                } else {
                                cmd.Parameters.Add(new SqlParameter("@11", SqlDbType.Bit)).Value = DBNull.Value; 
                                }
                                
                                insertCount++;
                                cn.Open();
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                    Debug.WriteLine("{status:'Line Inserted'}");
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
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
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
        //- Upload -

        [HttpPost]
        public JsonResult WpuUpload()
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
                            WorkSheet = Workbook.Worksheet("Raw Data");
                        }
                        catch
                        {
                            Debug.WriteLine("sheet not found!");
                        }
                        WorkSheet.FirstRow().Delete();//if you want to remove 1st row
                        string _sql = string.Format("INSERT INTO [dbo].wpu (sap_id,examinee_type,project_name,LOB,location,marks_obtained,total_marks,passing_percent,result)"
                        + " VALUES (@1, @2, @3, @4, @5, @6, @7, @8, @9");
                        using (SqlConnection cn = Utilities.getConn())
                        {
                            foreach (var row in WorkSheet.RowsUsed())
                            {
                                //do something here
                                var cmd = new SqlCommand(_sql, cn);
                                cmd.Parameters.Add(new SqlParameter("@1", SqlDbType.NVarChar))
                                    .Value = row.Cell(1).Value.ToString();
                                cmd.Parameters.Add(new SqlParameter("@2", SqlDbType.NVarChar))
                                    .Value = row.Cell(2).Value.ToString();
                                cmd.Parameters.Add(new SqlParameter("@3", SqlDbType.NVarChar))
                                    .Value = row.Cell(3).Value.ToString();
                                cmd.Parameters.Add(new SqlParameter("@4", SqlDbType.Int))
                                    .Value = Int32.Parse(row.Cell(4).Value.ToString());
                                cmd.Parameters.Add(new SqlParameter("@5", SqlDbType.Int))
                                    .Value = Int32.Parse(row.Cell(5).Value.ToString());
                                cmd.Parameters.Add(new SqlParameter("@6", SqlDbType.DateTime))
                                    .Value = row.Cell(6).GetDateTime();
                                cmd.Parameters.Add(new SqlParameter("@7", SqlDbType.DateTime))
                                    .Value = row.Cell(7).GetDateTime();
                                // DateTime.FromOADate(row.Cell(7).GetDouble() - For Excel DOubles as Dates
                                //Pass = 0, Fail = 1
                                if (row.Cell(8).Value.ToString().ToUpper().Equals("PASS"))
                                {
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Bit)).Value = false;
                                }
                                else if (row.Cell(8).Value.ToString().ToUpper().Equals("FAIL"))
                                {
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Bit)).Value = true;
                                }
                                else
                                {
                                    cmd.Parameters.Add(new SqlParameter("@8", SqlDbType.Bit)).Value = DBNull.Value;
                                }
                                if (row.Cell(9).Value.ToString().ToUpper().Equals("PASS"))
                                {
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Bit)).Value = false;
                                }
                                else if (row.Cell(9).Value.ToString().ToUpper().Equals("FAIL"))
                                {
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Bit)).Value = true;
                                }
                                else
                                {
                                    cmd.Parameters.Add(new SqlParameter("@9", SqlDbType.Bit)).Value = DBNull.Value;
                                } 
                                insertCount++;
                                cn.Open();
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                    Debug.WriteLine("{status:'Line Inserted'}");
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
                            String str2 = "{status:'OK',Count:'" + insertCount + "'}";
                            Debug.WriteLine("{status:'Upload Complete'}");
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
    }
}
