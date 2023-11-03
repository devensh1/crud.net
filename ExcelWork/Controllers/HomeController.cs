using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelWork.Models;
using OfficeOpenXml;

namespace ExcelWork.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        DataConsolidationEntities db = new DataConsolidationEntities();

        public HomeController()
        {
            // Set the LicenseContext based on your licensing status
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial
        }

        public ActionResult Index()
        {
            List<student> data = db.students.ToList();

            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(string));
                dt.Columns.Add("Department", typeof(string));

                foreach (var item in data)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = item.Id;
                    dr[1] = item.Name;
                    dr[2] = item.Age;
                    dr[3] = item.Department;
                    dt.Rows.Add(dr);
                }
                var memory = new MemoryStream();
                using (var Excel = new ExcelPackage(memory))
                {
                    var Worksheet = Excel.Workbook.Worksheets.Add("Student1");
                    Worksheet.Cells["A1"].LoadFromDataTable(dt);
                    Worksheet.Cells["A1:AN1"].Style.Font.Bold = true;
                    Worksheet.DefaultRowHeight = 25;
                    Session["mywork"] = Excel.GetAsByteArray();
                    return Json("", JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public ActionResult Download()
        {
            if (Session["mywork"] != null)
            {
                byte[] md = Session["mywork"] as byte[];
                return File(md, "application/octet-stream", "FileManager.xlsx");
            }

            return View();
        }
    }
}
