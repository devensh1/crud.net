using ExcelWork.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelWork.Controllers
{
    public class StudentController : Controller
    {
        // GET: Student
        DataConsolidationEntities db = new DataConsolidationEntities();
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Add(student model)
        {
            student obj = new student();
            obj.Id = model.Id;
            obj.Name = model.Name;
            obj.Age = model.Age;
            obj.Department = model.Department;
            db.students.Add(obj);
            db.SaveChanges();
            ModelState.Clear();
            return View("Index");
        }
        public ActionResult AddList()
        {
            var res = db.students.ToList();
            return View(res);
        }
    }
}