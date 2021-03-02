using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel_Import_ASPdotNETMVC.Models;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Excel_Import_ASPdotNETMVC.Controllers
{
    public class ProductController : Controller
    {
        // GET: Student
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase excelfile)
        {
            if(excelfile == null ||excelfile.ContentLength == 0)
            {
                ViewBag.error = "Please select a excel file";
                return View();
            }
            else
            {
                if(excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/File/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    // Read data from excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<Student> listproduct = new List<Student>();

                    for(int row = 2; row <= range.Rows.Count; row++)
                    {
                        Student s = new Student();
                        s.id = ((Excel.Range)range.Cells[row, 1]).Text;
                        s.Name = ((Excel.Range)range.Cells[row, 2]).Text;
                        s.Price = ((Excel.Range)range.Cells[row, 3]).Text;
                        s.Quantity = ((Excel.Range)range.Cells[row, 4]).Text;

                        listproduct.Add(s);
                    }
                    ViewBag.list = listproduct;
                     ViewBag.error = "Success";
                    return View(listproduct);
                }
                else
                {
                    ViewBag.error = "File type is incorrect only xls, xlsx is accepted";
                    return View();
                }
            }
            
        }
    }
}