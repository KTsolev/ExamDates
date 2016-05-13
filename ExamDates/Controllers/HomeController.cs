using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExamDates.Models;
using System.IO;
using System.Web.Routing;

namespace ExamDates.Controllers
{
    public class HomeController : Controller
    {
        public List<ExamModel> exams = new List<ExamModel>();
        // GET: Home
        //public ActionResult Index()
        //{
        //    return View();
        //}
        
        public ActionResult Index(string message)
        {
            if(message != null)
            ViewBag.Error =  message;
            return View();
        }

        [HttpPost]
        public ActionResult Result(FormCollection form)
        {
            string fileName = form["FileName"].ToString();
            ReadXLSFILE(fileName);
            ViewBag.Exams = exams;
            return View();
        }
        [HttpPost]
        public ActionResult Report(FormCollection form)
        {
            string fileName = form["FileName"].ToString();
            //check for reportName parameter value now
            //to do : Return something
            string error = string.Empty;
            
            if (fileName != null)
            {
                ViewBag.FileName = fileName;
            }
            else
            {
                error = "You have not uploaded file!Please choose excel file!";
                return (RedirectToAction("Index", new {message = error}));
            }
            if (!fileName.EndsWith(".xls") || !fileName.EndsWith(".xlsx"))
            {
                error = "You have uploaded wrong file!Please choose excel file!";
                return (RedirectToAction("Index", new { message = error }));
            }
            return View();
        }
        public void ReadXLSFILE(string filename) 
        {
            int rCnt = 0;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

           var range = xlWorkSheet.UsedRange;

            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
               string name = (string)(range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2;
               string date1 = (string)(range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2;
               string date2 = (string)(range.Cells[rCnt, 3] as Microsoft.Office.Interop.Excel.Range).Value2;
               ExamModel exam = new ExamModel();
               if(name!=null)
               exam.Name = name;
               if(date1 != null)
               exam.FirstDate = DateTime.Parse(date1);
               if(date2 != null)
               exam.SecondDate = DateTime.Parse(date2);
               if(name != null && (date1!=null || date2!=null))
               exams.Add(exam);
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
         private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
               Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}