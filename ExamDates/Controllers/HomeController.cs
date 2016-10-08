using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExamDates.Models;
using System.IO;
using System.Text;
using System.Web.Routing;
using System.Threading.Tasks;
using System.Net;

namespace ExamDates.Controllers
{
    public class HomeController : Controller
    {
        public List<ExamModel> exams = new List<ExamModel>();
        
        public ActionResult Index(string message)
        {
            if(message != null)
            ViewBag.Error =  message;
            return View();
        }

        [HttpPost]
        public ActionResult Result(String str)
        {
            ReadXLSFILE(str);
            exams.OrderBy(c => c.FirstDate).OrderBy(c => c.SecondDate);
            ViewBag.Exams = exams;
            return View();
        }

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            String fileNew = "";
            try
            {
                if (file != null && file.ContentLength > 0)
                {
                    var fileName = Path.GetFileName(file.FileName);
                    fileNew = fileName;
                    var path = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
                    file.SaveAs(path);
                }
                else if (Request.Files.Count > 0)
                {
                    file = Request.Files[0];
                    var fileName = Path.GetFileName(file.FileName);
                    fileNew = fileName;
                    var path = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
                    file.SaveAs(path);
                    Result(path);
                }

                if (fileNew != null)
                {
                    ViewBag.report = "File uploaded successfully!";
                    return View("Result");
                }
                else 
                {
                    ViewBag.report = "File upload failed!";
                    return View("Index");
                }
            }
            catch (Exception e) 
            {
                ViewBag.report = e.Message;
                return View("Index");
            }
            
            return View("Result");
        }
        public void ReadXLSFILE(string filename) 
        {
            int rCnt = 0;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            var range = xlWorkSheet.UsedRange;

           try
           {
               for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
               {
                   string name = (range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string date1 = (range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string regStartPerDate1 = (range.Cells[rCnt, 3] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string regEndPerDate1 = (range.Cells[rCnt, 4] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string date2 = (range.Cells[rCnt, 5] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string regStartPerDate2 = (range.Cells[rCnt, 6] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string regEndPerDate2 = (range.Cells[rCnt, 7] as Microsoft.Office.Interop.Excel.Range).Value2;
                   string session = (range.Cells[rCnt, 8] as Microsoft.Office.Interop.Excel.Range).Value2;
                   RegistrationDates regDiapason1 = new RegistrationDates();
                   RegistrationDates regDiapason2 = new RegistrationDates();
                   ExamModel exam = new ExamModel();
                   if (name != null)
                   {
                       exam.Name = name; 
                   }
                   if (date1 != null)
                   {
                       exam.FirstDate = DateTime.Parse(date1);
                   }
                   if (date2 != null) 
                   {
                       exam.SecondDate = DateTime.Parse(date2);
                   }
                   if (regStartPerDate1 != null) 
                   {
                       regDiapason1.RegStartDate = DateTime.Parse(regStartPerDate1);
                   }
                   if (regEndPerDate1 != null) 
                   {
                       regDiapason1.RegEndDate = DateTime.Parse(regEndPerDate1);
                   }
                   exam.RegPerFirstDate = regDiapason1;
                   if (regStartPerDate2 != null)
                   {
                       regDiapason2.RegStartDate = DateTime.Parse(regStartPerDate2);
                   }
                   if (regEndPerDate2 != null)
                   {
                       regDiapason2.RegEndDate = DateTime.Parse(regEndPerDate2);
                   }
                   exam.RegPerSecondDate = regDiapason2;
                   if (session != null || session != string.Empty) 
                   {
                       exam.Session = session;
                   }
                   if (name != null && (date1 != null || date2 != null)) 
                   {
                       exams.Add(exam);
                   }
               }
           }
           catch (Exception e) 
           {
               ViewBag.Error = "Something went wrong!";
           }
           finally 
           {
               xlWorkBook.Close(true, null, null);
               xlApp.Quit();

               releaseObject(xlWorkSheet);
               releaseObject(xlWorkBook);
               releaseObject(xlApp);
           }
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