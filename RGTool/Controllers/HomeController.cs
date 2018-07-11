using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RGTool.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
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

        [HttpPost]
        public string CreateExcel(string filepath, string ooxml)
        {
            try
            {
                SpreadsheetDocument document = SpreadsheetDocument.Create(@filepath, SpreadsheetDocumentType.Workbook, true);
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart sheetpart = workbookPart.AddNewPart<WorksheetPart>();
                sheetpart.Worksheet = new Worksheet(ooxml);

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(workbookPart), SheetId = 1, Name = "Test Sheet Name" };
                sheets.Append(sheet);

                document.Close();
            }
            catch (Exception ex)
            {
                return "错误原因：" + ex.Message;
            }
            return "CreateExcel function has finshed!";
        }
    }
}