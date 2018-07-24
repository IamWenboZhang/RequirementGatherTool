using RGTool.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RGTool.Controllers
{
    public class ExcelController : Controller
    {
        private static string ExcelTemplateFolderPath = "~/Excel/Template/";
        private static string ExcelFolderPath = "~/Excel/ExcelFile/";
        // GET: Excel
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Excel()
        {
            return View();
        }
        [HttpPost]
        public string CreateExcel(string templateName, string JsonContent, string Configdata)
        {
            //JsonContent=HttpUtility.UrlEncode(JsonContent, System.Text.Encoding.GetEncoding("UTF8")); 
            string result = "";
            JsonContent = HttpUtility.UrlDecode(JsonContent, System.Text.Encoding.GetEncoding("UTF-8"));
            Configdata = HttpUtility.UrlDecode(Configdata, System.Text.Encoding.GetEncoding("UTF-8"));
            try
            {
                string newDocName = DateTime.Now.ToString("yyyyMMddhhmm") + Guid.NewGuid() + ".xlsx";
                string path = Server.MapPath(ExcelTemplateFolderPath + templateName);
                string newDocPath = Server.MapPath(ExcelFolderPath + newDocName);
                ExcelUtil.ApplyTemplate(path, JsonContent, newDocPath, Configdata);
                result = "Success";
                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}