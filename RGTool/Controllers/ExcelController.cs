using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RGTool.Controllers
{
    public class ExcelController : Controller
    {
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

        //[HttpPost]
        //public bool InsertText(string docName, string text)
        //{
        //    bool result = false;
        //    try
        //    {
        //        string path = Server.MapPath(ConfigFilePath);
        //        Config config = new Config() { TDShortName = shortname, TDVersion = version, TDType = type, StartSection = startsection, EndSection = endsection };
        //        result = ConfigUtil.SavetoFile(config, path);
        //        return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        return result;
        //    }
        //}
    }
}