using Newtonsoft.Json;
using RGTool.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RGTool.Controllers
{
    public class ConfigController : Controller
    {
        private string ConfigFilePath = "~/ConfigFile/config.txt";
        // GET: Config
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Config()
        {
            
            return View();
        }

        [HttpPost]
        public string SetConfig(string shortname,string version,TypesofTD type, string startsection,string endsection)
        {
            string result = String.Empty;
            try
            {
                string path = Server.MapPath(ConfigFilePath);
                Config config = new Config() { TDShortName = shortname, TDVersion = version, TDType = type, StartSection = startsection, EndSection = endsection };
                //result = ConfigUtil.SavetoFile(config, path);
                result = JsonConvert.SerializeObject(config);
                return result;
            }
            catch(Exception ex)
            {
                return result;
            }
        }

        [HttpGet]
        public string GetConfig()
        {
            //Config result = new Config();
            string result = "";
            try
            {
                string path = Server.MapPath(ConfigFilePath);
                result = ConfigUtil.ReadConfigfromFile(path);
                return result;
            }
            catch(Exception ex)
            {
                return null;
            }
        }
    }
}