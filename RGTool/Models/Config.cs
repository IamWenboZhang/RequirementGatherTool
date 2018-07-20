using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace RGTool.Models
{
    public enum TypesofTD { None,Block,FileFormat,SOAP,WOPI }
    public class Config
    {
        public string TDVersion { get; set; }
        public string TDShortName { get; set; }
        public TypesofTD TDType { get; set; }
        public string StartSection { get; set; }
        public string EndSection { get; set; }
    }

    public class ConfigUtil
    {
        public static bool SavetoFile(Config config,string configfilepath)
        {
            bool result = false;
            try
            {
                var jsonStr = JsonConvert.SerializeObject(config);             
                File.WriteAllText(configfilepath, jsonStr);
                result = true;
            }
            catch(Exception ex)
            {
                result = false;
            }
            return result;
        }

        public static string ReadConfigfromFile(string configfilepath)
        {
            //Config result = new Config();
            string result = "";
            try
            {
                var jsonStr = File.ReadAllText(configfilepath);
                //result = JsonConvert.DeserializeObject<Config>(jsonStr);
                result = jsonStr;
            }
            catch(Exception ex)
            {
                return null;
            }
            return result;
        }
    }
}