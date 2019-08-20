using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.IO;
using UnitTestProject1.Model;
using ClassLibrary1;
using System.Linq;


namespace UnitTestProject1
{
    public class baseTest
    {
        public static List<IniModel> IniData;
        public static TestMethodList TestTool = new TestMethodList();
        public static string FormatNum2 = "{0:0.00}";
        public static string FormatNum3 = "{0:0.000}";
        public int AutoNum = 0;


        public static string readIni(string title, string data)
        {
            IniManager iniTool = new IniManager(@"C:\temp\log\manager.ini");
            try
            {
                string iniValue = iniTool.ReadIniFile(title, data, "error");
                Assert.IsFalse(iniValue.Contains("error"));
                return iniValue;
            }
            catch (Exception ex)
            {
                TestTool.writeLog("Trace", "----------------------------------------------------------------------------------------");
                TestTool.writeLog("Error", title + " | " + data + " | 可能不存在於INI中");
                TestTool.writeLog("Error", ex);
                TestTool.writeLog("Trace", "----------------------------------------------------------------------------------------");
                Trace.WriteLine(ex);
                throw;
            }

        }
        public static List<IniModel> ReadAllData(string path)
        {
            StreamReader reader = new StreamReader(path, System.Text.Encoding.UTF8);
            string[] content = reader.ReadToEnd().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            reader.Close();
            List<IniModel> result = new List<IniModel>();
            int index = -1;
            foreach (string s in content)
            {
                if (s.StartsWith("//"))
                {
                    continue;
                }
                if (s.StartsWith("["))
                {
                    result.Add(new IniModel() { IniType = s ,IniDetail = new List<Detail>() });
                    index++;
                }
                else
                {
                    
                        string[] ss = s.Split(new char[] { '=' }, 2);
                        List<string> sss = ss[1].Split(',').ToList();
                        result[index].IniDetail.Add(new Detail() { IniName = ss[0], Inivalue = sss });
                }
            }
            return result;
        }
      
    }
}
