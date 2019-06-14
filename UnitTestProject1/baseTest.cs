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
        public static List<IniModel> IniData = ReadAllData();
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
        public static List<IniModel> ReadAllData()
        {
            StreamReader reader = new StreamReader(@"C:\temp\log\manager.ini",System.Text.Encoding.Default);
            string[] content = reader.ReadToEnd().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            reader.Close();
            List<IniModel> result = new List<IniModel>();
            int index = -1;
            foreach (string s in content)
            {
                
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

        #region 通用
        



        /// <summary>
        /// 下拉式選單-依索引值
        /// </summary>
        /// <param name="selectIndex"></param>
        /// <param name="xpath"></param>
        /// <param name="driver"></param>
        public static void DropdownSelectIndex(int selectIndex, string xpath, IWebDriver driver)
        {
            IWebElement selectElem = driver.FindElement(By.XPath(xpath));
            SelectElement selectObj = new SelectElement(selectElem);
            selectObj.SelectByIndex(selectIndex);
        }
        /// <summary>
        /// 下拉式選單-依Value
        /// </summary>
        /// <param name="selectValue"></param>
        /// <param name="xpath"></param>
        /// <param name="driver"></param>
        public static void DropdownSelectValue(string selectValue, string xpath, IWebDriver driver)
        {
            IWebElement selectElem = driver.FindElement(By.XPath(xpath));
            SelectElement selectObj = new SelectElement(selectElem);
            selectObj.SelectByValue(selectValue);
        }
        /// <summary>
        /// 下拉式選單-依Text
        /// </summary>
        /// <param name="selectText"></param>
        /// <param name="xpath"></param>
        /// <param name="driver"></param>
        public static void DropdownSelectText(string selectText, string xpath, IWebDriver driver)
        {
            IWebElement selectElem = driver.FindElement(By.XPath(xpath));
            SelectElement selectObj = new SelectElement(selectElem);
            selectObj.SelectByText(selectText);
        }
        /// <summary>
        /// 下拉式選單-取的選取的值
        /// </summary>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public static string DropdownSelected(string xpath, IWebDriver driver)
        {
            IWebElement selectElem = driver.FindElement(By.XPath(xpath));
            SelectElement selectObj = new SelectElement(selectElem);
            IList<IWebElement> selected = selectObj.AllSelectedOptions;

            return selected[0].Text;
        }
        public static string DropdownSelected(IWebElement driver)
        {
            IWebElement selectElem = driver;
            SelectElement selectObj = new SelectElement(selectElem);
            IList<IWebElement> selected = selectObj.AllSelectedOptions;

            return selected[0].Text;
        }
        /// <summary>
        /// 當Checkbox是關著時，把他打勾
        /// </summary>
        /// <param name="xpath"></param>
        public static void CheckboxIsFalseToTrue(string xpath, IWebDriver driver)
        {
            IWebElement bikeCheckbox = driver.FindElement(By.XPath(xpath));
            if (!bikeCheckbox.Selected)
            {
                bikeCheckbox.Click();
            }

        }
        /// <summary>
        /// 當Checkbox是打勾時，把他取消
        /// </summary>
        /// <param name="xpath"></param>
        public static void CheckboxIsTrueToFalse(string xpath, IWebDriver driver)
        {
            IWebElement bikeCheckbox = driver.FindElement(By.XPath(xpath));
            if (bikeCheckbox.Selected)
            {
                bikeCheckbox.Click();
            }
        }
        /// <summary>
        /// 檢查資料是否符合
        /// </summary>
        /// <param name="data"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="colName"></param>
        public static void CheckDataObj(string data, string row, string col, string colName, IWebDriver driver)
        {
            Assert.AreEqual(data, driver.FindElement(By.XPath("//*[@id='data-table']/tbody/tr[" + row + "]/td[" + col + "]")).Text, "error：" + colName);
        }
        /// <summary>
        /// 檢查預設的Status有幾筆
        /// </summary>
        /// <param name="expected"></param>
        /// <param name="driver"></param>
        public static void CheckStatusCount(List<string> expected, IList<IWebElement> driver)
        {
            IList<IWebElement> StatusSelectList = driver;
            int NoExistCount = 0;
            foreach (var item in StatusSelectList)
            {
                if (!expected.Contains(item.GetAttribute("title")))
                {
                    Trace.WriteLine("不存在：" + item.GetAttribute("title"));
                    NoExistCount++;
                }
            }
            Assert.AreEqual(0, NoExistCount);
        }
        /// <summary>
        /// 按下Search後檢查Showing筆數
        /// </summary>
        /// <param name="count"></param>
        /// <param name="driver"></param>
        public static void CheckShowingCount(string count, IWebDriver driver)
        {
            driver.FindElement(By.XPath("//*[@id=\"btnSearch\"]")).Click();
            Assert.AreEqual("Showing 1 to " + count + " of " + count + " entries", driver.FindElement(By.XPath("//*[@id=\"data-table_info\"]")).Text);
        }
        /// <summary>
        /// 下載excel
        /// </summary>
        /// <param name="driver"></param>
        public static void ExcelDownload(IWebDriver driver)
        {
            driver.FindElement(By.XPath("//*[@id='data-table_wrapper']/div[2]/a[1]")).Click();
        }
        /// <summary>
        /// 將string轉數值後，再格式成0.00
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string FormatStr(int data, string format)
        {
            return string.Format(format, data);
        }
        public string Formatstr(DateTime data, string format)
        {
            return string.Format(format, data);
        }
        
        public int AutoNumber()
        {
            AutoNum = AutoNum + 1;
            return AutoNum;
        }
        public void AutoNumberClear()
        {
            AutoNum = 0;
        }

        #endregion
    }
}
