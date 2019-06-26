using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using System.Diagnostics;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Support.UI;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;

namespace ClassLibrary1
{
    public class TestMethodList
    {
        /// <summary>
        /// 連結到輸入的網址
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        public bool GoToUrl(IWebDriver driver,string url)
        {
            try
            {
                driver.Navigate().GoToUrl(url);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到元素名稱並進行點擊
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public bool ClickFindElementByName(IWebDriver driver, string Name)
        {
            try
            {
                driver.FindElement(By.Name(Name)).Click();
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到元素XPath並進行點擊
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool ClickFindElementByXPath(IWebDriver driver, string XPath)
        {
            try
            {
                driver.FindElement(By.XPath(XPath)).Click();
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到超連結內容並進行點擊
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="LinkText"></param>
        /// <returns></returns>
        public bool ClickFindElementByLinkText(IWebDriver driver, string LinkText)
        {
            try
            {
                driver.FindElement(By.LinkText(LinkText)).Click();
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到元素XPath再使用JS的方法進行點擊-當點擊不到時的最後手段
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool ClickJSFindElementByXPath(IWebDriver driver, string XPath)
        {
            try
            {
                IWebElement btn = driver.FindElement(By.XPath(XPath));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click()", btn);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        public bool ClearFindElementByXPath(IWebDriver driver, string XPath)
        {
            try
            {
                driver.FindElement(By.XPath(XPath)).Clear();
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到元素名稱並輸入鍵盤資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="SendKeysData"></param>
        /// <returns></returns>
        public bool SendKeysFindElementByName(IWebDriver driver, string Name,string SendKeysData)
        {
            try
            {
                driver.FindElement(By.Name(Name)).SendKeys(SendKeysData);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到元素XPath並輸入鍵盤資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="SendKeys"></param>
        /// <returns></returns>
        public bool SendKeysFindElementByXPath(IWebDriver driver, string XPath, string SendKeys)
        {
            try
            {
                driver.FindElement(By.XPath(XPath)).SendKeys(SendKeys);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 使用捲軸找到符合的文字，當找不到時會點擊下一頁的XPath
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="ClickXPath"></param>
        /// <returns></returns>
        public bool SearchTextGetEndandclicknextpage(IWebDriver driver, string XPath, string ClickXPath)
        {
            try
            {
                //使用滾動捲軸找到符合文字
                while (!ReelScroll(driver, XPath))
                {
                    //點擊下一頁
                    driver.FindElement(By.XPath(ClickXPath)).Click();
                }

                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 找到元素XPath再使用捲軸拖拉過去
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool WindowScroll(IWebDriver driver, string XPath)
        {
            try
            {
                bool result = ReelScroll(driver, XPath);
                if (result)
                {
                    passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                    return true;
                }
                else
                {
                    errorMessage("未找到資料", System.Reflection.MethodBase.GetCurrentMethod().Name);
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 等待時間
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Time"></param>
        /// <returns></returns>
        public bool ThreadSleep(IWebDriver driver, string Time)
        {
            try
            {
                Thread.Sleep(Convert.ToInt16(Time));
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 依網頁的Title內容與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqualDriverTitle(IWebDriver driver, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.Title);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 依元素XPath找到內容並與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="CheckString"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool AssertAreEqualByXPathText(IWebDriver driver, string XPath,string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.XPath(XPath)).Text);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 依元素XPath找到value的資料與輸入的文字進行驗證-例如input的value
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqualByXPathValue(IWebDriver driver, string XPath, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.XPath(XPath)).GetAttribute("value"));
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 下拉式選單-依Value選取
        /// </summary>
        /// <param name="selectValue"></param>
        /// <param name="xpath"></param>
        /// <param name="driver"></param>
        public bool DropdownSelectValueFindElementByXPath(IWebDriver driver, string XPath  ,string selectValue )
        {
            try
            {
                IWebElement selectElem = driver.FindElement(By.XPath(XPath));
                SelectElement selectObj = new SelectElement(selectElem);
                selectObj.SelectByValue(selectValue);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 下拉式選單-依Text選取
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="selectValue"></param>
        /// <returns></returns>
        public bool DropdownSelectTextFindElementByXPath(IWebDriver driver, string XPath, string selectText)
        {
            try
            {
                IWebElement selectElem = driver.FindElement(By.XPath(XPath));
                SelectElement selectObj = new SelectElement(selectElem);
                selectObj.SelectByText(selectText);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 可以根據輸入的值選擇下拉式選單，和並檢查輸入值與下拉式選單值不同資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="ListSelectText"></param>
        /// <returns></returns>
        public bool DropdownSelectTextEveryoneInputFindElementByXPath(IWebDriver driver, string XPath, string ListSelectText)
        {
            try
            {
                List<string> optionsSelectList = SplitBigScratch(ListSelectText);
                IWebElement selectElem = driver.FindElement(By.XPath(XPath));
                IList<IWebElement> options = selectElem.FindElements(By.TagName("option"));
                int NoExistCount = 0;
                for (int i = 0; i < options.Count; i++)
                {
                    if (optionsSelectList.Contains(options[i].Text))
                    {
                        options[i].Click();
                    }
                    else
                    {
                        writeLog("Info ", "不存在：" + options[i].Text);
                        Trace.WriteLine("不存在：" + options[i].Text);
                        NoExistCount++;
                    }
                }
                if (NoExistCount==0)
                {
                    passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                }
                else
                {
                    errorMessage("選項不存在", System.Reflection.MethodBase.GetCurrentMethod().Name);
                }
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        /// <summary>
        /// 檢查select選擇了那些值
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="ClassName"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool CheckSelectListValuetFindElementByClassName(IWebDriver driver, string ClassName, string ListData)
        {
            try
            {
                List<string> expected = SplitBigScratch(ListData);

                IList<IWebElement> StatusSelectList = driver.FindElements(By.ClassName(ClassName));
                int NoExistCount = 0;
                foreach (var item in StatusSelectList)
                {
                    if (!expected.Contains(item.GetAttribute("title")))
                    {
                        writeLog("Info ", "不存在：" + item.GetAttribute("title"));
                        Trace.WriteLine("不存在：" + item.GetAttribute("title"));
                        NoExistCount++;
                    }
                }
                Assert.AreEqual(0, NoExistCount);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
            //public bool TempTest(IWebDriver driver, string data, string XPath)
            //{
            //    try
            //    {
            //        data = data.Replace("{", "");
            //        data = data.Substring(0, data.Length - 1);
            //        List<string> list = new List<string>(data.Split('}'));
            //        foreach (string t in list)
            //        {
            //            writeLog("Info ", t);
            //        }
            //        writeLog("Info ", XPath);
            //        passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
            //        return true;
            //    }
            //    catch (Exception ex)
            //    {
            //        errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
            //        return false;
            //    }
        }

        


        #region 通用
        /// <summary>
        /// 執行成功訊息
        /// </summary>
        /// <param name="functionName"></param>
        /// <returns></returns>
        public bool passMessage(string functionName)
        {
            //取函式名稱
            writeLog("Info ", "PASS：" + functionName);
            Trace.WriteLine("PASS：" + functionName);
            return true;
        }
        /// <summary>
        /// 執行失敗訊息
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        public bool errorMessage(Exception ex, string functionName)
        {
            //取函式名稱
            writeLog("Error","Fail：" + functionName);
            Trace.WriteLine("Fail：" + functionName);
            writeLog("Error", ex);
            Trace.WriteLine(ex);
            return false;
        }
        public bool errorMessage(string errorMessage, string functionName)
        {
            //取函式名稱
            writeLog("Error", "Fail：" + functionName);
            Trace.WriteLine("Fail：" + functionName);
            writeLog("Error", errorMessage);
            Trace.WriteLine(errorMessage);
            return false;
        }

        /// <summary>
        /// 寫入Log
        /// </summary>
        /// <param name="title">類型</param>
        /// <param name="message">訊息</param>
        public void writeLog(string title, string message)
        {
            string dateFile = DateTime.Now.ToString("yyyy-MM-dd");
            string FilePath = @"C:\temp\log\" + dateFile;
            if (!Directory.Exists(FilePath))
            {
                //新增資料夾
                Directory.CreateDirectory(FilePath);
            }

            FilePath = FilePath + @"\test.log";
            StreamWriter sw = new StreamWriter(FilePath, true);
            sw.WriteLine(DateTime.Now + " | " + title + " | " + message);
            sw.Flush();
            sw.Close();
        }
        public void writeLog(string title, Exception message)
        {
            string dateFile = DateTime.Now.ToString("yyyy-MM-dd");
            string FilePath = @"C:\temp\log\" + dateFile;
            if (!Directory.Exists(FilePath))
            {
                //新增資料夾
                Directory.CreateDirectory(FilePath);
            }

            FilePath = FilePath + @"\test.log";
            StreamWriter sw = new StreamWriter(FilePath, true);
            sw.WriteLine(DateTime.Now + " | " + title + " | " + message);
            sw.Flush();
            sw.Close();
        }
        /// <summary>
        /// 去除大括號
        /// </summary>
        /// <param name="ListData"></param>
        /// <returns></returns>
        private static List<string> SplitBigScratch(string ListData)
        {
            ListData = ListData.Replace("{", "");
            ListData = ListData.Substring(0, ListData.Length - 1);
            List<string> expected = new List<string>(ListData.Split('}'));
            return expected;
        }
        /// <summary>
        /// 滾動捲軸
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        private bool ReelScroll(IWebDriver driver, string XPath)
        {
            try
            {
                IWebElement eles = driver.FindElement(By.XPath(XPath));
                int elesPostionX = eles.Location.X;
                int elesPostionY = eles.Location.Y;
                string js = "window.scroll(" + elesPostionX + "," + elesPostionY + ")";
                ((IJavaScriptExecutor)driver).ExecuteScript(js);
               return true;
            }
            catch (Exception)
            {

                return false;
            }
            
        }
        #endregion
    }
}
