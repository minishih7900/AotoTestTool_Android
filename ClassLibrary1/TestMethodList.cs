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
        /// 找到元素XPath再使用捲軸拖拉過去
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool WindowScroll(IWebDriver driver, string XPath)
        {
            try
            {
                IWebElement eles = driver.FindElement(By.XPath(XPath));
                int elesPostionX = eles.Location.X;
                int elesPostionY = eles.Location.Y;
                string js = "window.scroll(" + elesPostionX + "," + elesPostionY + ")";
                ((IJavaScriptExecutor)driver).ExecuteScript(js);
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

        public bool TempTest(IWebDriver driver, string data)
        {
            try
            {
                data = data.Replace("{", "");
                data = data.Substring(0, data.Length - 1);
                List<string> list = new List<string>(data.Split('}'));
                foreach (string t in list)
                {
                    writeLog("Info ", t);
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
        #endregion
    }
}
