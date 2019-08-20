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
using System.Drawing;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Interactions;

namespace ClassLibrary1
{
    public class TestMethodList
    {
        #region A
        /// <summary>
        /// 依網頁的Title內容與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_DriverTitle(IWebDriver driver, string CheckString)
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

        //--------------------------------------------------------------------//
        /// <summary>
        /// 依元素XPath找到內容並與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Text_ByXPath(IWebDriver driver, string XPath, string CheckString)
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
        /// 依元素Id找到內容並與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Text_ById(IWebDriver driver, string Id, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.Id(Id)).Text);
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
        /// 依元素Name找到內容並與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Text_ByName(IWebDriver driver, string Name, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.Name(Name)).Text);
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
        /// 依元素TagName找到內容並與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="TagName"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Text_ByTagName(IWebDriver driver, string TagName, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.TagName(TagName)).Text);
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
        /// 依元素ClassName找到內容並與輸入的文字進行驗證
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="ClassName"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Text_ByClassName(IWebDriver driver, string ClassName, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.ClassName(ClassName)).Text);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        //--------------------------------------------------------------------//

        //--------------------------------------------------------------------//
        /// <summary>
        /// 依元素XPath找到value的資料與輸入的文字進行驗證-例如input的value
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Value_ByXPath(IWebDriver driver, string XPath, string CheckString)
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
        /// 依元素Id找到value的資料與輸入的文字進行驗證-例如input的value
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Value_ById(IWebDriver driver, string Id, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.Id(Id)).GetAttribute("value"));
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
        /// 依元素Name找到value的資料與輸入的文字進行驗證-例如input的value
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Value_ByName(IWebDriver driver, string Name, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.Name(Name)).GetAttribute("value"));
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
        /// 依元素TagName找到value的資料與輸入的文字進行驗證-例如input的value
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="TagName"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Value_ByTagName(IWebDriver driver, string TagName, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.TagName(TagName)).GetAttribute("value"));
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
        /// 依元素ClassName找到value的資料與輸入的文字進行驗證-例如input的value
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="ClassName"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_Value_ByClassName(IWebDriver driver, string ClassName, string CheckString)
        {
            try
            {
                Assert.AreEqual(CheckString, driver.FindElement(By.ClassName(ClassName)).GetAttribute("value"));
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        //--------------------------------------------------------------------//


        //--------------------------------------------------------------------//
        /// <summary>
        /// 依元素XPath檢查元素的顏色是否符合
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="item"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_GetCssColor_ByXPath(IWebDriver driver, string XPath,string item, string CheckString)
        {
            try
            {
                var CssValue = driver.FindElement(By.XPath(XPath)).GetCssValue(item);
                string hex = RgbToHex(CssValue);
                Assert.AreEqual(CheckString.ToUpper(),hex);
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
        /// 依元素Id檢查元素的顏色是否符合
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="item"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_GetCssColor_ById(IWebDriver driver, string Id, string item, string CheckString)
        {
            try
            {
                var CssValue = driver.FindElement(By.Id(Id)).GetCssValue(item);
                string hex = RgbToHex(CssValue);
                Assert.AreEqual(CheckString.ToUpper(), hex);
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
        /// 依元素Name檢查元素的顏色是否符合
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="item"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_GetCssColor_ByName(IWebDriver driver, string Name, string item, string CheckString)
        {
            try
            {
                var CssValue = driver.FindElement(By.Name(Name)).GetCssValue(item);
                string hex = RgbToHex(CssValue);
                Assert.AreEqual(CheckString.ToUpper(), hex);
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        //--------------------------------------------------------------------//

        //--------------------------------------------------------------------//
        /// <summary>
        /// 判斷測試元素XPath是否存在，可以自行輸入是要存在還是不存在。
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="exist"></param>
        /// <returns></returns>
        public bool AssertAreEqual_ElementExist_ByXPath(IWebDriver driver, string XPath, string exist)
        {
            try
            {
                bool temp = doesWebElementExist(driver,"XPath" ,XPath);
                Assert.AreEqual(exist.ToLower(), temp.ToString().ToLower());
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
        /// 判斷測試元素Id是否存在，可以自行輸入是要存在還是不存在。
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="exist"></param>
        /// <returns></returns>
        public bool AssertAreEqual_ElementExist_ById(IWebDriver driver, string Id, string exist)
        {
            try
            {
                bool temp = doesWebElementExist(driver, "Id", Id);
                Assert.AreEqual(exist.ToLower(), temp.ToString().ToLower());
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
        /// 判斷測試元素Name是否存在，可以自行輸入是要存在還是不存在。
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="exist"></param>
        /// <returns></returns>
        public bool AssertAreEqual_ElementExist_ByName(IWebDriver driver, string Name, string exist)
        {
            try
            {
                bool temp = doesWebElementExist(driver, "Name", Name);
                Assert.AreEqual(exist.ToLower(), temp.ToString().ToLower());
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        //--------------------------------------------------------------------//

        //--------------------------------------------------------------------//
        /// <summary>
        /// 檢查select選擇了那些值-適用多選取時
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="ClassName"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool AssertAreEqual_SelectListValuet_ByClassName(IWebDriver driver, string ClassName, string ListData)
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

        }
        //--------------------------------------------------------------------//

        //--------------------------------------------------------------------//
        /// <summary>
        /// 依元素XPath檢查元素選取值
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_DropdownSelectValue_ByXPath(IWebDriver driver, string XPath, string CheckString)
        {
            try
            {
                IWebElement Select = driver.FindElement(By.XPath(XPath));
                SelectElement dataSelect = new SelectElement(Select);
                IList<IWebElement> selected = dataSelect.AllSelectedOptions;
                Assert.AreEqual(CheckString, selected[0].Text);

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
        /// 依元素Id檢查元素選取值
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_DropdownSelectValue_ById(IWebDriver driver, string Id, string CheckString)
        {
            try
            {
                IWebElement Select = driver.FindElement(By.Id(Id));
                SelectElement dataSelect = new SelectElement(Select);
                IList<IWebElement> selected = dataSelect.AllSelectedOptions;
                Assert.AreEqual(CheckString, selected[0].Text);

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
        /// 依元素名稱檢查元素選取值
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="CheckString"></param>
        /// <returns></returns>
        public bool AssertAreEqual_DropdownSelectValue_ByName(IWebDriver driver, string Name, string CheckString)
        {
            try
            {
                IWebElement Select = driver.FindElement(By.Name(Name));
                SelectElement dataSelect = new SelectElement(Select);
                IList<IWebElement> selected = dataSelect.AllSelectedOptions;
                Assert.AreEqual(CheckString, selected[0].Text);

                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }

        }

        //--------------------------------------------------------------------//
        #endregion

        #region C
        /// <summary>
        /// 找到元素ID並進行點擊
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public bool ClickFindElement_ById(IWebDriver driver, string Id)
        {
            try
            {
                driver.FindElement(By.Id(Id)).Click();
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
        public bool ClickFindElement_ByName(IWebDriver driver, string Name)
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
        public bool ClickFindElement_ByXPath(IWebDriver driver, string XPath)
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
        /// 找到超連結文字並進行點擊
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="LinkText"></param>
        /// <returns></returns>
        public bool ClickFindElement_ByLinkText(IWebDriver driver, string LinkText)
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
        public bool ClickJSFindElement_ByXPath(IWebDriver driver, string XPath)
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

        //--------------------------------------------------------------------//
        /// <summary>
        /// 找到元素XPath，清除內容，例如：把輸入框內的文字清除
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool ClearFindElement_ByXPath(IWebDriver driver, string XPath)
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
        /// 找到元素Id，清除內容，例如：把輸入框內的文字清除
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <returns></returns>
        public bool ClearFindElement_ById(IWebDriver driver, string Id)
        {
            try
            {
                driver.FindElement(By.Id(Id)).Clear();
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
        /// 找到元素name，清除內容，例如：把輸入框內的文字清除
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public bool ClearFindElement_ByName(IWebDriver driver, string Name)
        {
            try
            {
                driver.FindElement(By.Name(Name)).Clear();
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        //--------------------------------------------------------------------//

        #endregion

        #region D
        //--------------------------------------------------------------------//
        /// <summary>
        /// 下拉式選單-元素Xpath-依Value選取
        /// </summary>
        /// <param name="selectValue"></param>
        /// <param name="xpath"></param>
        /// <param name="driver"></param>
        public bool Dropdown_SelectValue_ByXPath(IWebDriver driver, string XPath, string selectValue)
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
        /// 下拉式選單-元素Id-依Value選取
        /// </summary>
        /// <param name="selectValue"></param>
        /// <param name="Id"></param>
        /// <param name="driver"></param>
        public bool Dropdown_SelectValue_ById(IWebDriver driver, string Id, string selectValue)
        {
            try
            {
                IWebElement selectElem = driver.FindElement(By.Id(Id));
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
        /// 下拉式選單-元素Name-依Value選取
        /// </summary>
        /// <param name="selectValue"></param>
        /// <param name="Name"></param>
        /// <param name="driver"></param>
        public bool Dropdown_SelectValue_ByName(IWebDriver driver, string Name, string selectValue)
        {
            try
            {
                IWebElement selectElem = driver.FindElement(By.Name(Name));
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
        //--------------------------------------------------------------------//

        //--------------------------------------------------------------------//
        /// <summary>
        /// 下拉式選單-元素XPath-依Text選取
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="selectValue"></param>
        /// <returns></returns>
        public bool Dropdown_SelectText_ByXPath(IWebDriver driver, string XPath, string selectText)
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
        /// 下拉式選單-元素Id-依Text選取
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="selectValue"></param>
        /// <returns></returns>
        public bool Dropdown_SelectText_ById(IWebDriver driver, string Id, string selectText)
        {
            try
            {
                IWebElement selectElem = driver.FindElement(By.Id(Id));
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
        /// 下拉式選單-元素Name-依Text選取
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="selectValue"></param>
        /// <returns></returns>
        public bool Dropdown_SelectText_ByName(IWebDriver driver, string Name, string selectText)
        {
            try
            {
                IWebElement selectElem = driver.FindElement(By.Name(Name));
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
        //--------------------------------------------------------------------//

        //--------------------------------------------------------------------//
        /// <summary>
        /// 下拉式選單依元素XPath，可以根據輸入的值選擇下拉式選單，和並檢查輸入值與下拉式選單值不同資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="ListSelectText"></param>
        /// <returns></returns>
        public bool Dropdown_SelectTextEveryoneInput_ByXPath(IWebDriver driver, string XPath, string ListSelectText)
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
                if (NoExistCount == 0)
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
        /// 下拉式選單依元素Id，可以根據輸入的值選擇下拉式選單，和並檢查輸入值與下拉式選單值不同資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="ListSelectText"></param>
        /// <returns></returns>
        public bool Dropdown_SelectTextEveryoneInput_ById(IWebDriver driver, string Id, string ListSelectText)
        {
            try
            {
                List<string> optionsSelectList = SplitBigScratch(ListSelectText);
                IWebElement selectElem = driver.FindElement(By.Id(Id));
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
                if (NoExistCount == 0)
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
        /// 下拉式選單依元素Name，可以根據輸入的值選擇下拉式選單，和並檢查輸入值與下拉式選單值不同資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Name"></param>
        /// <param name="ListSelectText"></param>
        /// <returns></returns>
        public bool Dropdown_SelectTextEveryoneInput_ByName(IWebDriver driver, string Name, string ListSelectText)
        {
            try
            {
                List<string> optionsSelectList = SplitBigScratch(ListSelectText);
                IWebElement selectElem = driver.FindElement(By.Name(Name));
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
                if (NoExistCount == 0)
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

        //--------------------------------------------------------------------//

        /// <summary>
        /// 當有多個頁籤時，可切換頁籤，0為第1個頁籤，1為第2個頁籤
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool Driver_SwitchTo(IWebDriver driver, string index)
        {
            try
            {
                ReadOnlyCollection<string> tabs = driver.WindowHandles;
                driver.SwitchTo().Window(tabs[Convert.ToInt32(index)]);//first tab
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
        /// 關閉瀏覽器
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public bool Driver_Quit(IWebDriver driver)
        {
            try
            {
                driver.Quit();
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
        /// 關閉當前使用頁籤
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public bool Driver_Close(IWebDriver driver)
        {
            try
            {
                driver.Close();
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
       
        #endregion

        #region G
        /// <summary>
        /// 連結到輸入的網址
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        public bool GoToUrl(IWebDriver driver, string url)
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
        #endregion

        #region I
        /// <summary>
        /// 新增瀏覽器頁籤
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public bool IJavaScriptExecutor_BrowserAddPaging(IWebDriver driver,string url)
        {
            try
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("window.open('"+ url + "','_blank');");
                passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage(ex, System.Reflection.MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
        #endregion

        #region P
        /// <summary>
        /// 找到元素XPath再使用捲軸拖拉過去
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool PullDownScroll_ByXPath(IWebDriver driver, string XPath)
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
        /// 使用捲軸找到符合的文字，當找不到時會點擊下一頁的XPath
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="ClickXPath"></param>
        /// <returns></returns>
        public bool PullDownScroll_ByXPath_ClickNextPage(IWebDriver driver, string XPath, string ClickXPath)
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

        #endregion

        #region S

        /// <summary>
        /// 找到元素XPath並輸入鍵盤資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <param name="SendKeys"></param>
        /// <returns></returns>
        public bool SendKeys_ByXPath(IWebDriver driver, string XPath, string SendKeys)
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
        /// 找到元素Id並輸入鍵盤資料
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="Id"></param>
        /// <param name="SendKeys"></param>
        /// <returns></returns>
        public bool SendKeys_ById(IWebDriver driver, string Id, string SendKeys)
        {
            try
            {
                driver.FindElement(By.Id(Id)).SendKeys(SendKeys);
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
        public bool SendKeys_ByName(IWebDriver driver, string Name, string SendKeysData)
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
        
        #endregion

        #region T
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
        public IWebDriver TestRememberMe(IWebDriver driver, string url)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.open();");
            ReadOnlyCollection<string> tabs = driver.WindowHandles;
            driver.Close();
            driver.SwitchTo().Window(tabs[1]);//first tab
            js.ExecuteScript("window.close();");
            string sourceFile = @"C:\selenum\AutomationProfile\";
            string destinationFile = @"C:\temp\selenum\AutomationProfile\";
            directoryCopy(sourceFile, destinationFile);
            Thread.Sleep(3000);
            ChromeOptions options = new ChromeOptions();
            Thread.Sleep(3000);
            options.AddArgument(@"user-data-dir=C:\temp\selenum\AutomationProfile\");
            IWebDriver driver1 = new ChromeDriver(@"C:\temp\driver", options);
            driver1.Manage().Window.Maximize();
            driver1.Navigate().GoToUrl(url);
            passMessage(System.Reflection.MethodBase.GetCurrentMethod().Name);
            return driver1;
            

        }

        #endregion


        public bool TempTest(IWebDriver driver, string XPath,string data)
        {
            try
            {
                var CssValue = driver.FindElement(By.XPath(XPath)).GetCssValue(data);
                RgbToHex(CssValue);
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
            StreamWriter sw = new StreamWriter(FilePath, true, System.Text.Encoding.Default);
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
            StreamWriter sw = new StreamWriter(FilePath, true, System.Text.Encoding.Default);
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
        /// <summary>
        /// 判斷元素是否存在
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="XPath"></param>
        /// <returns></returns>
        public bool doesWebElementExist(IWebDriver driver, string FindElementName,string FindElement)
        {
            try
            {
                switch (FindElementName)
                {
                    case "XPath":
                        driver.FindElement(By.XPath(FindElement));
                        break;
                    case "Id":
                        driver.FindElement(By.Id(FindElement));
                        break;
                    case "Name":
                        driver.FindElement(By.Name(FindElement));
                        break;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public string RgbToHex(string rgb)
        {
            Regex reg = new Regex(@"[0-9].+[0-9]");
            string[] temprbg = reg.Match(rgb).ToString().Split(',');
            Color myColor = Color.FromArgb(Convert.ToInt32(temprbg[0]), Convert.ToInt32(temprbg[1]), Convert.ToInt32(temprbg[2]));
            string hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");

            return hex;
        }
        public void directoryCopy(string sourceDirectory, string targetDirectory)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(sourceDirectory);
                //获取目录下（不包含子目录）的文件和子目录
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)     //判断是否文件夹
                    {
                        if (!Directory.Exists(targetDirectory + "\\" + i.Name))
                        {
                            //目标目录下不存在此文件夹即创建子文件夹
                            Directory.CreateDirectory(targetDirectory + "\\" + i.Name);
                        }
                        //递归调用复制子文件夹
                        directoryCopy(i.FullName, targetDirectory + "\\" + i.Name);
                    }
                    else
                    {
                        //不是文件夹即复制文件，true表示可以覆盖同名文件
                        File.Copy(i.FullName, targetDirectory + "\\" + i.Name, true);
                    }
                }
            }
            catch (Exception ex)
            {
                writeLog("Error", "複製文件出現異常" + ex.Message);
            }
        }
        #endregion
    }
}
