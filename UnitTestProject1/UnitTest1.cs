using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using ClassLibrary1;
using System.IO;
using UnitTestProject1.Model;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1:baseTest
    {
        //初始設定
        private static IWebDriver driver;
        private static ChromeOptions options;
        private StringBuilder verificationErrors;
        private bool acceptNextAlert = true;

        //ClassTEST執行前必執行程序
        public static void InitializeClass()
        {
            //Chrome設定
            options = new ChromeOptions();
            //download.default_directory 預設下載會存在C:\temp\excel
            options.AddUserProfilePreference("download.default_directory", @"C:\temp\excel");
            //設定Chrome使用者設定資料目錄
            //options.AddArgument(@"user-data-dir=C:\selenum\AutomationProfile\");
            //自行設定路徑或直接從Nuget安裝
            driver = new ChromeDriver(@"C:\temp\driver", options);
            //瀏覽器最大化
            driver.Manage().Window.Maximize();
            //等待時間10秒
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        }
        //ClassTEST執行完最後再執行
        public static void CleanupClass()
        {
            try
            {
                //driver.Quit();// quit does not close the window
                //driver.Close();
                //driver.Dispose();
            }
            catch (Exception)
            {
                // Ignore errors if unable to close the browser
            }
        }
        

        //TestMethod測試內容
        [TestMethod]
        public void T1()
        {

            // 取得資料夾內所有檔案
            foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"C:\temp\log\", "TEST*.ini"))
            {
                TestTool.writeLog("Trace", "----開始讀取----" + Path.GetFileName(fname));
                List<IniModel>  IniData = ReadAllData(fname);
                if (IniData.Count>0)
                {
                    InitializeClass();
                    TestFunction(IniData);
                }
            }
        }

        private static void TestFunction(List<IniModel> IniData)
        {
            var funName = "";
            bool stop = true;
            try
            {
                foreach (var item in IniData)
                {
                    TestTool.writeLog("Trace", "----測試開始----" + item.IniType);

                    foreach (var item2 in item.IniDetail)
                    {
                        funName = item2.IniName;

                        switch (item2.IniName)
                        {
                            case "AssertAreEqual_DriverTitle":
                                stop = TestTool.AssertAreEqual_DriverTitle(driver, item2.Inivalue[0]);
                                break;
                            case "AssertAreEqual_Text_ByXPath":
                                stop = TestTool.AssertAreEqual_Text_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Text_ById":
                                stop = TestTool.AssertAreEqual_Text_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Text_ByName":
                                stop = TestTool.AssertAreEqual_Text_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Text_ByTagName":
                                stop = TestTool.AssertAreEqual_Text_ByTagName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Text_ByClassName":
                                stop = TestTool.AssertAreEqual_Text_ByClassName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Value_ByXPath":
                                stop = TestTool.AssertAreEqual_Value_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Value_ById":
                                stop = TestTool.AssertAreEqual_Value_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Value_ByName":
                                stop = TestTool.AssertAreEqual_Value_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Value_ByTagName":
                                stop = TestTool.AssertAreEqual_Value_ByTagName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_Value_ByClassName":
                                stop = TestTool.AssertAreEqual_Value_ByClassName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_GetCssColor_ByXPath":
                                stop = TestTool.AssertAreEqual_GetCssColor_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1], item2.Inivalue[2]);
                                break;
                            case "AssertAreEqual_GetCssColor_ById":
                                stop = TestTool.AssertAreEqual_GetCssColor_ById(driver, item2.Inivalue[0], item2.Inivalue[1], item2.Inivalue[2]);
                                break;
                            case "AssertAreEqual_GetCssColor_ByName":
                                stop = TestTool.AssertAreEqual_GetCssColor_ByName(driver, item2.Inivalue[0], item2.Inivalue[1], item2.Inivalue[2]);
                                break;
                            case "AssertAreEqual_ElementExist_ByXPath":
                                stop = TestTool.AssertAreEqual_ElementExist_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_ElementExist_ById":
                                stop = TestTool.AssertAreEqual_ElementExist_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_ElementExist_ByName":
                                stop = TestTool.AssertAreEqual_ElementExist_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_SelectListValuet_ByClassName":
                                stop = TestTool.AssertAreEqual_SelectListValuet_ByClassName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_DropdownSelectValue_ByXPath":
                                stop = TestTool.AssertAreEqual_DropdownSelectValue_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_DropdownSelectValue_ById":
                                stop = TestTool.AssertAreEqual_DropdownSelectValue_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqual_DropdownSelectValue_ByName":
                                stop = TestTool.AssertAreEqual_DropdownSelectValue_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "ClickFindElement_ById":
                                stop = TestTool.ClickFindElement_ById(driver, item2.Inivalue[0]);
                                break;
                            case "ClickFindElement_ByName":
                                stop = TestTool.ClickFindElement_ByName(driver, item2.Inivalue[0]);
                                break;
                            case "ClickFindElement_ByXPath":
                                stop = TestTool.ClickFindElement_ByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "ClickFindElement_ByLinkText":
                                stop = TestTool.ClickFindElement_ByLinkText(driver, item2.Inivalue[0]);
                                break;
                            case "ClickJSFindElement_ByXPath":
                                stop = TestTool.ClickJSFindElement_ByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "ClearFindElement_ByXPath":
                                stop = TestTool.ClearFindElement_ByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "ClearFindElement_ById":
                                stop = TestTool.ClearFindElement_ById(driver, item2.Inivalue[0]);
                                break;
                            case "ClearFindElement_ByName":
                                stop = TestTool.ClearFindElement_ByName(driver, item2.Inivalue[0]);
                                break;
                            case "Dropdown_SelectValue_ByXPath":
                                stop = TestTool.Dropdown_SelectValue_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectValue_ById":
                                stop = TestTool.Dropdown_SelectValue_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectValue_ByName":
                                stop = TestTool.Dropdown_SelectValue_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectText_ByXPath":
                                stop = TestTool.Dropdown_SelectText_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectText_ById":
                                stop = TestTool.Dropdown_SelectText_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectText_ByName":
                                stop = TestTool.Dropdown_SelectText_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectTextEveryoneInput_ByXPath":
                                stop = TestTool.Dropdown_SelectTextEveryoneInput_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectTextEveryoneInput_ById":
                                stop = TestTool.Dropdown_SelectTextEveryoneInput_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Dropdown_SelectTextEveryoneInput_ByName":
                                stop = TestTool.Dropdown_SelectTextEveryoneInput_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "Driver_SwitchTo":
                                stop = TestTool.Driver_SwitchTo(driver, item2.Inivalue[0]);
                                break;
                            case "Driver_Quit":
                                stop = TestTool.Driver_Quit(driver);
                                break;
                            case "Driver_Close":
                                stop = TestTool.Driver_Close(driver);
                                break;
                            case "GoToUrl":
                                stop = TestTool.GoToUrl(driver, item2.Inivalue[0]);
                                break;
                            case "IJavaScriptExecutor_BrowserAddPaging":
                                stop = TestTool.IJavaScriptExecutor_BrowserAddPaging(driver, item2.Inivalue[0]);
                                break;
                            case "PullDownScroll_ByXPath":
                                stop = TestTool.PullDownScroll_ByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "PullDownScroll_ByXPath_ClickNextPage":
                                stop = TestTool.PullDownScroll_ByXPath_ClickNextPage(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "SendKeys_ByXPath":
                                stop = TestTool.SendKeys_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "SendKeys_ById":
                                stop = TestTool.SendKeys_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "SendKeys_ByName":
                                stop = TestTool.SendKeys_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "ThreadSleep":
                                stop = TestTool.ThreadSleep(driver, item2.Inivalue[0]);
                                break;
                            case "TestRememberMe":
                                driver = TestTool.TestRememberMe(driver, item2.Inivalue[0]);
                                break;
                            //case "TempTest":
                            //    stop = TestTool.TempTest(driver, item2.Inivalue[0], item2.Inivalue[1]);
                            //    break;
                            default:
                                TestTool.writeLog("error", item2.IniName + "：不存在!!");
                                break;
                        }
                        Assert.IsTrue(stop);
                    }

                    TestTool.writeLog("Trace", "----測試結束----" + item.IniType);
                    TestTool.writeLog("Trace", "----------------------------------------------------------------------------------------");

                }
            }
            catch (Exception ex)
            {
                if (stop)
                {
                    TestTool.writeLog("error", "有問題的方法，請確認參數設定是否正確：" + funName);
                    TestTool.writeLog("error", ex);
                    Trace.WriteLine("Fail：" + funName);
                    Trace.WriteLine(ex);
                    Assert.IsFalse(true);
                }
                else
                {
                    Assert.IsFalse(true);
                }
            }
        }
    }


}
