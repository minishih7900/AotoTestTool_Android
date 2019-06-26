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
        private StringBuilder verificationErrors;
        private bool acceptNextAlert = true;

        //ClassTEST執行前必執行程序
        [ClassInitialize]
        public static void InitializeClass(TestContext testContext)
        {
            //Chrom設定
            ChromeOptions options = new ChromeOptions();
            //download.default_directory 預設下載會存在C:\temp\excel
            options.AddUserProfilePreference("download.default_directory", @"C:\temp\excel");
            //自行設定路徑或直接從Nuget安裝
            driver = new ChromeDriver(@"C:\temp\driver", options);
            //瀏覽器最大化
            driver.Manage().Window.Maximize();
            //等待時間10秒
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        }
        //ClassTEST執行完最後再執行
        [ClassCleanup]
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
        //TestTestMethod前執行的程序
        [TestInitialize]
        public void InitializeTest()
        {
            verificationErrors = new StringBuilder();
        }
        //TestTestMethod後執行的程序
        [TestCleanup]
        public void CleanupTest()
        {
            Assert.AreEqual("", verificationErrors.ToString());
        }

        //TestMethod測試內容
        [TestMethod]
        public void T1()
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
                            case "GoToUrl":
                                stop = TestTool.GoToUrl(driver, item2.Inivalue[0]);
                                break;
                            case "ClickFindElementByName":
                                stop = TestTool.ClickFindElementByName(driver, item2.Inivalue[0]);
                                break;
                            case "ClickFindElementByXPath":
                                stop = TestTool.ClickFindElementByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "ClickFindElementByLinkText":
                                stop = TestTool.ClickFindElementByLinkText(driver, item2.Inivalue[0]);
                                break;
                            case "ClickJSFindElementByXPath":
                                stop = TestTool.ClickJSFindElementByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "ClearFindElementByXPath":
                                stop = TestTool.ClearFindElementByXPath(driver, item2.Inivalue[0]);
                                break;
                            case "DropdownSelectValueFindElementByXPath":
                                stop = TestTool.DropdownSelectValueFindElementByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "DropdownSelectTextFindElementByXPath":
                                stop = TestTool.DropdownSelectTextFindElementByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "DropdownSelectTextEveryoneInputFindElementByXPath":
                                stop = TestTool.DropdownSelectTextEveryoneInputFindElementByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "SendKeysFindElementByName":
                                stop = TestTool.SendKeysFindElementByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "SendKeysFindElementByXPath":
                                stop = TestTool.SendKeysFindElementByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "SearchTextGetEndandclicknextpage":
                                stop = TestTool.SearchTextGetEndandclicknextpage(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "WindowScroll":
                                stop = TestTool.WindowScroll(driver, item2.Inivalue[0]);
                                break;
                            case "ThreadSleep":
                                stop = TestTool.ThreadSleep(driver, item2.Inivalue[0]);
                                break;
                            case "AssertAreEqualDriverTitle":
                                stop = TestTool.AssertAreEqualDriverTitle(driver, item2.Inivalue[0]);
                                break;
                            case "AssertAreEqualByXPathText":
                                stop = TestTool.AssertAreEqualByXPathText(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "AssertAreEqualByXPathValue":
                                stop = TestTool.AssertAreEqualByXPathValue(driver, item2.Inivalue[0], item2.Inivalue[1]);
                                break;
                            case "CheckSelectListValuetFindElementByClassName":
                                stop = TestTool.CheckSelectListValuetFindElementByClassName(driver, item2.Inivalue[0], item2.Inivalue[1]);
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
