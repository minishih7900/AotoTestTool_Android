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
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Remote;
using System.Reflection;
using System.Linq;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1:baseTest
    {
        //初始設定
        private static IWebDriver driver;
        private static ChromeOptions options;
        private static AndroidDriver<AndroidElement> androidDriver;
        private static DesiredCapabilities Dcap;
        private static string TestDriver;

        //ClassTEST執行前必執行程序
        public static void InitializeClass()
        {
            TestDriver = readIni("DefaultSet", "Driver");
            if (TestDriver == "AndroidDriver")
            {
                Dcap = new DesiredCapabilities();
                Dcap.SetCapability("platformName", readIni("AndroidSet", "platformName")); //手機系統
                Dcap.SetCapability("platformVersion", readIni("AndroidSet", "platformVersion")); //手機系統版本
                Dcap.SetCapability("deviceName", readIni("AndroidSet", "deviceName")); //手機或模擬器名稱
                Dcap.SetCapability("appPackage", readIni("AndroidSet", "appPackage")); //apk包裝名稱
                Dcap.SetCapability("appActivity", readIni("AndroidSet", "appActivity")); //app執行名稱
                androidDriver = new AndroidDriver<AndroidElement>(new Uri(readIni("AndroidSet", "Uri")), Dcap);
                
               }
            else
            {
                //Chrome設定
                options = new ChromeOptions();
                //download.default_directory 預設下載會存在C:\temp\excel
                options.AddUserProfilePreference("download.default_directory", readIni("ChromeSet", "downloadDefaultDirectoryPath"));
                //設定Chrome使用者設定資料目錄
                //options.AddArgument(@"user-data-dir=C:\selenum\AutomationProfile\");
                //自行設定路徑或直接從Nuget安裝
                driver = new ChromeDriver(readIni("ChromeSet", "driverPath"), options);
                //瀏覽器最大化
                driver.Manage().Window.Maximize();
                //等待10秒
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            }
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
            
            //取得資料夾內所有檔案
            foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"C:\temp\log\", "TEST*.ini"))
            {
                TestTool.writeLog("Trace", "----開始讀取----" + Path.GetFileName(fname));
                List<IniModel> IniData = ReadAllData(fname);
                if (IniData.Count > 0)
                {
                    InitializeClass();
                    TestFunction(IniData);
                }
            }

            //輸出方法清單
            //TestMethodList temp = new TestMethodList();
            //Type temType = temp.GetType();
            //foreach (var prop in temType.GetMethods())
            //{
            //    TestTool.writeLog("Trace", prop.ToString());
            //}


        }

        private static void TestFunction(List<IniModel> IniData)
        {
            var funName = "";
            int PassTotal = 0;
            bool stop = true;
            try
            {
                foreach (var item in IniData)
                {
                    TestTool.writeLog("Trace", "----測試開始----" + item.IniType);

                    foreach (var item2 in item.IniDetail)
                    {
                        stop = InvokeMethod(item2);
                        //stop = SwitchMethod(stop, item2);

                        Assert.IsTrue(stop);
                    }
                    PassTotal += item.IniDetail.Count;
                    TestTool.writeLog("Trace", "----測試結束----" + item.IniType);
                    TestTool.writeLog("Trace", string.Format(@"----PASS：{0}筆----" , item.IniDetail.Count));
                    TestTool.writeLog("Trace", "----------------------------------------------------------------------------------------");

                }
                TestTool.writeLog("Trace", string.Format(@"----總PASS：{0}筆----", PassTotal));
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

        private static bool InvokeMethod(Detail item2)
        {
            bool stop;
            //預設設定
            Type TestToolType = TestTool.GetType();
            ConstructorInfo magicConstructor = TestToolType.GetConstructor(Type.EmptyTypes);
            object magicClassObject = magicConstructor.Invoke(new object[] { });
            //查詢方法名稱
            MethodInfo magicMethod = TestToolType.GetMethods().Where(p => p.Name == item2.IniName && p.ToString().Contains(TestDriver)).FirstOrDefault();
            //取得方法有幾個Parameter
            int parameters = magicMethod.GetParameters().Count();
            //宣告組要執行的parameters字串
            object[] arr = new object[parameters];
            if (TestDriver == "AndroidDriver")
            {
                arr[0] = androidDriver;
            }
            else
            {
                arr[0] = driver;
            }
            for (int i = 1; i < parameters; i++)
            {
                arr[i] = item2.Inivalue[i - 1];
            }
            //執行方法
            stop = Convert.ToBoolean(magicMethod.Invoke(magicClassObject, arr));
            return stop;
        }


        public static bool InvokeStringMethod(string methodName, object[] stringParam)
        {
            TestMethodList dd = new TestMethodList();
            Type magicType = dd.GetType();
            ConstructorInfo magicConstructor = magicType.GetConstructor(Type.EmptyTypes);
            object magicClassObject = magicConstructor.Invoke(new object[] { });

            //MethodInfo magicMethod = magicType.GetMethod(methodName);
            MethodInfo magicMethod = magicType.GetMethods().Where(p =>p.Name==methodName && p.ToString().Contains("IWebDriver")).FirstOrDefault();
            var MethodEnd = magicMethod.Invoke(magicClassObject,  stringParam );
            return (bool)MethodEnd;
        }

        private static bool SwitchMethod(bool stop, Detail item2)
        {
            #region chrome
            switch (item2.IniName)
            {
                case "AssertAreEqual_DriverTitle":
                    stop = TestTool.AssertAreEqual_DriverTitle(driver, item2.Inivalue[0]);
                    break;
                case "AssertAreEqual_Text_ByXPath":
                    if (TestDriver.ToLower() == "android")
                    {
                        stop = TestTool.AssertAreEqual_Text_ByXPath(androidDriver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    else
                    {
                        stop = TestTool.AssertAreEqual_Text_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    break;
                case "AssertAreEqual_Text_ById":
                    if (TestDriver.ToLower() == "android")
                    {
                        stop = TestTool.AssertAreEqual_Text_ById(androidDriver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    else
                    {
                        stop = TestTool.AssertAreEqual_Text_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
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
                case "AssertAreEqual_Text_ByCssSelector":
                    stop = TestTool.AssertAreEqual_Text_ByCssSelector(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "AssertAreEqual_Text_ByXPath_TrimSpaceReplaceLine":
                    stop = TestTool.AssertAreEqual_Text_ByXPath_TrimSpaceReplaceLine(driver, item2.Inivalue[0], item2.Inivalue[1]);
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
                case "AssertAreEqual_ElementExist_ByLinkText":
                    stop = TestTool.AssertAreEqual_ElementExist_ByLinkText(driver, item2.Inivalue[0], item2.Inivalue[1]);
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
                case "AssertAreEqual_DropdownCount_ById":
                    stop = TestTool.AssertAreEqual_DropdownCount_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "AssertAreEqual_DropdownCount_ByXPath":
                    stop = TestTool.AssertAreEqual_DropdownCount_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "AssertAreEqual_DropdownCount_ByName":
                    stop = TestTool.AssertAreEqual_DropdownCount_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "AssertAreEqual_Android_messageToast_ByXPath":
                    stop = TestTool.AssertAreEqual_Android_messageToast_ByXPath(androidDriver, item2.Inivalue[0]);
                    break;
                case "AssertIsTrue_Selected_ByXPath":
                    stop = TestTool.AssertIsTrue_Selected_ByXPath(driver, item2.Inivalue[0]);
                    break;
                case "AssertIsTrue_DriverTitle":
                    stop = TestTool.AssertIsTrue_DriverTitle(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "Action_ContextClick_ByXPath":
                    stop = TestTool.Action_ContextClick_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "Action_MoveToElement_ByCssSelector":
                    stop = TestTool.Action_MoveToElement_ByCssSelector(driver, item2.Inivalue[0]);
                    break;
                case "Alert_OKCancel":
                    stop = TestTool.Alert_OKCancel(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "ClickFindElement_ById":
                    if (TestDriver.ToLower() == "android")
                    {
                        stop = TestTool.ClickFindElement_ById(androidDriver, item2.Inivalue[0]);
                    }
                    else
                    {
                        stop = TestTool.ClickFindElement_ById(driver, item2.Inivalue[0]);
                    }
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
                    stop = TestTool.PullDownScroll_ByXPath_ClickNextPage(driver, item2.Inivalue[0], item2.Inivalue[1], item2.Inivalue[2]);
                    break;
                case "PullRightScroll_ByXPath":
                    stop = TestTool.PullRightScroll_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "PullLeftScroll_ByXPatht":
                    stop = TestTool.PullLeftScroll_ByXPatht(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "SendKeys_ByXPath":
                    stop = TestTool.SendKeys_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "SendKeys_ById":
                    if (TestDriver.ToLower() == "android")
                    {
                        stop = TestTool.SendKeys_ById(androidDriver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    else
                    {
                        stop = TestTool.SendKeys_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    break;
                case "SendKeys_ByName":
                    stop = TestTool.SendKeys_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "SendKeys_SendWait":
                    stop = TestTool.SendKeys_SendWait(driver, item2.Inivalue[0]);
                    break;
                case "swipeToUp":
                    stop = TestTool.swipeToUp(androidDriver, item2.Inivalue[0]);
                    break;
                case "swipeToDown":
                    stop = TestTool.swipeToDown(androidDriver, item2.Inivalue[0]);
                    break;
                case "swipeToLeft":
                    stop = TestTool.swipeToDown(androidDriver, item2.Inivalue[0]);
                    break;
                case "swipeToRight":
                    stop = TestTool.swipeToDown(androidDriver, item2.Inivalue[0]);
                    break;
                case "ThreadSleep":
                    stop = TestTool.ThreadSleep(driver, item2.Inivalue[0]);
                    break;
                case "TestRememberMe":
                    driver = TestTool.TestRememberMe(driver, item2.Inivalue[0]);
                    break;
                case "Table_GetTrRowNumerAndClickHref_ByXPath":
                    stop = TestTool.Table_GetTrRowNumerAndClickHref_ByXPath(driver, item2.Inivalue[0], item2.Inivalue[1], item2.Inivalue[2], item2.Inivalue[3], item2.Inivalue[4]);
                    break;
                case "WebDriverWait_AssertAreEqual_Text_ById":
                    if (TestDriver.ToLower() == "android")
                    {
                        stop = TestTool.WebDriverWait_AssertAreEqual_Text_ById(androidDriver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    else
                    {
                        stop = TestTool.WebDriverWait_AssertAreEqual_Text_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    }
                    break;
                case "WebDriverWait_Dropdown_SelectText_ByName":
                    stop = TestTool.WebDriverWait_Dropdown_SelectText_ByName(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "WebDriverWait_Dropdown_SelectText_ById":
                    stop = TestTool.WebDriverWait_Dropdown_SelectText_ById(driver, item2.Inivalue[0], item2.Inivalue[1]);
                    break;
                case "TempTest":
                    stop = TestTool.TempTest(driver, item2.Inivalue[0], item2.Inivalue[1], item2.Inivalue[2]);
                    break;
                default:
                    TestTool.writeLog("error", item2.IniName + "：不存在!!");
                    break;
            }
            #endregion
            return stop;
        }
    }


}
