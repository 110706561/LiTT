using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Remote;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Drawing.Imaging;

namespace FormTestOtoCare
{
    class StartTest
    {
        public AndroidDriver<AndroidElement> driver;

        private string filelog = "UnitTest-" + DateTime.Now.ToString("(ddMMyyyy_HHmmss)") + ".log";
        private string LogDirectory = "C:\\COMPANY AAB\\UnitTestOtosales\\LogDirectory\\";
        private DesiredCapabilities capabilities;
        private string a;
        private string test = "";

        Capabilities capabilitesInvoke;

        public void AddLog(string strDesc)
        {
            try
            {
                bool Dir = Directory.Exists(LogDirectory);
                if (Dir == false)
                {
                    Directory.CreateDirectory(LogDirectory);
                }
                StreamWriter sw2 = new StreamWriter(LogDirectory + filelog, true, Encoding.ASCII);
                sw2.WriteLine(strDesc);
                sw2.Close();
            }

            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }
        }

        [TestInitialize]
        public void BeforeAll()
        {
            capabilitesInvoke = new Capabilities();
            capabilitesInvoke.CapabilitiesDevice();
        }

        [TestCleanup]
        public void AfterAll()
        {
            Capabilities.driver.Quit();

        }

        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void Scenario1()
        {
            Login login = new Login();
            login.LoginAction();
        }

        //public void Login(String usernameTxt, String passwordTxt)
        //{
        //    try
        //    {
        //        var username = driver.FindElementById("com.aab.mobilesalesnative:id/etUser");
        //        username.Click();

        //        //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
        //        username.SendKeys(usernameTxt);

        //        //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
        //        driver.HideKeyboard();

        //        var password = driver.FindElementById("com.aab.mobilesalesnative:id/etPassword");
        //        password.Click();
        //        //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
        //        password.SendKeys(passwordTxt);

        //        //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
        //        driver.HideKeyboard();

        //        var btnLogin = driver.FindElementById("com.aab.mobilesalesnative:id/btnLogin");
        //        btnLogin.Click();

        //        //File scrFile = ((ITakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
        //        a = "C:\\Company AAB\\UnitTestOtosales\\Screenshot\\images\\"+System.Reflection.MethodBase.GetCurrentMethod().Name.ToString() + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
        //        driver.GetScreenshot().SaveAsFile(a, ImageFormat.Jpeg);

        //        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
        //        wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//android.widget.ImageView[contains(@resource-id, 'android:id/up')]")));

        //        string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Success message : Success - " + System.Reflection.MethodBase.GetCurrentMethod().Name;
        //        AddLog(errorMessage);
        //        Trace.WriteLine(errorMessage);
        //    }

        //    catch(Exception ex)
        //    {
        //        string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Error message : " + System.Reflection.MethodBase.GetCurrentMethod().Name + " | " + ex.Message.ToString();
        //        AddLog(errorMessage);
        //        Trace.WriteLine(errorMessage);
        //    }
        //}

        public void TestPremiumSimulation()
        {
            try
            {
                //Vehicle Info
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));

                if (driver.FindElementById("com.aab.mobilesalesnative:id/drawerView").Displayed)
                {
                    var PremiumSimulation = driver.FindElementByXPath("//android.widget.ListView[@resource-id='com.aab.mobilesalesnative:id/drawerMenu']//android.widget.RelativeLayout[@index='0']");
                    PremiumSimulation.Click();

                    //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                }

                driver.FindElementById("com.aab.mobilesalesnative:id/etVCode").Click();

                //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                string b = "V05802";
                driver.FindElementById("com.aab.mobilesalesnative:id/etVCode").SendKeys(b);
                driver.HideKeyboard();
                driver.PressKeyCode(66);


                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));

                driver.FindElementById("com.aab.mobilesalesnative:id/spUsage").Click();
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));

                driver.FindElementByXPath("//android.widget.TextView[@text='Commercial Longterm']").Click();


                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementById("com.aab.mobilesalesnative:id/spRegion").Click();
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementByXPath("//android.widget.ListView[@resource-id='com.aab.mobilesalesnative:id/listSpinner']//android.widget.LinearLayout[@index='1']").Click();

                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementByXPath("//android.widget.TextView[@text='COVER']").Click();
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));

                a = "C:\\Company AAB\\UnitTestOtosales\\Screenshot\\images\\" + System.Reflection.MethodBase.GetCurrentMethod().Name.ToString() + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                driver.GetScreenshot().SaveAsFile(a, ImageFormat.Jpeg);

                //Cover
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementById("com.aab.mobilesalesnative:id/spBasicCover").Click();
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementByXPath("//android.widget.ListView[@resource-id='com.aab.mobilesalesnative:id/listSpinner']//android.widget.LinearLayout[@index='1']").Click();

                test = driver.FindElementById("com.aab.mobilesalesnative:id/etSumInsured").Text.ToString();
                //AddLog(driver.FindElementById("com.aab.mobilesalesnative:id/etSumInsured").Text);

                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementByXPath("//android.widget.TextView[@text='SUMMARY']").Click();
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));


                a = "C:\\Company AAB\\UnitTestOtosales\\Screenshot\\images\\" + System.Reflection.MethodBase.GetCurrentMethod().Name.ToString() + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                driver.GetScreenshot().SaveAsFile(a, ImageFormat.Jpeg);

                //Summary
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementByXPath("//android.widget.TextView[@resource-id='com.aab.mobilesalesnative:id/action_save_premium_simulation']").Click();
                a = "C:\\Company AAB\\UnitTestOtosales\\Screenshot\\images\\" + System.Reflection.MethodBase.GetCurrentMethod().Name.ToString() + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                driver.GetScreenshot().SaveAsFile(a, ImageFormat.Jpeg);

                //Assign Prospect
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementById("com.aab.mobilesalesnative:id/etProspectName").SendKeys("Alohadance");
                driver.FindElementById("com.aab.mobilesalesnative:id/etProspectPhone1").SendKeys("08123987123");

                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementById("com.aab.mobilesalesnative:id/spAssignAO").Click();

                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementByXPath("//android.widget.ListView[@resource-id='com.aab.mobilesalesnative:id/listSpinner']//android.widget.LinearLayout[@index='1']").Click();

                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                driver.FindElementById("com.aab.mobilesalesnative:id/action_save").Click();

                a = "C:\\Company AAB\\UnitTestOtosales\\Screenshot\\images\\" + System.Reflection.MethodBase.GetCurrentMethod().Name.ToString() + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                driver.GetScreenshot().SaveAsFile(a, ImageFormat.Jpeg);

                string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Success message : Success - " + System.Reflection.MethodBase.GetCurrentMethod().Name;
                AddLog(errorMessage);
                Trace.WriteLine(errorMessage);
            }

            catch (Exception ex)
            {
                string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Error message : " + System.Reflection.MethodBase.GetCurrentMethod().Name + " | " + ex.Message.ToString();
                AddLog(errorMessage);
                Trace.WriteLine(errorMessage);
            }
        }

        public void TaskList()
        {

        }

        public void InputProspect()
        {
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
            if (driver.FindElementById("com.aab.mobilesalesnative:id/drawerView").Displayed)
            {
                var InputProspect = driver.FindElementByXPath("//android.widget.ListView[@resource-id='com.aab.mobilesalesnative:id/drawerMenu']//android.widget.RelativeLayout[@index='1']");
                InputProspect.Click();

                //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
            }

            string prospectName = "Widuri";
            driver.FindElementById("com.aab.mobilesalesnative:id/etProspectName").SendKeys(prospectName);

            driver.FindElementById("com.aab.mobilesalesnative:id/btnBrowse1").Click();

            //element dari Redmi Note 2
            driver.FindElementByXPath("//android.widget.ListView[@resource-id='android:id/list']//android.view.View[@index='3']").Click();

            driver.FindElementById("com.aab.mobilesalesnative:id/ivKTP").Click();

            driver.FindElementByXPath("//android.widget.TextView[@text='Gallery']").Click();

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));

            driver.FindElementByXPath("//android.widget.GridView[@resource-id='android.miui:id/resolver_grid']//android.widget.LinearLayout[@index='1']").Click();

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));

            driver.Tap(1, 130, 2000, 1);

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
        }
    }
}
