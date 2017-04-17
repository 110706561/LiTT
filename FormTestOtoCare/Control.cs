using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.IO;
using System.Data;
using System.Threading;
using System.Data.SqlClient;

using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Principal;

using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Android.Interfaces;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.MultiTouch;

using Appium.Interfaces.Generic.SearchContext;
using Newtonsoft.Json;
using System.Collections.ObjectModel;
using System.Drawing;

using System.Text.RegularExpressions;


namespace FormTestOtoCare
{
    class Control 
    {
        public static string lastElement { get; set; }
        public static string lastFeatureAction { get; set; }
        public static string lastAction { get; set; }
        public static string dateTimeCreatedXls { get; set; }
        public static string actionPath { get; set; }
        public static int tempNum { get; set; }
        public static string insertText { get; set; }
        public static int successFlag { get; set; }
        public static string duration { get; set; }
        public static string actualValue { get; set; }
        //public static string desc { get; set; }

        public static Dictionary<string, string> myDictionary = new Dictionary<string, string>();
        //private static Dictionary<string, string> myDictionaries = new Dictionary<string, string>();
        private static string destinationJPG;

        //private string LogDirectory;// = "C:\\COMPANY AAB\\UnitTestOtosales\\LogDirectory\\";

        private string filelog = "UnitTest-" + DateTime.Now.ToString("(ddMMyyyy_HHmmss)") + ".log";
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";
        private static string serverUri1, serverUri2, serverUri3;
        private int errorFlag = 0;
        private int switchPortFlag = 0;
        private Configuration config;
        public static char[] delimiterChar = { '.' };
        public static int dummyIterator = 0;
        public Thread tid;

        private DateTime endTime;
        public static DateTime startTime {get; set; }

        public static AndroidDriver<AndroidElement> tempDriver2=null, tempDriver3=null;

        public Control()
        {
            config = new Configuration();
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            serverUri1 = ini.GetValue("serverUri1", "AndroidCapabilities");
            serverUri2 = ini.GetValue("serverUri2", "AndroidCapabilities");
            serverUri3 = ini.GetValue("serverUri3", "AndroidCapabilities");
        }

        public void _xCheckEnabledByElementId(string action, string elementId, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = "Check result element : " + elementId;
                lastAction = action;
                if (Capabilities.driver.FindElementById(elementId).GetAttribute("enabled").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                {
                    duration=DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("enabled").ToString();
                    successFlag = 0;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckEnabledByXPath(string action, string xPath, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = "Check result xpath : " + xPath;
                lastAction = action;
                if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("enabled").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("enabled").ToString();
                    successFlag = 0;
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckCheckedByElementId(string action, string elementId, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = "Check result element : " + elementId;
                lastAction = action;
                if (Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                {
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString();
                    successFlag = 0;
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckCheckedByXPath(string action, string xPath, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = "Check result xpath : " + xPath;
                lastAction = action;
                if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString();
                    successFlag = 0;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckFocusableByElementId(string action, string elementId, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = "Check result element : " + elementId;
                lastAction = action;
                if (Capabilities.driver.FindElementById(elementId).GetAttribute("focusable").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                {
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("focusable").ToString();
                    successFlag = 0;
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckFocusableByXPath(string action, string xPath, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = "Check result xpath : " + xPath;
                lastAction = action;
                if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("focusable").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("focusable").ToString();
                    successFlag = 0;
                    duration = DurationCalculate(startTime, endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xClickByElementId(string action, string elementId, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = elementId;
                lastAction = action;
                Capabilities.driver.FindElementById(elementId).Click();
                duration=DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xClickByXPath(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = xPath;
                lastAction = action;
                Capabilities.driver.FindElementByXPath(xPath).Click();
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckColorByXPath(string action, string xPath, string resultExpectation, int waitSeconds, string desc)
        {
             try
             {
                lastElement = "CheckColor";
                 lastAction = action;
                 destinationJPG = FormTest.pictureFile + lastFeatureAction + "-" + lastAction + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                 Capabilities.driver.GetScreenshot().SaveAsFile(destinationJPG, ImageFormat.Jpeg);
                 endTime = DateTime.Now;
                 duration=DurationCalculate(startTime,endTime);
                 config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
             }
             catch (Exception ex)
             {
                 _xCatchControl(ex.Message, desc);
             }
        }

        public void _xDismissByElementId(string action, string elementId, string type, int waitSeconds, string desc)
        {
            try
            {
                lastElement = elementId;
                lastAction = action;
                Capabilities.driver.PressKeyCode(187);
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Capabilities.driver.FindElementById(elementId).Click();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xDismissByXPath(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                lastElement = xPath;
                lastAction = action;
                Capabilities.driver.PressKeyCode(187);
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Capabilities.driver.FindElementByXPath(xPath).Click();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWriteByElementId(string action, string elementId, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = elementId;
                lastAction = action;
                //Capabilities.driver.FindElementById(elementId).Clear();
                //_xDeleteText(Capabilities.driver.FindElementById(elementId));
                Capabilities.driver.FindElementById(elementId).ReplaceValue(textSent);
                _xHideKeyboard();
                //_xPressKeyCode(66);
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWriteByXPath(string action, string xPath, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc,1);
                endTime = DateTime.Now;
                lastElement = xPath;
                lastAction = action;
                //Capabilities.driver.FindElementByXPath(xPath).Clear();
                Capabilities.driver.FindElementByXPath(xPath).ReplaceValue(textSent);
                //_xPressKeyCode(action, 66);
                _xHideKeyboard();
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWriteNumericByXPath(string action, string xPath, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = xPath;
                lastAction = action;
                Capabilities.driver.FindElementByXPath(xPath).Clear();
                int numKeys;
                for (int x = 0; x < textSent.Length; x++)
                {
                    numKeys = int.Parse(textSent.Substring(x, 1));
                    if (numKeys == 0) { Capabilities.driver.PressKeyCode(7); }
                    else if (numKeys == 1) { Capabilities.driver.PressKeyCode(8); }
                    else if (numKeys == 2) { Capabilities.driver.PressKeyCode(9); }
                    else if (numKeys == 3) { Capabilities.driver.PressKeyCode(10); }
                    else if (numKeys == 4) { Capabilities.driver.PressKeyCode(11); }
                    else if (numKeys == 5) { Capabilities.driver.PressKeyCode(12); }
                    else if (numKeys == 6) { Capabilities.driver.PressKeyCode(13); }
                    else if (numKeys == 7) { Capabilities.driver.PressKeyCode(14); }
                    else if (numKeys == 8) { Capabilities.driver.PressKeyCode(15); }
                    else if (numKeys == 9) { Capabilities.driver.PressKeyCode(16); }
                }
                _xHideKeyboard();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWriteNumericWithoutClearByXPath(string action, string xPath, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = xPath;
                lastAction = action;
                
                int numKeys;
                for (int x = 0; x < textSent.Length; x++)
                {
                    numKeys = int.Parse(textSent.Substring(x, 1));
                    if (numKeys == 0) { Capabilities.driver.PressKeyCode(7); }
                    else if (numKeys == 1) { Capabilities.driver.PressKeyCode(8); }
                    else if (numKeys == 2) { Capabilities.driver.PressKeyCode(9); }
                    else if (numKeys == 3) { Capabilities.driver.PressKeyCode(10); }
                    else if (numKeys == 4) { Capabilities.driver.PressKeyCode(11); }
                    else if (numKeys == 5) { Capabilities.driver.PressKeyCode(12); }
                    else if (numKeys == 6) { Capabilities.driver.PressKeyCode(13); }
                    else if (numKeys == 7) { Capabilities.driver.PressKeyCode(14); }
                    else if (numKeys == 8) { Capabilities.driver.PressKeyCode(15); }
                    else if (numKeys == 9) { Capabilities.driver.PressKeyCode(16); }
                }
                _xHideKeyboard();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWriteEmptySpaceByXPath(string action, string xPath, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastElement = xPath;
                lastAction = action;
                Capabilities.driver.FindElementByXPath(xPath).SendKeys(" ");
                _xHideKeyboard();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public static void _xHideKeyboard()
        {
            try
            {
                Capabilities.driver.HideKeyboard();
            }
            catch (Exception ex)
            { }
        }

        public void _xPressKeyCode(string action, int key)
        {
            try
            {
                Capabilities.driver.PressKeyCode(key);
                //lastElement = "";
                //lastAction = action;
                //lastAction = "";
            }
            catch (Exception ex)
            { 
                
            }
        }

        public static string DurationCalculate(DateTime startTime, DateTime endTime)
        {
            int milisecond, second, minute, duration;

            milisecond = (endTime - startTime).Milliseconds;
            second = (endTime - startTime).Seconds;
            minute = (endTime - startTime).Minutes;

            duration = (milisecond + (second * 1000) + (minute * 60000));

            return duration.ToString();
        }

        public void _xWaitUntil(string action, string nameId, int second, string type, string desc, int flag)
        {
            try
            {
                //startTime =DateTime.Now;
                WebDriverWait wDriverWait = new WebDriverWait(Capabilities.driver, TimeSpan.FromMilliseconds(second));
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "waiting for : " + nameId;
                switch (type.ToUpper())
                {
                    case "PATH": wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(nameId)));
                        break;
                    case "ELEMENT": wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id(nameId)));
                        break;
                }

                if (action.ToUpper().Equals("WAIT"))
                {
                    if (flag == 0)
                    {
                        duration = DurationCalculate(startTime,endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                    }   
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xNoActionSupported(string action, string element)
        {
            try
            {
                lastAction = action;
                lastElement = element;
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Not Supported", "No action triggered","","");
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xNoClassSupported(string classes)
        {
            try
            {
                lastAction = "";
                lastElement = "";
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Not Supported alternate class or element not ready", "Class not declared","","");
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xNoClassOrElement(string action)
        {
            try
            {
                lastAction = action;
                lastElement = "";
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Not supported alternate class or element not ready", "Class not declared or element not found", "", "");
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xDelay(string action, int msecond)
        {
            try
            {
                lastElement = "";
                lastAction = action;
                Thread.Sleep(msecond);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xScreenShot(string action, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastElement = "Screenshot";
                lastAction = action;
                destinationJPG = FormTest.pictureFile + lastFeatureAction + "-" + lastAction + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                Capabilities.driver.GetScreenshot().SaveAsFile(destinationJPG, ImageFormat.Jpeg);
                endTime = DateTime.Now;
                duration=DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckResultById(string action, string elementId, string resultExpectation, string ElementClass, string type, int waitSeconds, string desc)
        {
            //startTime =DateTime.Now;
            _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            lastElement = "Check result element : " + elementId;
            lastAction = action;

            switch (ElementClass.ToString().ToUpper())
            {
                #region check result action
                case "EDITTEXT":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Contains(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue +" and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "TEXTVIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Contains(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "CHECKBOX":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "VIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).GetAttribute("name").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("name").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "SWITCH":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "BUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Contains(resultExpectation))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "RADIOBUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Contains(resultExpectation))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                #endregion

                default:
                    _xCatchControl("class type not supported", "");
                    break;
            }
        }

        public void _xCheckResultByXPath(string action, string xPath, string resultExpectation, string ElementClass, string type, int waitSeconds, string desc)
        {
            //startTime =DateTime.Now;
            _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            lastElement = "Check result xpath : " + xPath;
            lastAction = action;

            switch (ElementClass.ToString().ToUpper())
            {
                #region check result action
                case "EDITTEXT":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Contains(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "TEXTVIEW":
                    try
                    {
                        string fixedStringOne = Regex.Replace(Capabilities.driver.FindElementByXPath(xPath).Text.ToString(), @"\s+", String.Empty).Replace("\n", String.Empty);
                        string fixedStringTwo = Regex.Replace(resultExpectation.ToString(), @"\s+", String.Empty).Replace("\n", String.Empty); ;

                        bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
                        //if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Contains(resultExpectation.ToString()))
                        if(isEqual)
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "CHECKBOX":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "VIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("name").ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("name").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "CHECKEDTEXTVIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Contains(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "SWITCH":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString().ToUpper().ToUpper().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "BUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Contains(resultExpectation))
                        {
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime,endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "RADIOBUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Contains(resultExpectation))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                #endregion

                default:
                    _xCatchControl("class type not supported", "");
                    break;

            }
        }

        public void _xCheckResultEmptyById(string action, string elementId, string resultExpectation, string ElementClass, string type, int waitSeconds, string desc)
        {
            //startTime = DateTime.Now;
            _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            lastElement = "Check result element : " + elementId;
            lastAction = action;

            switch (ElementClass.ToString().ToUpper())
            {
                #region check result action
                case "EDITTEXT":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "TEXTVIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "CHECKBOX":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "VIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).GetAttribute("name").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("name").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "SWITCH":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "BUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Equals(resultExpectation))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "RADIOBUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementById(elementId).Text.ToString().Equals(resultExpectation))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementById(elementId).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                #endregion

                default:
                    _xCatchControl("class type not supported", "");
                    break;
            }
        }

        public void _xCheckResultEmptyByXPath(string action, string xPath, string resultExpectation, string ElementClass, string type, int waitSeconds, string desc)
        {
            //startTime = DateTime.Now;
            _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            lastElement = "Check result xpath : " + xPath;
            lastAction = action;

            switch (ElementClass.ToString().ToUpper())
            {
                #region check result empty
                case "EDITTEXT":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "TEXTVIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "CHECKBOX":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                case "VIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("name").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("name").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "CHECKEDTEXTVIEW":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "SWITCH":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString().ToUpper().ToUpper().Equals(resultExpectation.ToString()))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).GetAttribute("checked").ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "BUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(resultExpectation))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;

                case "RADIOBUTTON":
                    try
                    {
                        if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(resultExpectation))
                        {
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                        }
                        else
                        {
                            actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                            successFlag = 0;
                            duration = DurationCalculate(startTime, endTime);
                            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                        }
                    }
                    catch (Exception ex)
                    {
                        _xCatchControl(ex.Message, desc);
                    }
                    break;
                #endregion

                default:
                    _xCatchControl("class type not supported", "");
                    break;

            }
        }

        public void _xCheckResultDateTime(string xPath, string action, string resultExpectation, string type, string desc)
        {
            try
            {
                System.Globalization.DateTimeFormatInfo convert = new System.Globalization.DateTimeFormatInfo();
                IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
                string[] temp = resultExpectation.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                string[] tempDateTime = temp[0].ToString().Split(new string[] { "&" }, StringSplitOptions.None);
                string[] tempValue = temp[2].ToString().Split(new string[] { "&" }, StringSplitOptions.None);
                DateTime setDateTime;
                if (temp[1].ToUpper().Equals("NOW"))
                    setDateTime = DateTime.Now;
                else
                    setDateTime = DateTime.Parse(temp[1], culture, System.Globalization.DateTimeStyles.AssumeLocal);
                lastAction = action;
                lastElement = xPath;

                if (temp[0].Contains("yyyy") || temp[0].Contains("MM") || temp[0].Contains("dd"))
                {
                    for (int k = 0; k < tempDateTime.Length; k++)
                    {
                        //tahun
                        if (tempDateTime[k].Equals("yyyy"))
                        {
                            setDateTime = setDateTime.AddYears(int.Parse(tempValue[k]));
                        }
                        //bulan
                        else if (tempDateTime[k].Equals("MM") || tempDateTime[k].Equals("MMM"))
                        {
                            setDateTime = setDateTime.AddMonths(int.Parse(tempValue[k]));
                        }
                        //tanggal
                        else if (tempDateTime[k].Equals("dd"))
                        {
                            setDateTime = setDateTime.AddDays(int.Parse(tempValue[k]));
                        }
                    }
                }
                else if (temp[0].Contains("HH") || temp[0].Contains("mm"))
                {
                    for (int j = 0; j < tempDateTime.Length; j++)
                    {
                        //jam
                        if (tempDateTime[j].Equals("HH"))
                        {
                            setDateTime = setDateTime.AddHours(int.Parse(tempValue[j]));
                        }
                        //menit
                        else if (tempDateTime[j].Equals("mm"))
                        {
                            setDateTime = setDateTime.AddMinutes(int.Parse(tempValue[j]));
                        }
                    }
                }

                if (resultExpectation != null && resultExpectation.Length > 0)
                {
                    if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(setDateTime.ToString(temp[3])))
                    {
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", "Not Use", desc);
                    }
                    else
                    {
                        actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                        successFlag = 0;
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + setDateTime.ToString(temp[3]), "Not Use", desc);
                    }
                }
                else
                {
                    if (Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Equals(setDateTime.ToString(temp[3])))
                    {
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", "Not Use", desc);
                    }
                    else
                    {
                        actualValue = Capabilities.driver.FindElementByXPath(xPath).Text.ToString();
                        successFlag = 0;
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + setDateTime.ToString(temp[3]), "Not Use", desc);
                    }
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCatchControl(string message, string desc)
        {
            if (message.Contains("not found"))
            {
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Not Found", message.ToString(),"","");
            }
            else
            {
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Error", message.ToString(),"","");
            }

            string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Error message : " + lastFeatureAction + " | " + lastElement + "|" + message;
            //_xAddLog(errorMessage);
            errorFlag = 1;
            successFlag = 0;
        }

        public void _xSuccessControl()
        {
            if (errorFlag != 1)
            {
                string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Success message : Success - " + lastFeatureAction;
                //_xAddLog(errorMessage);
            }
        }

        private bool IsAllDigits(string s)
        {
            foreach (char c in s)
            {
                if (!char.IsDigit(c))
                    return false;
            }
            return true;
        }

        public string _xComposeString(DataRow data, int b)
        {
            string path = "";
            if (data.Table.Rows[b]["Action"].ToString().ToUpper().Equals("SORT"))
            {
                path = "//" + data.Table.Rows[b]["Class"].ToString();
            }
            else
            {
                switch (data.Table.Rows[b]["Class"].ToString().ToUpper().Split(delimiterChar)[2].ToString())
                {
                    case "TEXTVIEW":
                        if (data.Table.Rows[b]["Key"].ToString()[0].Equals('^'))
                            path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@text='" + data.Table.Rows[b]["Key"].ToString() + "']";
                        else
                        {
                            if (IsAllDigits(data.Table.Rows[b]["Key"].ToString()) == true)//cek semua key isinya angka
                                path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Key"].ToString() + "']";
                            else
                                path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@text='" + data.Table.Rows[b]["Key"].ToString() + "']";
                        }
                        break;
                    case "EDITTEXT": path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                        break;
                    case "CHECKEDTEXTVIEW":
                        if (data.Table.Rows[b]["ResourceId"].ToString() == "")
                        {
                            if (IsAllDigits( data.Table.Rows[b]["Key"].ToString()) == true)//cek semua key isinya angka
                                path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Key"].ToString() + "']";
                            else
                                path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@text='" + data.Table.Rows[b]["Key"].ToString() + "']";
                        }
                        else
                            path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";

                        break;
                    case "BUTTON":
                        if (data.Table.Rows[b]["ResourceId"].ToString() == "")
                        {
                            path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@text='" + data.Table.Rows[b]["Key"].ToString() + "']";
                        }
                        else
                        {
                            path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                        }

                        break;
                    //case ""
                    #region Buildpath Rule
                    /* BUILDPATH RULE
                case "LISTVIEW":
                    if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                    else
                        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                    break;
                //case "LINEARLAYOUT"     : 
                //    if(data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "RELATIVELAYOUT"   : 
                //    if(data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "NUMBERPICKER"     : 
                //    if(data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "IMAGEVIEW"        : 
                //    if(data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "FRAMELAYOUT":
                //    if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "WEBVIEW":
                //    if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "VIEW":
                //    if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "CHECKEDTEXTVIEW":
                //    if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                //case "IMAGEBUTTON":
                //    if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                //    else
                //        path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Value"].ToString() + "']";
                //    break;
                 */
                    #endregion
                    default:
                        if (data.Table.Rows[b]["ResourceId"].ToString().Length > 2)
                            path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@resource-id='" + data.Table.Rows[b]["ResourceId"].ToString() + "']";
                        else
                            path = "//" + data.Table.Rows[b]["Class"].ToString() + "[@index='" + data.Table.Rows[b]["Key"].ToString() + "']";
                        break;
                }
            }
            return path;
        }

        public string _xCheckBuildPath(DataRow data, int b, bool isCheckBaseAlternate)
        {
            string tempText = "";
            if (AndroidAction.flagAlternateMaster == 0)
            {
                return AndroidAction.unaclass;
            }
            else
            {
                if (b > 0)
                {

                    if (data.Table.Rows[b - 1]["Action"].ToString().ToUpper().Equals("BUILDPATH"))
                    {
                        tempText = Control.actionPath + _xComposeString(data, b);
                        if (!isCheckBaseAlternate)
                            Control.actionPath = "";
                    }
                    else
                    {
                        tempText = _xComposeString(data, b);
                    }
                }

                else
                {
                    tempText = _xComposeString(data, b);
                }
                return tempText;
            }
        }

        public void _xSwipe(string action, string timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = "";
                string[] timeExec = timeExecution.Split(',');
                if (timeExec.Length == 5)
                {
                    Capabilities.driver.Swipe(Convert.ToInt32(timeExec[0]), Convert.ToInt32(timeExec[1]), Convert.ToInt32(timeExec[2]), Convert.ToInt32(timeExec[3]), Convert.ToInt32(timeExec[4]));
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Wrong format value, please input with format value 'start x, start y, end x, end y, interval time'",duration,desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeUp(string action, int timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Thread.Sleep(2000);
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Swipe(Capabilities.driver.Manage().Window.Size.Width / 2, Capabilities.driver.Manage().Window.Size.Height * 4 / 5, Capabilities.driver.Manage().Window.Size.Width / 2, Capabilities.driver.Manage().Window.Size.Height / 5, timeExecution);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
            //Thread.Sleep(2000);
        }

        public void _xSwipeDown(string action, int timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Thread.Sleep(2000);
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Swipe(Capabilities.driver.Manage().Window.Size.Width / 2, Capabilities.driver.Manage().Window.Size.Height / 5, Capabilities.driver.Manage().Window.Size.Width / 2, Capabilities.driver.Manage().Window.Size.Height * 4 / 5, timeExecution);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
            //Thread.Sleep(2000);
        }

        public void _xSwipeHorizontalBar(string action,string xPath, int value, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Thread.Sleep(2000);
                AndroidElement target = Capabilities.driver.FindElementByXPath(xPath);
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Swipe(target.Location.X, target.Location.Y + (target.Size.Height / 2), target.Location.X + target.Size.Width / 100 * value, target.Location.Y + (target.Size.Height / 2),1000);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeVerticalBar(string action, string xPath, int value, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Thread.Sleep(2000);
                AndroidElement target = Capabilities.driver.FindElementByXPath(xPath);
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Swipe(target.Location.X, target.Location.Y + (target.Size.Height / 2), target.Location.X , (target.Location.Y + (target.Size.Height / 2)) + target.Size.Width / 100 * value, 1000);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeRight(string action, int timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Thread.Sleep(2000);
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Swipe(Capabilities.driver.Manage().Window.Size.Width * 7 / 10, Capabilities.driver.Manage().Window.Size.Height / 2, Capabilities.driver.Manage().Window.Size.Width * 3 / 10, Capabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }

            //Thread.Sleep(2000);
        }

        public void _xSwipeLeft(string action, int timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Thread.Sleep(2000);
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Swipe(Capabilities.driver.Manage().Window.Size.Width * 3 / 10, Capabilities.driver.Manage().Window.Size.Height / 2, Capabilities.driver.Manage().Window.Size.Width * 7 / 10, Capabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
            //Thread.Sleep(2000);
        }

        public void _xDateSwipeUp(string value)
        {
            try
            {
                int startX = Capabilities.driver.FindElementByXPath(Control.actionPath).Coordinates.LocationInViewport.X + 20;
                int startY = Capabilities.driver.FindElementByXPath(Control.actionPath).Coordinates.LocationInViewport.Y + 250;
                Capabilities.driver.Swipe(startX, startY, startX, startY - 200, Convert.ToInt32(value));
                Control.actionPath = "";
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xDateSwipeDown(string value)
        {
            try
            {
                int startX = Capabilities.driver.FindElementByXPath(Control.actionPath).Coordinates.LocationInViewport.X + 20;
                int startY = Capabilities.driver.FindElementByXPath(Control.actionPath).Coordinates.LocationInViewport.Y + 250;
                Capabilities.driver.Swipe(startX, startY, startX, startY + 200, Convert.ToInt32(value));
                Control.actionPath = "";
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xCheckExistByElementId(string action, string elementId, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = elementId;
                if (Capabilities.driver.FindElementById(elementId).Displayed)
                {
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    successFlag = 0;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Element " + elementId + " not found",duration,desc);
                }
            }
            catch (Exception ex)
            {
                successFlag = 0;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Element " + elementId + " not found", duration, desc);
            }
        }

        public void _xCheckExistByXpath(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = xPath;
                if (Capabilities.driver.FindElementByXPath(xPath).Displayed)
                {
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    successFlag = 0;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Element " + xPath + " not found", duration, desc);
                }
            }
            catch (Exception ex)
            {
                successFlag = 0;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Element " + xPath + " not found", duration,desc);
            }
        }

        public void _xGetByElementId(string action, string elementId, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = elementId;
                myDictionary.Add(value, Capabilities.driver.FindElementById(elementId).Text.ToString());
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xGetElementByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = xPath;
                myDictionary.Add(value, Capabilities.driver.FindElementByXPath(xPath).Text.ToString());
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xGetElementByXPathSubstring(string action, string xPath, string ResultVariable, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = xPath;
                string[] teks = value.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                myDictionary.Add(ResultVariable, Capabilities.driver.FindElementByXPath(xPath).Text.ToString().Substring(int.Parse(teks[0]), int.Parse(teks[1])));
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xPostByElementId(string action, string elementId, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc ,1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = elementId;
                _xWriteByElementId(action, elementId, myDictionary[value].ToString(), type, waitSeconds, desc);
                duration=DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                //myDictionaries[value].Tables[0].Rows[b][""]
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xPostByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc , 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = xPath;
                _xWriteByElementId(action, xPath, myDictionary[value].ToString(), type, waitSeconds, desc);
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);                
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xBackDriver(string action, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastElement = "";
                lastAction = action;
                Thread.Sleep(5000);
                Capabilities.driver.PressKeyCode(4);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                duration = (int.Parse(duration) - 5000).ToString();
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }

        }

        public void _xSwipeNotificationBar(string action, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = "";
                Capabilities.driver.OpenNotifications();
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }

        }

        public void _xCheckToast(string action, string value)
        {
            try
            {
                lastElement = "";
                lastAction = action;
                String toastdestinationJPG = FormTest.pictureFile + "Toast-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                Capabilities.driver.GetScreenshot().SaveAsFile(toastdestinationJPG, ImageFormat.Jpeg);
                Thread.Sleep(1000);
            } 
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xDeleteTextByElementId(string elementId, string action, int value, string type, int waitSeconds, string desc)
        {
            try
            {
                _xWaitUntil(action, elementId, waitSeconds, type, "", 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                AndroidElement target = Capabilities.driver.FindElementById(elementId);
                TouchAction touch = new TouchAction(Capabilities.driver);
                touch.Press(target.Location.X + target.Size.Width - 5, target.Location.Y + (target.Size.Height / 2)).Release();
                touch.Perform();
                for (int a = 0; a < value; a++)
                    Capabilities.driver.LongPressKeyCode(67);
                _xHideKeyboard();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xDeleteTextByXPath(string xPath, string action, int value, string type, int waitSeconds, string desc)
        {
            try
            {
                _xWaitUntil(action, xPath, waitSeconds, type, "", 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                AndroidElement target = Capabilities.driver.FindElementByXPath(xPath);
                TouchAction touch = new TouchAction(Capabilities.driver);
                touch.Press(target.Location.X + target.Size.Width - 5, target.Location.Y + (target.Size.Height / 2)).Release();
                touch.Perform();
                for (int a = 0; a < value; a++)
                    Capabilities.driver.LongPressKeyCode(67);
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xClearTextByXPath(string xPath, string action, string type, int waitSeconds, string desc)
        {
            try
            {
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                _xWaitUntil(action, xPath, waitSeconds, type, "", 1);
                Capabilities.driver.FindElementByXPath(xPath).Clear();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xDeleteNumeric(string action, int value, string type, int waitSeconds, string desc)
        {
            try
            {
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                for (int a = 0; a < value; a++)
                {
                   Capabilities.driver.PressKeyCode(67);
                }

                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xNetworkChange(string action, string value)
        {
            try
            {
                lastElement = "";
                lastAction = action;
                switch (value.ToUpper())
                {
                    case "AIRPLANE": Capabilities.driver.ConnectionType = ConnectionType.AirplaneMode;
                        break;
                    case "ALLNETWORK": Capabilities.driver.ConnectionType = ConnectionType.AllNetworkOn;
                        break;
                    case "DATAONLY": Capabilities.driver.ConnectionType = ConnectionType.DataOnly;
                        break;
                    case "WIFIONLY": Capabilities.driver.ConnectionType = ConnectionType.WifiOnly;
                        break;
                    case "NONE": Capabilities.driver.ConnectionType = ConnectionType.None;
                        break;
                    default: _xNoActionSupported(action, value);
                        break;
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xLongPressByElementId(string action, string elementId, int durationx, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now; 
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Tap(1, Capabilities.driver.FindElementById(elementId), durationx);
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xLongPressByXPath(string action, string xPath, int durationx, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Tap(1, Capabilities.driver.FindElementByXPath(xPath), durationx);
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xZoomByElementId(string action, string elementId, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Zoom(Capabilities.driver.FindElementById(elementId));
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xZoomByXPath(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Zoom(Capabilities.driver.FindElementByXPath(xPath));
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xTap(string action, int X, int Y, int durationx, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = "";
                Capabilities.driver.Tap(1, X, Y, durationx);
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckSorted(string action, string elementId, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = elementId;
                IReadOnlyCollection<AndroidElement> nameOfList = Capabilities.driver.FindElementsByXPath(elementId);
                if (value.ToUpper().ToString().Equals("ASCENDING"))
                {
                    if (_xIsSortedAscending(nameOfList))
                    {
                        duration = DurationCalculate(startTime,endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "List sorted ascending",duration,desc);
                    }
                    else
                    {
                        successFlag = 0;
                        duration = DurationCalculate(startTime,endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else
                {
                    if (_xIsSortedDescending(nameOfList))
                    {
                        duration = DurationCalculate(startTime,endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "List sorted descending", duration,desc);
                    }
                    else
                    {
                        successFlag = 0;
                        duration = DurationCalculate(startTime,endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "List not sorted descending",duration,desc);
                    }
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckSorted2(string action, string elementId, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = elementId;
                IReadOnlyCollection<AndroidElement> nameOfList = Capabilities.driver.FindElementsByXPath(elementId);
                if (value.ToUpper().ToString().Equals("ASCENDING"))
                {
                    if (_xIsSortedAscending(nameOfList))
                    {
                        duration = DurationCalculate(startTime, endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "List sorted ascending", duration, desc);
                    }
                    else
                    {
                        successFlag = 0;
                        duration = DurationCalculate(startTime, endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else
                {
                    if (_xIsSortedDescending(nameOfList))
                    {
                        duration = DurationCalculate(startTime, endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "List sorted descending", duration, desc);
                    }
                    else
                    {
                        successFlag = 0;
                        duration = DurationCalculate(startTime, endTime);
                        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "List not sorted descending", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public static bool _xIsSortedAscending(IReadOnlyCollection<AndroidElement> nameOfList)
        {
            for (int i = 1; i < nameOfList.Count; i++)
            {
                if (nameOfList.ElementAt(i - 1).Text != nameOfList.ElementAt(i).Text)
                {
                    if (string.Compare(nameOfList.ElementAt(i - 1).Text, nameOfList.ElementAt(i).Text) > 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public static bool _xIsSortedDescending(IReadOnlyCollection<AndroidElement> nameOfList)
        {
            for (int i = 1; i < nameOfList.Count; i++)
            {
                if (nameOfList.ElementAt(i - 1).Text != nameOfList.ElementAt(i).Text)
                {
                    if (string.Compare(nameOfList.ElementAt(i - 1).Text, nameOfList.ElementAt(i).Text) < 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public void _xCheckUserAD(string action, string userpass, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = "";
                string[] up = userpass.Split(';');
                DirectoryEntry root = new DirectoryEntry("LDAP://beyond.asuransi.astra.co.id", up[0].ToString(), up[1].ToString(), AuthenticationTypes.Secure);

                DirectorySearcher dSearch = new DirectorySearcher(root);
                SearchResultCollection dResult = dSearch.FindAll();
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "Username and password are correct",duration,desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }

            #region try AD
            //if (dResult.Count > 0)
            //{
            //    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "Username and password are correct");
            //}
            //else
            //{

            //}
            //using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "beyond.asuransi.astra.co.id"))
            //{
            //    DirectoryEntry root = new DirectoryEntry(
            //    // validate the credentials
            //    string[] up = userpass.Split(';');
            //    bool isValid = pc.ValidateCredentials(up[0].ToString(), up[1].ToString());
            //    if (isValid)
            //    {
            //        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "Username and password are correct");
            //    }
            //    else
            //    {
            //        config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Username and password are incorrect");
            //    }
            //}
            #endregion
        }

        public void _xCheckPortrait(string action, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = "";
                ScreenOrientation so;
                switch (resultExpectation.ToUpper())
                {
                    case "PORTRAIT":
                        so = ScreenOrientation.Portrait;
                        break;
                    case "LANDSCAPE":
                        so = ScreenOrientation.Landscape;
                        break;
                    default:
                        so = ScreenOrientation.Portrait;
                        break;
                }

                if (Capabilities.driver.Orientation.Equals(so))
                {
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
                else
                {
                    successFlag = 0;
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Screen orientation not match", duration,desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public static bool _xCheckBaseClass(string classes, int waitSeconds, string type)
        {
            try
            {
                if (type.ToUpper().Equals("ELEMENT"))
                {
                    WebDriverWait wDriverWait = new WebDriverWait(Capabilities.driver, TimeSpan.FromMilliseconds(waitSeconds));
                    wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id(classes)));

                    //if (Capabilities.driver.FindElementByClassName(classes).Displayed)
                    if (Capabilities.driver.FindElementById(classes).Displayed)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    WebDriverWait wDriverWait = new WebDriverWait(Capabilities.driver, TimeSpan.FromMilliseconds(waitSeconds));
                    wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(classes)));

                    //if (Capabilities.driver.FindElementByClassName(classes).Displayed)
                    if (Capabilities.driver.FindElementByXPath(classes).Displayed)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

            }
            catch (Exception ex)
            {
                return false;
            }
        }

        #region Rate (not used)
        public void _xRateByElementId(string action, string elementId, int value, string type, int waitSeconds)
        {
            try
            {
                _xWaitUntil(action, elementId, waitSeconds, type, "", 1);
                AndroidElement target = Capabilities.driver.FindElementById(elementId);
                TouchAction touch = new TouchAction(Capabilities.driver);
                touch.Press(target.Location.X + target.Size.Width / 10 * ((2 * value) - 1), target.Location.Y + (target.Size.Height / 2)).Release();
                touch.Perform();
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xRateByXPath(string action, string xPath, int value, string type, int waitSeconds)
        {
            try
            {
                _xWaitUntil(action, xPath, waitSeconds, type, "", 1);
                AndroidElement target = Capabilities.driver.FindElementById(xPath);
                TouchAction touch = new TouchAction(Capabilities.driver);
                touch.Press(target.Location.X + target.Size.Width / 10 * ((2 * value) - 1), target.Location.Y + (target.Size.Height / 2)).Release();
                touch.Perform();
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }
        #endregion

        public void _xCheckNotExistByElementId(string action, string elementId, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = elementId;
                if (Capabilities.driver.FindElementById(elementId).Displayed)
                {
                    successFlag = 0;
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Element " + elementId + " found", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
            }
            catch (Exception ex)
            {
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
        }

        public void _xCheckNotExistByXPath(string action, string xPath, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                lastAction = action;
                lastElement = xPath;
                if (Capabilities.driver.FindElementByXPath(xPath).Displayed)
                {
                    successFlag = 0;
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Failed", "Element " + xPath + " found", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = DurationCalculate(startTime,endTime);
                    config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
                }
            }
            catch (Exception ex)
            {
                endTime = DateTime.Now;
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
        }

        public void _xUpdateDateTimeByXPath(string xPath, string action, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                _xDelay(lastAction, 1000);
                Capabilities.driver.Tap(1, Capabilities.driver.FindElementByXPath(xPath), 1000);
                _xDelay(lastAction, 500);
                //Capabilities.driver.FindElementByXPath(xPath).Clear();
                Capabilities.driver.FindElementByXPath(xPath).ReplaceValue(value);
                _xDelay(lastAction, 500);
                Capabilities.driver.PressKeyCode(66);
                _xHideKeyboard();
                duration = DurationCalculate(startTime, endTime);
                lastAction = action;
                lastElement = xPath;
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xUpdateDateTimeWithIntervalByXPath(string xPath, string action, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                System.Globalization.DateTimeFormatInfo convert = new System.Globalization.DateTimeFormatInfo();
                DateTime TimeNow = DateTime.Now;
                startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                lastAction = action;
                lastElement = xPath;
                Capabilities.driver.Tap(1, Capabilities.driver.FindElementByXPath(xPath), 1000);
                Capabilities.driver.FindElementByXPath(xPath).Clear();

                if (value.ToString()[0].Equals('H'))
                    value = TimeNow.AddHours(double.Parse(value.Substring(2))).Hour.ToString();
                else if (value.Substring(0, 2).Equals("MI"))
                    value = TimeNow.AddMinutes(double.Parse(value.Substring(2))).Minute.ToString();
                else if (value.Substring(0, 4).Equals("YYYY"))
                    value = TimeNow.AddYears(int.Parse(value.Substring(5))).Year.ToString();
                else if (value.Substring(0, 2).Equals("DD"))
                    value = TimeNow.AddDays(double.Parse(value.Substring(3))).Day.ToString();
                else if (value.Substring(0, 2).Equals("MO"))
                    value = convert.GetMonthName(TimeNow.AddMonths(int.Parse(value.Substring(3))).Month).Substring(0, 3);

                Capabilities.driver.FindElementByXPath(xPath).SendKeys(value);
                Capabilities.driver.PressKeyCode(66);
                _xHideKeyboard();
                duration = DurationCalculate(startTime, endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xSetDateTime(string action, string Value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                System.Globalization.DateTimeFormatInfo convert = new System.Globalization.DateTimeFormatInfo();
                IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
                string[] temp = Value.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                string[] tempDateTime = temp[0].ToString().Split(new string[] { "&" }, StringSplitOptions.None);
                string[] tempValue = temp[2].ToString().Split(new string[] { "&" }, StringSplitOptions.None);
                DateTime setDateTime;
                if (temp[1].ToUpper().Equals("NOW"))
                    setDateTime = DateTime.Now;
                else
                    setDateTime = DateTime.Parse(temp[1], culture, System.Globalization.DateTimeStyles.AssumeLocal);
                string[] setDate = new string[] { "'2'", "'0'", "'1'", "'2'"};
                string[] setTime = new string[] { "'0'", "'2'" };
                string[] value = new string[4];
                string buttonOkPath = "", buttonCancelPath = "";
                
                if (temp[0].Contains("yyyy") || temp[0].Contains("MM") || temp[0].Contains("dd"))
                {
                    for (int k = 0; k < tempDateTime.Length; k++)
                    {
                        //tahun
                        if (tempDateTime[k].Equals("yyyy"))
                        {
                            setDateTime = setDateTime.AddYears(int.Parse(tempValue[k]));
                        }
                        //bulan
                        else if (tempDateTime[k].Equals("MM") || tempDateTime[k].Equals("MMM"))
                        {
                            setDateTime = setDateTime.AddMonths(int.Parse(tempValue[k]));
                        }
                        //tanggal
                        else if (tempDateTime[k].Equals("dd"))
                        {
                            setDateTime = setDateTime.AddDays(int.Parse(tempValue[k]));
                        }
                    }

                    if (setDateTime.Month == 12 || setDateTime.Month == 1)
                    {
                        value[0] = (setDateTime.Year - 1).ToString();
                        //Code dibawah ditambahkan karena problem kalender untuk bulan Jan-Dec/Dec-Jan pada OS Kitkat
                        value[3] = setDateTime.Year.ToString();
                    }
                    else
                    {
                        value[0] = setDateTime.Year.ToString();
                        //Code dibawah ditambahkan karena problem kalender untuk bulan Jan-Dec/Dec-Jan pada OS Kitkat
                        value[3] = setDateTime.Year.ToString();
                    }
                    value[1] = convert.GetMonthName(setDateTime.Month).Substring(0, 3);
                    value[2] = setDateTime.Day.ToString();

                    DataRow[] dr = FormTest.dsMasterResourceId.Tables[0].Select("OperatingSystem = '" + Capabilities.osVersion + "'" + " AND Modul like '%calendar%'");
                    switch(dr[0]["Modul"].ToString().ToUpper())
                    {
                        case  "SMALLCALENDAR":
                        //SetDateTime untuk OS 4(KitKat)
                        string tempX, tempCheck;
                        bool flagCheckSmallCalendar = true;
                        for (int i = 0; i < setDate.Length; i++)
                        {
                            for (int j = 0; j < dr.Length; j++)
                            {
                                if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONOK"))
			                    buttonOkPath = dr[j]["AndroidResourceId"].ToString();

		                        else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONCANCEL"))
			                    buttonCancelPath = dr[j]["AndroidResourceId"].ToString();
                            }
                            for (int h = 0; h < 5; h++)
                            {
                                string tempXPath = dr[0]["AndroidResourceId"].ToString();
                                tempXPath = tempXPath.Replace("#Key", setDate[i]);
                                _xWaitUntil(action, tempXPath, waitSeconds, type, desc, 1);
                                if (i == 0 || i == 2 || i == 3)
                                {
                                    tempX = int.Parse(value[i]).ToString();
                                    tempCheck = int.Parse(Capabilities.driver.FindElementByXPath(tempXPath).Text).ToString();
                                }
                                else
                                {
                                    tempX = value[i].ToString();
                                    tempCheck = Capabilities.driver.FindElementByXPath(tempXPath).Text;
                                }

                                if (!tempX.Equals(tempCheck))
                                {
                                    _xDelay(lastAction, 1000);
                                    Capabilities.driver.Tap(1, Capabilities.driver.FindElementByXPath(tempXPath), 1000);
                                    _xDelay(lastAction, 500);
                                    Capabilities.driver.FindElementByXPath(tempXPath).ReplaceValue(value[i]);
                                    _xDelay(lastAction, 500);
                                    Capabilities.driver.PressKeyCode(66);
                                    _xHideKeyboard();
                                }
                                else if(i==3 && tempX.Equals(tempCheck))
                                {
                                    flagCheckSmallCalendar = false;
                                    break;
                                }
                            }
                            if (!flagCheckSmallCalendar)
                                Capabilities.driver.FindElementByXPath(buttonOkPath).Click();
                        }
                        break;
                        //Batas SetDateTime untuk OS 4(KitKat)

                        case "BIGCALENDAR":
                        //SetDateTime untuk OS 5(Lollipop)
                        string year="",month="",day="",listViewDateMonthPath="",listviewYearPath="";
                        int yearIndex=0;
                        for (int i = 0; i < dr.Length; i++)
                        {
                            //inisialisasi value awal dari year,month,day,listViewDateMonth,listviewYear
                            for(int j=0;j<dr.Length;j++)
                            {
                                if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("YEAR"))
                                {
                                    year = dr[j]["AndroidResourceId"].ToString();
                                    yearIndex = j;
                                }
                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("MONTH"))
                                    month = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("DAY"))
                                    day = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("LISTVIEWDATEMONTH"))
                                    listViewDateMonthPath = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("LISTVIEWYEAR"))
                                    listviewYearPath = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONOK"))
                                    buttonOkPath = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONCANCEL"))
                                    buttonCancelPath = dr[j]["AndroidResourceId"].ToString();
                            }

                            DateTime dateExist, dateTo;
                            dateExist = DateTime.Parse(Capabilities.driver.FindElementByXPath(day).Text + " " + Capabilities.driver.FindElementByXPath(month).Text + " " + Capabilities.driver.FindElementByXPath(year).Text);
                            dateTo = DateTime.Parse(value[2] + " " + value[1] + " " + value[3]);

                            //setYear
                            Capabilities.driver.FindElementByXPath(dr[yearIndex]["AndroidResourceId"].ToString()).Click();
                            IReadOnlyCollection<AndroidElement> nameOfListYear = Capabilities.driver.FindElementsByXPath(listviewYearPath);
                            AndroidElement targetYear = Capabilities.driver.FindElementByXPath(listViewDateMonthPath);
                            bool flagCheckYear = true;
                            while (flagCheckYear)
                            {
                                if (Capabilities.driver.FindElementByXPath(listviewYearPath + "[@index='0']").Text.Equals(value[3]))
                                {
                                    _xDelay("", 500);
                                    Capabilities.driver.FindElementByXPath(listviewYearPath + "[@index='0']").Click();
                                    flagCheckYear = false;
                                    break;
                                }
                                else
                                {
                                    for (int z = 1; z < nameOfListYear.Count; z++)
                                    {
                                        _xDelay("", 500);
                                        //if (nameOfListYear.ElementAt(z).Text.Equals(value[3]))
                                        try
                                        {
                                            if (Capabilities.driver.FindElementByXPath(listviewYearPath + "[@index='" + z + "']").Text.Equals(value[3]))
                                            {
                                                //klik tahun pada element z
                                                Capabilities.driver.FindElementByXPath(listviewYearPath + "[@index='" + z + "']").Click();
                                                flagCheckYear = false;
                                                break;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            break;
                                        }
                                    }
                                }

                                if(dateExist.Year>dateTo.Year && flagCheckYear==true)
                                {
                                    //Swipe dari atas ke bawah
                                    Capabilities.driver.Swipe(targetYear.Location.X + (targetYear.Size.Width / 4), targetYear.Location.Y + 5, targetYear.Location.X + (targetYear.Size.Width / 4), (targetYear.Location.Y + targetYear.Size.Height) - 5,int.Parse(Capabilities.speedSwipe));
                                    _xDelay("", 1000);
                                }
                                else if (dateExist.Year < dateTo.Year && flagCheckYear == true)
                                {
                                    //Swipe dari atas ke bawah
                                    Capabilities.driver.Swipe(targetYear.Location.X + (targetYear.Size.Width / 4), (targetYear.Location.Y + targetYear.Size.Height) - 5, targetYear.Location.X + (targetYear.Size.Width / 4), targetYear.Location.Y + 5, int.Parse(Capabilities.speedSwipe));
                                    _xDelay("", 1000);
                                }
                                else if (dateExist.Year == dateTo.Year)
                                {
                                    _xDelay("", 500);
                                    //Capabilities.driver.FindElementByXPath(listviewYearPath + "[@text='"+dateExist.Year.ToString()+"']").Click();
                                    break;
                                }
                            }

                            //SetDateMonth
                            IReadOnlyCollection<AndroidElement> nameOfListDateMonth = Capabilities.driver.FindElementsByXPath(listViewDateMonthPath);
                            AndroidElement targetDateMonth = Capabilities.driver.FindElementByXPath(listViewDateMonthPath);
                            bool flagCheckDateMonth = true;
                            
                            while (flagCheckDateMonth)
                            {
                                if (Capabilities.driver.FindElementsByXPath(listViewDateMonthPath + "//android.view.View[@index='0']/*").Count > 1)
                                {
                                    IReadOnlyCollection<AndroidElement> nameOfList = Capabilities.driver.FindElementsByXPath(listViewDateMonthPath + "//android.view.View[@index='0']/*");
                                    if (nameOfList.ElementAt(0).GetAttribute("name").ToString().Contains(convert.GetMonthName(dateTo.Month)))
                                    {
                                        Capabilities.driver.FindElementByXPath(listViewDateMonthPath + "//android.view.View[@index='0']").Click();
                                    }
                                }

                                if (Capabilities.driver.FindElementsByXPath(listViewDateMonthPath + "//android.view.View[@index='1']/*").Count > 1)
                                {
                                     IReadOnlyCollection<AndroidElement> nameOfList = Capabilities.driver.FindElementsByXPath(listViewDateMonthPath + "//android.view.View[@index='1']/*");
                                     if (nameOfList.ElementAt(0).GetAttribute("name").ToString().Contains(convert.GetMonthName(dateTo.Month)))
                                     {
                                         Capabilities.driver.FindElementByXPath(listViewDateMonthPath + "//android.view.View[@index='1']").Click();
                                     }
                                }
                                dateExist = DateTime.Parse(Capabilities.driver.FindElementByXPath(day).Text + " " + Capabilities.driver.FindElementByXPath(month).Text + " " + Capabilities.driver.FindElementByXPath(year).Text);
                                _xDelay("", 1000);
                               
                                if (convert.GetMonthName(dateExist.Month).Substring(0, 3).ToUpper().Equals(value[1].ToUpper()))
                                {
                                    if (Capabilities.driver.FindElementsByXPath(listViewDateMonthPath + "//android.view.View[@index='0']/*").Count > 1)
                                        Capabilities.driver.FindElementByXPath(listViewDateMonthPath + "//android.view.View[@index='0']//android.view.View[@index='" + (int.Parse(value[2]) - 1).ToString() + "']").Click();
                                    else
                                        Capabilities.driver.FindElementByXPath(listViewDateMonthPath + "//android.view.View[@index='1']//android.view.View[@index='" + (int.Parse(value[2]) - 1).ToString() + "']").Click();
                                    _xDelay("", 1000);
                                    flagCheckDateMonth = false;
                                    break;
                                }

                                if (dateExist.Month > dateTo.Month && flagCheckDateMonth==true)
                                {
                                    //Swipe dari atas ke bawah
                                    Capabilities.driver.Swipe(targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4), targetDateMonth.Location.Y + 5, targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4), (targetDateMonth.Location.Y + targetDateMonth.Size.Height*4/5) - 5, int.Parse(Capabilities.speedSwipe));
                                }
                                else if(dateExist.Month < dateTo.Month && flagCheckDateMonth==true)
                                { 
                                    //Swipe dari bawah ke atas
                                    Capabilities.driver.Swipe(targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4), (targetDateMonth.Location.Y + targetDateMonth.Size.Height*4/5) - 5, targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4), targetDateMonth.Location.Y + 5, int.Parse(Capabilities.speedSwipe));
                                }
                                _xDelay("", 1000);
                            }
                            if (!flagCheckDateMonth)
                                break;
                        }
                        //Batas SetDateTime untuk OS 5(Lollipop) 
                        Capabilities.driver.FindElementByXPath(buttonOkPath).Click();
                        break;

                        case "BIGCALENDAR2":
                        //SetDateTime untuk OS 6(Marshmallow)
                        string yearx = "", daymonthtemp = "", dayx = "", dayx2 = "", monthx = "", monthx2 = "", listViewDateMonthPathx = "", listviewYearPathx = "", listDay = ""
                        ,btnNextDay="",btnPreviousDay="";
                        int yearIndexx = 0;
                        for (int i = 0; i < dr.Length; i++)
                        {
                            //inisialisasi value awal dari year,month,day,listViewDateMonth,listviewYear
                            for (int j = 0; j < dr.Length; j++)
                            {
                                if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("YEAR"))
                                {
                                    yearx = dr[j]["AndroidResourceId"].ToString();
                                    yearIndexx = j;
                                }
                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("DAYMONTH"))
                                    daymonthtemp = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("LISTVIEWDATEMONTH"))
                                    listViewDateMonthPathx = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("LISTDAY"))
                                    listDay = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("LISTVIEWYEAR"))
                                    listviewYearPathx = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONOK"))
                                    buttonOkPath = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONCANCEL"))
                                    buttonCancelPath = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONNEXTDAY"))
                                    btnNextDay = dr[j]["AndroidResourceId"].ToString();

                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONPREVIOUSDAY"))
                                    btnPreviousDay = dr[j]["AndroidResourceId"].ToString();
                            }

                            DateTime dateExist, dateTo;
                            string[] daymonth = Capabilities.driver.FindElementByXPath(daymonthtemp).Text.Split(new string[] { " " }, StringSplitOptions.None);
                            dayx = daymonth[2];
                            monthx = daymonth[1];
                            dateExist = DateTime.Parse(dayx + " " + monthx + " " + Capabilities.driver.FindElementByXPath(yearx).Text);
                            dateTo = DateTime.Parse(value[2] + " " + value[1] + " " + value[3]);

                            //setYear
                            Capabilities.driver.FindElementByXPath(dr[yearIndexx]["AndroidResourceId"].ToString()).Click();
                            IReadOnlyCollection<AndroidElement> nameOfListYear = Capabilities.driver.FindElementsByXPath(listviewYearPathx);
                            AndroidElement targetYear = Capabilities.driver.FindElementByXPath(listViewDateMonthPathx);
                            bool flagCheckYear = true;
                            while (flagCheckYear)
                            {
                                for (int z = 0; z < nameOfListYear.Count; z++)
                                {

                                    if (nameOfListYear.ElementAt(z).Text.Equals(value[3]))
                                    {
                                        //klik tahun pada element z
                                        Capabilities.driver.FindElementByXPath(listviewYearPathx + "[@index='" + z + "']").Click();
                                        flagCheckYear = false;
                                        break;
                                    }
                                }

                                if (dateExist.Year > dateTo.Year && flagCheckYear == true)
                                {
                                    //Swipe dari atas ke bawah
                                    Capabilities.driver.Swipe(targetYear.Location.X + (targetYear.Size.Width / 4), targetYear.Location.Y + 5, targetYear.Location.X + (targetYear.Size.Width / 4), (targetYear.Location.Y + targetYear.Size.Height) - 5, int.Parse(Capabilities.speedSwipe));
                                    _xDelay("", 1000);
                                }
                                else if (dateExist.Year < dateTo.Year && flagCheckYear == true)
                                {
                                    //Swipe dari atas ke bawah
                                    Capabilities.driver.Swipe(targetYear.Location.X + (targetYear.Size.Width / 4), (targetYear.Location.Y + targetYear.Size.Height) - 5, targetYear.Location.X + (targetYear.Size.Width / 4), targetYear.Location.Y + 5, int.Parse(Capabilities.speedSwipe));
                                    _xDelay("", 1000);
                                }
                            }

                            //SetDateMonth
                            _xDelay("", 1000);
                            AndroidElement targetDateMonth = Capabilities.driver.FindElementByXPath(listDay);
                            bool flagCheckDateMonth = true;
                            
                            while (flagCheckDateMonth)
                            {
                                Capabilities.driver.FindElementByXPath(listDay + "//com.android.internal.widget.ViewPager[@resource-id='android:id/day_picker_view_pager']").Click();
                                //Capabilities.driver.FindElementByXPath(listDay + "//android.view.View[@resource-id='android:id/month_view']//android.view.View[@index='26']").Click();
                                _xDelay("", 1000);
                                string[] daymonth2 = Capabilities.driver.FindElementByXPath(daymonthtemp).Text.Split(new string[] { " " }, StringSplitOptions.None);
                                dayx2 = daymonth2[2];
                                monthx2 = daymonth2[1];
                                dateExist = DateTime.Parse(dayx2 + " " + monthx2 + " " + Capabilities.driver.FindElementByXPath(yearx).Text);
                                
                                if (convert.GetMonthName(dateExist.Month).Substring(0, 3).ToUpper().Equals(value[1].ToUpper()))
                                {
                                    Capabilities.driver.FindElementByXPath(listDay + "//android.view.View[@resource-id='android:id/month_view']//android.view.View[@index='" + (int.Parse(value[2]) - 1).ToString() + "']").Click();
                                    _xDelay("", 1000);
                                    flagCheckDateMonth = false;
                                    break;
                                }

                                if (dateExist.Month > dateTo.Month && flagCheckDateMonth == true)
                                {
                                    //Swipe dari kiri ke kanan
                                    //Capabilities.driver.Swipe(targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4), targetDateMonth.Location.Y + 5, (targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4)) - 5, targetDateMonth.Location.Y + 5, int.Parse(Capabilities.speedSwipe));
                                    Capabilities.driver.FindElementByXPath(btnPreviousDay).Click();
                                }
                                else if (dateExist.Month < dateTo.Month && flagCheckDateMonth == true)
                                {
                                    //Swipe dari kanan ke kiri
                                    //Capabilities.driver.Swipe((targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4)) - 5, targetDateMonth.Location.Y + 5, targetDateMonth.Location.X + (targetDateMonth.Size.Width / 4), targetDateMonth.Location.Y + 5, int.Parse(Capabilities.speedSwipe));
                                    Capabilities.driver.FindElementByXPath(btnNextDay).Click();
                                }
                                _xDelay("", 1000);
                            }
                            if (!flagCheckDateMonth)
                                break;
                        }
                        //Batas SetDateTime untuk OS 6(Marshmallow) 
                        Capabilities.driver.FindElementByXPath(buttonOkPath).Click();
                        break;
                    }
                }
                else if (temp[0].Contains("HH") || temp[0].Contains("mm"))
                {
                    for (int j = 0; j < tempDateTime.Length; j++)
                    {
                        //jam
                        if (tempDateTime[j].Equals("HH"))
                        {
                            setDateTime = setDateTime.AddHours(int.Parse(tempValue[j]));
                        }
                        //menit
                        else if (tempDateTime[j].Equals("mm"))
                        {
                            setDateTime = setDateTime.AddMinutes(int.Parse(tempValue[j]));
                        }
                    }
                    value[0] = setDateTime.Hour.ToString();
                    value[1] = setDateTime.Minute.ToString();

                    DataRow[] dr = FormTest.dsMasterResourceId.Tables[0].Select("OperatingSystem = '" + Capabilities.osVersion + "'" + " AND Modul like '%time%'");
                    switch(dr[0]["Modul"].ToString().ToUpper())
                    {
                       case "SMALLTIME":
                       string tempX, tempCheck, tempXPath="";
                       bool flagCheckSmallTime = true;
                        for (int i = 0; i < setTime.Length; i++)
                        {
                            for (int j = 0; j < dr.Length; j++)
                            {
                                if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("TIME"))
                                    tempXPath = dr[j]["AndroidResourceId"].ToString();
                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONOK"))
                                    buttonOkPath = dr[j]["AndroidResourceId"].ToString();
                                else if (dr[j]["UniqueIdentifier"].ToString().ToUpper().Equals("BUTTONCANCEL"))
                                    buttonCancelPath = dr[j]["AndroidResourceId"].ToString();
                            }
                            tempXPath = tempXPath.Replace("#Key", setTime[i]);      
                            _xWaitUntil(action, tempXPath, waitSeconds, type, desc, 1);

                            for (int h = 0; h < 5; h++)
                            {
                                tempX = int.Parse(value[i]).ToString();
                                tempCheck = int.Parse(Capabilities.driver.FindElementByXPath(tempXPath).Text).ToString();
                                if (!tempX.Equals(tempCheck))
                                {
                                    _xDelay(lastAction, 1000);
                                    Capabilities.driver.Tap(1, Capabilities.driver.FindElementByXPath(tempXPath), 1000);
                                    _xDelay(lastAction, 500);
                                    Capabilities.driver.FindElementByXPath(tempXPath).ReplaceValue(value[i]);
                                    _xDelay(lastAction, 500);
                                    Capabilities.driver.PressKeyCode(66);
                                    _xHideKeyboard();
                                }
                                else if (i == 1 && tempX.Equals(tempCheck))
                                {
                                    flagCheckSmallTime = false;
                                    break;
                                }
                            }
                            if(!flagCheckSmallTime)
                                Capabilities.driver.FindElementByXPath(buttonOkPath).Click();
                        }
                    break;
                    }
                }
                
                lastAction = action;
                lastElement = "";
                DateTime endTimedate = DateTime.Now;
                duration = DurationCalculate(startTime, endTimedate);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", (int.Parse(duration)/10).ToString(), desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xWriteWithoutHide(string action, string xPath, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc,1);
                endTime = DateTime.Now;
                lastElement = xPath;
                lastAction = action;
                Capabilities.driver.FindElementByXPath(xPath).Clear();
                Capabilities.driver.FindElementByXPath(xPath).SendKeys(textSent);
                duration = DurationCalculate(startTime,endTime);
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void SwitchPort(string pathFile, string key, string action, string desc, int delaySwitchPort)
        {
            try
            {
                if (key.ToUpper() == "DRIVER 1")
                {
                    if (Capabilities.tempDriver1 != null)
                    {
                        Capabilities.driver = Capabilities.tempDriver1;
                        _xDelay("", delaySwitchPort);
                    }

                    if (tid != null)
                    { tid.Abort(); }
                }
                else if (key.ToUpper() == "DRIVER 2")
                {
                    if (tempDriver2 != null)
                    {
                        Capabilities.driver = tempDriver2;
                        _xDelay("", delaySwitchPort);
                    }
                    else
                    {
                        //Eksekusi .bat file untuk memanggil driver appium
                        System.Diagnostics.Process.Start(pathFile);
                        _xDelay("", delaySwitchPort);
                        //Switch Port
                        Capabilities.driver = new AndroidDriver<AndroidElement>(new Uri(serverUri2), Capabilities._capabilities, TimeSpan.FromSeconds(180));
                        tempDriver2 = Capabilities.driver;
                        _xDelay("", delaySwitchPort);
                    }
                }
                else if (key.ToUpper() == "DRIVER 3")
                {
                    if (tempDriver3 != null)
                    {
                        Capabilities.driver = tempDriver3;
                        _xDelay("", delaySwitchPort);
                    }
                    else
                    {
                        //Eksekusi .bat file untuk memanggil driver appium
                        System.Diagnostics.Process.Start(pathFile);
                        
                        _xDelay("", delaySwitchPort);
                        //Switch Port
                        Capabilities.driver = new AndroidDriver<AndroidElement>(new Uri(serverUri3), Capabilities._capabilities, TimeSpan.FromSeconds(180));
                        tempDriver3 = Capabilities.driver;
                        _xDelay("", delaySwitchPort);
                    }
                }
                lastAction = action;
                lastElement = key;
                config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", "Not Use", desc);
            }
            catch (Exception ex)
            {
                // Log the exception
            }
        }

        public void WaitForThread(string xPath)
        {
            WebDriverWait wDriverWait = new WebDriverWait(Capabilities.driver, TimeSpan.FromMilliseconds(1000000));
            wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(xPath)));
        }

        public void ThreadWait(string xPath, string action, string desc)
        {
            lastAction = action;
            lastElement = xPath;
            tid = new Thread(() => WaitForThread(xPath));
            tid.Start();
            config._xInsertRowData(dateTimeCreatedXls, lastFeatureAction, lastAction, lastElement, "Success", "", "Not Use", desc);
        }

        #region addlog
        //public void _xAddLog(string strDesc)
        //{
        //    try
        //    {
        //        bool Dir = Directory.Exists(LogDirectory);
        //        if (Dir == false)
        //        {
        //            Directory.CreateDirectory(LogDirectory);
        //        }
        //        StreamWriter sw2 = new StreamWriter(LogDirectory + filelog, true, Encoding.ASCII);
        //        sw2.WriteLine(strDesc);
        //        sw2.Close();

        //        Trace.WriteLine(strDesc);
        //    }
        //    catch (Exception ex)
        //    {
        //        Trace.WriteLine(ex.Message);
        //    }
        //}
        #endregion
    }
}
