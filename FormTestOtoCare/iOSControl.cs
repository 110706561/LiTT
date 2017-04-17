using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.iOS;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Support.UI;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Drawing.Imaging;

namespace FormTestOtoCare
{
    class iOSControl
    {
        private Configuration config;
        private DateTime startTime, endTime;
        private static string serverUri1, serverUri2, serverUri3;
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";
        private string duration;
        private static string destinationJPG;
        public Thread tid;
        public IOSDriver<IOSElement> tempDriver=null,tempDriver2=null,tempDriver3=null;

        public iOSControl()
        {
            config = new Configuration();
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            serverUri1 = ini.GetValue("serverUri1", "IOSCapabilities");
            serverUri2 = ini.GetValue("serverUri2", "IOSCapabilities");
        }

        public void _xWaitUntil(string action, string nameId, int second, string type, string desc, int flag)
        {
            try
            {
                startTime = DateTime.Now;
                WebDriverWait wDriverWait = new WebDriverWait(iOSCapabilities.driver, TimeSpan.FromMilliseconds(second));
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "waiting for : " + nameId;
                switch (type.ToUpper())
                {
                    case "PATH": wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(nameId)));
                        break;
                    case "ELEMENT": wDriverWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id(nameId)));
                        break;
                }

                if (action.ToUpper().Equals("WAIT") && flag == 0)
                {
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCatchControl(string message, string desc)
        {
            endTime = DateTime.Now;
            duration = duration = Control.DurationCalculate(startTime, endTime);
            if (message.Contains("not found"))
            {
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Not Found", message.ToString(), "", "");
            }
            else
            {
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Error", message.ToString(), "", "");
            }

            Control.successFlag = 0;
        }

        public static void _xHideKeyboard(string XPath)
        {
            try
            {
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).SendKeys(Keys.Return);
            }
            catch (Exception ex)
            { }
        }

        public void _xClick(string action, string XPath, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, XPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = XPath;
                Control.lastAction = action;
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).Click();
                duration = duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xTap(string action, string XPath, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, XPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = XPath;
                Control.lastAction = action;
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).Tap(1,1);
                duration = duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xDBPostByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                duration = Control.DurationCalculate(startTime, endTime);
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).Clear();
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).SendKeys(Control.myDictionary[value]);
                Control._xHideKeyboard();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWrite(string action, string XPath, string textSent, string type, int waitSeconds, string desc)
        {
            try 
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, XPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = XPath;
                Control.lastAction = action;
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).Click();
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).Clear();
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).SendKeys(textSent);

                _xHideKeyboard(XPath);
                duration = duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);

            }
            catch(Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckSorted(string action, string xPath, string value, string type, int waitSeconds, string desc, string key)
        {
            try
            {
                string[] listOfIndex = key.Split('-');
                startTime = DateTime.Now;
                string tempPath = xPath.Replace("[SortKey]", "[" + listOfIndex[0] + "]");
                _xWaitUntil(action, tempPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                if (value.ToUpper().Equals("ASCENDING"))
                {
                    if (_xIsSortedAscending(xPath, listOfIndex))
                    {
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted ascending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else if (value.ToUpper().Equals("DESCENDING"))
                {
                    if (_xIsSortedDescending(xPath, listOfIndex))
                    {
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted descending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted descending", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        private bool _xIsSortedAscending(string xPath, string[] listOfIndex)
        {
            List<string> ListOfText = new List<string>();
            int startIndex = int.Parse(listOfIndex[0]);
            int endIndex = int.Parse(listOfIndex[1]);
            string newPath;
            for (int a = startIndex; a < endIndex; a++)
            {
                newPath = xPath.Replace("[SortKey]", "[" + a.ToString() + "]");
                ListOfText.Add(iOSCapabilities.driver.FindElement(By.XPath(newPath)).Text.ToString());
                if (a > startIndex)
                {
                    if (ListOfText.ElementAt(a - startIndex - 1) != ListOfText.ElementAt(a - startIndex))
                    {
                        #region Only for string
                        /*if (string.Compare(ListOfText.ElementAt(a - startIndex - 1), ListOfText.ElementAt(a - startIndex)) > 0)
                        {
                            return false;
                        }*/
                        #endregion
                        //Check for string and numeric using LINQ method
                        var OrderedByDesc = ListOfText.OrderByDescending(d => d);
                        if (ListOfText.SequenceEqual(OrderedByDesc))
                            return false;
                    }
                }
            }
            return true;
        }

        private bool _xIsSortedDescending(string xPath, string[] listOfIndex)
        {
            List<string> ListOfText = new List<string>();
            int startIndex = int.Parse(listOfIndex[0]);
            int endIndex = int.Parse(listOfIndex[1]);
            string newPath;
            for (int a = startIndex; a < endIndex; a++)
            {
                newPath = xPath.Replace("[SortKey]", "[" + a.ToString() + "]");
                ListOfText.Add(iOSCapabilities.driver.FindElement(By.XPath(newPath)).Text.ToString());
                if (a > startIndex)
                {
                    if (ListOfText.ElementAt(a - startIndex - 1) != ListOfText.ElementAt(a - startIndex))
                    {
                        #region Only for string
                        /*if (string.Compare(ListOfText.ElementAt(a - startIndex - 1), ListOfText.ElementAt(a - startIndex)) < 0)
                        {
                            return false;
                        }*/
                        #endregion
                        //Check for string and numeric using LINQ method
                        var OrderedByAsc = ListOfText.OrderBy(d => d);
                        if (ListOfText.SequenceEqual(OrderedByAsc))
                            return false;
                    }
                }
            }
            return true;
        }

        public void _xCheckSortedDate(string action, string xPath, string value, string type, int waitSeconds, string desc, string key)
        {
            try
            {
                string[] listOfIndex = key.Split('-');
                startTime = DateTime.Now;
                string tempPath = xPath.Replace("[SortKey]", "[" + listOfIndex[0] + "]");
                _xWaitUntil(action, tempPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                if (value.ToUpper().Equals("ASCENDING"))
                {
                    if (_xIsSortedDateAscending(xPath, listOfIndex))
                    {
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted ascending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else if (value.ToUpper().Equals("DESCENDING"))
                {
                    if (_xIsSortedDateDescending(xPath, listOfIndex))
                    {
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted descending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        duration = Control.DurationCalculate(startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted descending", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        private bool _xIsSortedDateAscending(string xPath, string[] listOfIndex)
        {
            IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
            List<string> ListOfText = new List<string>();
            List<string> ListValue = new List<string>();
            int startIndex = int.Parse(listOfIndex[0]);
            int endIndex = int.Parse(listOfIndex[1]);
            string newPath;
            for (int a = startIndex; a < endIndex; a++)
            {
                newPath = xPath.Replace("[SortKey]", "[" + a.ToString() + "]");
                ListOfText.Add(iOSCapabilities.driver.FindElement(By.XPath(newPath)).Text.ToString());
                if (a > startIndex)
                {
                    if (ListOfText.ElementAt(a - startIndex - 1) != ListOfText.ElementAt(a - startIndex))
                    {
                        #region Only for string
                        /*if (string.Compare(ListOfText.ElementAt(a - startIndex - 1), ListOfText.ElementAt(a - startIndex)) > 0)
                        {
                            return false;
                        }*/
                        #endregion
                        //Check for date with DateTimeCompare method
                        DateTime dateBefore = DateTime.Parse(ListOfText.ElementAt(a - startIndex - 1).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                        DateTime dateAfter = DateTime.Parse(ListOfText.ElementAt(a - startIndex).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);

                        //Descending greater than 0
                        //Ascending less than 0
                        //Same is 0
                        if (DateTime.Compare(dateBefore,dateAfter)>0)//Descending check because of return FALSE
                            return false;
                    }
                }
            }
            return true;
        }

        private bool _xIsSortedDateDescending(string xPath, string[] listOfIndex)
        {
            IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
            List<string> ListOfText = new List<string>();
            List<string> ListValue = new List<string>();
            int startIndex = int.Parse(listOfIndex[0]);
            int endIndex = int.Parse(listOfIndex[1]);
            string newPath;
            for (int a = startIndex; a < endIndex; a++)
            {
                newPath = xPath.Replace("[SortKey]", "[" + a.ToString() + "]");
                ListOfText.Add(iOSCapabilities.driver.FindElement(By.XPath(newPath)).Text.ToString());
                if (a > startIndex)
                {
                    if (ListOfText.ElementAt(a - startIndex - 1) != ListOfText.ElementAt(a - startIndex))
                    {
                        #region Only for string
                        /*if (string.Compare(ListOfText.ElementAt(a - startIndex - 1), ListOfText.ElementAt(a - startIndex)) > 0)
                        {
                            return false;
                        }*/
                        #endregion
                        //Check for date with DateTimeCompare method
                        DateTime dateBefore = DateTime.Parse(ListOfText.ElementAt(a - startIndex - 1).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                        DateTime dateAfter = DateTime.Parse(ListOfText.ElementAt(a - startIndex).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);

                        //Descending greater than 0
                        //Ascending less than 0
                        //Same is 0
                        if (DateTime.Compare(dateBefore, dateAfter) < 0)//Ascending check because of return FALSE
                            return false;
                    }
                }
            }
            return true;
        }

        #region delta action
        public void _xDismissByXPath(string action, string type, int waitSeconds, string desc)
        {
            try
            {
                Control.lastElement = "";
                Control.lastAction = action;
                tempDriver = iOSCapabilities.driver;
                //iOSCapabilities.driver.BackgroundApp(5);
                iOSCapabilities.driver.CloseApp();
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipe(string action, string timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                string[] timeExec = timeExecution.Split(',');
                if (timeExec.Length == 5)
                {
                    iOSCapabilities.driver.Swipe(Convert.ToInt32(timeExec[0]), Convert.ToInt32(timeExec[1]), Convert.ToInt32(timeExec[2]), Convert.ToInt32(timeExec[3]), Convert.ToInt32(timeExec[4]));
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Wrong format value, please input with format value 'start x, start y, end x, end y, interval time'", duration, desc);
                }
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
                Control.lastElement = xPath;
                Control.lastAction = action;
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).SendKeys(" ");
                Control._xHideKeyboard();
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xWriteWithoutHide(string action, string xPath, string textSent, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = xPath;
                Control.lastAction = action;
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).SendKeys(textSent);
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xDeleteText(string action, string XPath, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, XPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = XPath;
                Control.lastAction = action;
                iOSCapabilities.driver.FindElement(By.XPath(XPath)).Clear();
                _xHideKeyboard(XPath);
                duration = duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
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
                Control.lastElement = "Check result xpath : " + xPath;
                Control.lastAction = action;
                string tempValue="";
                string[] tempConcat = xPath.Split(new string[] { "/" }, StringSplitOptions.None);

                if (tempConcat[tempConcat.Length-1].ToUpper().Contains("UIASWITCH"))
                {
                    if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() == "1")
                        tempValue = "True";
                    else if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() == "0")
                        tempValue = "False";
                }
                else if (tempConcat[tempConcat.Length - 1].ToUpper().Contains("UIATEXTFIELD")||
                    tempConcat[tempConcat.Length - 1].ToUpper().Contains("UIABUTTON")||
                    tempConcat[tempConcat.Length - 1].ToUpper().Contains("UIATABLECELL"))
                {
                    if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).Enabled)
                        tempValue = "True";
                    else 
                        tempValue = "False";
                }

                if (tempValue.ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.actualValue = tempValue;
                    Control.successFlag = 0;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckExistByXpath(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("visible").ToUpper() == "TRUE")
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " not found", duration, desc);
                }
            }
            catch (Exception ex)
            {
                Control.successFlag = 0;
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " not found", duration, desc);
            }
        }

        public void _xCheckNotExistByXPath(string action, string xPath, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
               
                if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("visible").ToUpper() == "TRUE")
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " found", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
              
            }
            catch (Exception ex)
            {
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
        }

        public void _xCheckResultByXPath(string action, string xPath, string resultExpectation, string type, int waitSeconds, string desc)
        {
            //startTime =DateTime.Now;
            _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            Control.lastElement = "Check result xpath : " + xPath;
            Control.lastAction = action;

            string teks = "";
            if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() != null &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() != "" &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString().Length > 0)
            {
                teks = iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString();
            }
            else if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString() != null &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString() != "" &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString().Length > 0)
            {
                teks = iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString();
            }

            string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation, @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void _xWriteWithCheckResultByXPath(string action, string xPath, string value, string resultExpectation, string type, int waitSeconds, string desc)
        {
            //startTime =DateTime.Now;
            _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            Control.lastElement = "Check result xpath : " + xPath;
            Control.lastAction = action;

            iOSCapabilities.driver.FindElement(By.XPath(xPath)).SendKeys(value);
            string teks = "";
            string[] splitTeks;
            if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() != null &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() != "" &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString().Length > 0)
            {
                splitTeks = iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString().Split(new string[] { "," }, StringSplitOptions.None);
                teks = splitTeks[0];
            }
            else if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString() != null &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString() != "" &&
                iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString().Length > 0)
            {
                splitTeks = iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString().Split(new string[] { "," }, StringSplitOptions.None);
                teks = splitTeks[0];
            }

            string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation, @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void _xCheckResultEmptyByXPath(string action, string xPath, string resultExpectation, string type, int waitSeconds, string desc)
        {
            //startTime =DateTime.Now;
            _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
            endTime = DateTime.Now;
            Control.lastElement = "Check result xpath : " + xPath;
            Control.lastAction = action;
            
            if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).Text.ToString().Equals(resultExpectation.ToString()))
            {
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = iOSCapabilities.driver.FindElement(By.XPath(xPath)).Text.ToString();
                Control.successFlag = 0;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void _xSetDateTime(string action, string xPath, string Value, string type, int waitSeconds, string desc)
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
                string[] setDate = new string[] { "3", "1", "2"};
                string[] setTime = new string[] { "1", "2" };
                string[] value = new string[3];
                string tempXPath="";

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
                    value[0] = setDateTime.Year.ToString();
                    value[1] = convert.GetMonthName(setDateTime.Month);
                    value[2] = setDateTime.Day.ToString();

                    //SetDate
                    
                    for (int i = 0; i < setDate.Length; i++)
                    {
                        tempXPath = xPath.Replace("#Key", setDate[i]);
                        iOSCapabilities.driver.FindElement(By.XPath(tempXPath)).SendKeys(value[i]);
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
                    value[0] = setDateTime.Hour.ToString("00");
                    value[1] = setDateTime.Minute.ToString("00");
                    for (int i = 0; i < setTime.Length; i++)
                    {
                        tempXPath = xPath.Replace("#Key", setTime[i]);
                        iOSCapabilities.driver.FindElement(By.XPath(tempXPath)).SendKeys(value[i]);
                    }
                }
                Control.lastAction = action;
                Control.lastElement = xPath;
                DateTime endTimedate = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTimedate);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", (int.Parse(duration) / 10).ToString(), desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xCheckResultDateTime(string action, string xPath, string resultExpectation, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
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

                string[] value = new string[3];
                string tempResultExpectation = "", tempActualDate = "", tempActualTime="";
                string actualDay="", actualMonth="", actualYear="", actualHour="", actualMinute="";

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
                    value[0] = setDateTime.Year.ToString();
                    value[1] = convert.GetMonthName(setDateTime.Month);
                    value[2] = setDateTime.Day.ToString();
                    tempResultExpectation = value[1] + " " + value[2] + ", " + value[0];

                    actualMonth = iOSCapabilities.driver.FindElement(By.XPath(xPath.Replace("#Key","1"))).Text;
                    actualDay = iOSCapabilities.driver.FindElement(By.XPath(xPath.Replace("#Key","2"))).Text;
                    actualYear = iOSCapabilities.driver.FindElement(By.XPath(xPath.Replace("#Key","3"))).Text;
                    tempActualDate = actualMonth + " " + actualDay + ", " + actualYear;
                    Control.lastAction = action;
                    Control.lastElement = xPath;
                    //CheckDate
                    if (resultExpectation != null && resultExpectation.Length > 0)
                    {
                        
                        if (tempActualDate.Equals(tempResultExpectation))
                        {
                            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", "Not Use", desc);
                        }
                        else
                        {
                            Control.actualValue = tempActualDate;
                            Control.successFlag = 0;
                            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + tempResultExpectation, "Not Use", desc);
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
                    value[0] = setDateTime.Hour.ToString("00");
                    value[1] = setDateTime.Minute.ToString("00");
                    tempResultExpectation = value[0] + ":" + value[1];

                    
                    actualHour = iOSCapabilities.driver.FindElement(By.XPath(xPath.Replace("#Key", "1"))).Text;
                    actualHour = Regex.Replace(actualHour, "[^0-9]+", string.Empty);
                    actualMinute = iOSCapabilities.driver.FindElement(By.XPath(xPath.Replace("#Key", "2"))).Text;
                    actualMinute = Regex.Replace(actualMinute, "[^0-9]+", string.Empty);
                    tempActualTime = actualHour + ":" + actualMinute;
                    Control.lastAction = action;
                    Control.lastElement = xPath;

                    //CheckTime
                    if (resultExpectation != null && resultExpectation.Length > 0)
                    {

                        if (tempActualTime.Equals(tempResultExpectation))
                        {
                            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", "Not Use", desc);
                        }
                        else
                        {
                            Control.actualValue = tempActualTime;
                            Control.successFlag = 0;
                            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + tempResultExpectation, "Not Use", desc);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xGetElementByXPathSubstring(string action, string xPath, string ResultVariable, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                string[] teks = value.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                Control.myDictionary.Add(ResultVariable, iOSCapabilities.driver.FindElement(By.XPath(xPath)).Text.ToString().Substring(int.Parse(teks[0]), int.Parse(teks[1])));
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xScreenShot(string action, string desc)
        {
            try
            {
                startTime =DateTime.Now;
                Control.lastElement = "Screenshot";
                Control.lastAction = action;
                destinationJPG = FormTest.pictureFile + Control.lastFeatureAction + "-" + Control.lastAction + "-" + DateTime.Now.ToString("[dd-MM-yyyy HH-mm-ss]") + ".jpg";
                iOSCapabilities.driver.GetScreenshot().SaveAsFile(destinationJPG, ImageFormat.Jpeg);
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeDownNotificationBar(string action, int timeExecution, string desc)
        {
            try
            {
                startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height - iOSCapabilities.driver.Manage().Window.Size.Height, iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                endTime = DateTime.Now;
                duration = (int.Parse(Control.DurationCalculate(startTime, endTime))/10+timeExecution).ToString();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeUpNotificationBar(string action, int timeExecution, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height, iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                endTime = DateTime.Now;
                duration = (int.Parse(Control.DurationCalculate(startTime, endTime)) / 10 + timeExecution).ToString();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
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
                startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height * 4 / 5, iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height / 5, timeExecution);
                endTime = DateTime.Now;
                duration = (int.Parse(Control.DurationCalculate(startTime, endTime)) / 10 + timeExecution).ToString();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeDown(string action, int timeExecution, string desc)
        {
            try
            {
                startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height / 5, iOSCapabilities.driver.Manage().Window.Size.Width / 2, iOSCapabilities.driver.Manage().Window.Size.Height * 4 / 5, timeExecution);
                endTime = DateTime.Now;
                duration = (int.Parse(Control.DurationCalculate(startTime, endTime)) / 10 + timeExecution).ToString();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeHorizontalBar(string action, string xPath, int value, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                IOSElement target = iOSCapabilities.driver.FindElement(By.XPath(xPath));
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(target.Location.X, target.Location.Y + (target.Size.Height / 2), target.Location.X + target.Size.Width / 100 * value, target.Location.Y + (target.Size.Height / 2), 1000);
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeSideMenu(string action, string xPath, int value, string desc)
        {
            try
            {
                startTime =DateTime.Now;
                IOSElement target = iOSCapabilities.driver.FindElement(By.XPath(xPath));
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(target.Location.X + target.Size.Width, target.Location.Y + (target.Size.Height / 2), target.Location.X, target.Location.Y + (target.Size.Height / 2), value);
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
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
                IOSElement target = iOSCapabilities.driver.FindElement(By.XPath(xPath));
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(target.Location.X, target.Location.Y + (target.Size.Height / 2), target.Location.X, (target.Location.Y + (target.Size.Height / 2)) + target.Size.Width / 100 * value, 1000);
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
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
                Control.lastElement = "Check result xpath : " + xPath;
                Control.lastAction = action;
                string tempValue = "";

                if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() == "1")
                    tempValue = "True";
                else
                    tempValue = "False";

                if (tempValue.ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.actualValue = tempValue;
                    Control.successFlag = 0;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xDelay(string action, int msecond)
        {
            try
            {
                Control.lastElement = "";
                Control.lastAction = action;
                Thread.Sleep(msecond);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }

        public void _xZoomByXPath(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Zoom(iOSCapabilities.driver.FindElement(By.XPath(xPath)));
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckPortrait(string action, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
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

                if (iOSCapabilities.driver.Orientation.Equals(so))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Screen orientation not match, the actual orientation is: " + so.ToString(), duration, desc);
                }
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
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width * 3 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, iOSCapabilities.driver.Manage().Window.Size.Width * 7 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                //iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width * 7 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, iOSCapabilities.driver.Manage().Window.Size.Width * 3 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeLeft(string action, int timeExecution, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = "";
                iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width * 7 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, iOSCapabilities.driver.Manage().Window.Size.Width * 3 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                //iOSCapabilities.driver.Swipe(iOSCapabilities.driver.Manage().Window.Size.Width * 3 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, iOSCapabilities.driver.Manage().Window.Size.Width * 7 / 10, iOSCapabilities.driver.Manage().Window.Size.Height / 2, timeExecution);
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xResetAppByXPath(string action, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime =DateTime.Now;
                Control.lastElement = "";
                Control.lastAction = action;
                iOSCapabilities.driver.ResetApp();
                endTime = DateTime.Now;
                duration = (int.Parse(Control.DurationCalculate(startTime, endTime))).ToString();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        #endregion

        public void _xNoActionSupported(string action, string element)
        {
            try
            {
                Control.lastAction = action;
                Control.lastElement = element;
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Not Supported", "No action triggered", "", "");
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, "");
            }
        }
        //20161208
        public void _xGet(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                string tempValue = "";
                if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() != "" ||
                    iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString().Length > 0 ||
                    iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString() != null)
                {
                    tempValue = iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("value").ToString();
                }
                else if (iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString() != "" ||
                    iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString().Length > 0 ||
                    iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString() != null)
                {
                    tempValue = iOSCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("name").ToString();
                }
                Control.myDictionary.Add(value, tempValue);
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xGetSize(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                    string [] tempConcat = iOSCapabilities.driver.FindElement(By.XPath(xPath)).Text.ToString().Split(new string[] { "of" }, StringSplitOptions.None);
                Control.myDictionary.Add(value, tempConcat[tempConcat.Length-1]);
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckNotExistBySize(string action, string xPath, string value, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                string[] tempConcat = value.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                string a = Control.myDictionary[tempConcat[0]].ToString();
                string b = Control.myDictionary[tempConcat[1]].ToString();
                if ((int.Parse(a)-int.Parse(b))==0)
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element found", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
            }
            catch (Exception ex)
            {
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
        }

        public void _xPost(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                _xWrite(action, xPath, Control.myDictionary[value].ToString(), type, waitSeconds, desc);
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xCatchControl(ex.Message, desc);
            }
        }

        public void _xSwipeListToDelete(string action, string xPath, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                IOSElement target = iOSCapabilities.driver.FindElement(By.XPath(xPath));

                //int y = target.Location.Y + target.Size.Height / 2;
                //int xStart = (int)(target.Location.X + target.Size.Width * 0.8);
                //int xEnd = (int)(target.Location.X + target.Size.Width * 0.2);
                //TouchAction tAction = new TouchAction(iOSCapabilities.driver);
                iOSCapabilities.driver.Swipe((int)(target.Location.X + target.Size.Width * 0.8), target.Location.Y + (target.Size.Height / 2), (int)(target.Location.X + target.Size.Width * 0.2), target.Location.Y + (target.Size.Height / 2), 1000);
                /*tAction.Press(xStart, y)
                    .Wait(1000)
                    .MoveTo(xEnd, y)
                    .Release();*/
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);  
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
                    iOSCapabilities.driver.Dispose();
                    iOSCapabilities.driver = new IOSDriver<IOSElement>(new Uri(serverUri1), iOSCapabilities._capabilities, TimeSpan.FromSeconds(1000));
                    Thread.Sleep(5000);
                }
                else if (key.ToUpper() == "DRIVER 2")
                {
                    //Switch Port
                    iOSCapabilities.driver.Dispose();
                    iOSCapabilities.driver = new IOSDriver<IOSElement>(new Uri(serverUri2), iOSCapabilities._capabilities, TimeSpan.FromSeconds(1000));
                    Thread.Sleep(5000);
                }
                Control.lastAction = action;
                Control.lastElement = key;
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", "Not Use", desc);
            }
            catch (Exception ex)
            {
                // Log the exception
            }
        }
    }
}
