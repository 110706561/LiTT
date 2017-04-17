using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using AutoItX3Lib;
using System.Text.RegularExpressions;
using System.Drawing;


namespace FormTestOtoCare
{
    class WebControl
    {
        private Configuration config;
        private DateTime endTime;
        private int errorFlag = 0;
        private string duration { get; set; }
        private string actualValue { get; set; }
        
        public WebControl()
        {
            config = new Configuration();
        }

        public static string DurationCalculate(DateTime startTime, DateTime endTime, int delayTime)
        {
            int milisecond, second, minute, duration;

            milisecond = (endTime - startTime).Milliseconds;
            second = (endTime - startTime).Seconds;
            minute = (endTime - startTime).Minutes;

            duration = (milisecond + (second * 1000) + (minute * 60000))-delayTime;

            return duration.ToString();
        }

        public void Wait(int timeout, string type, string element, string action, string desc, int flag)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(WebCapabilities.driver, TimeSpan.FromMilliseconds(timeout));           
                Control.lastAction = action;
                Control.lastElement = "waiting for : " + element;

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(element)));

                if (action.ToUpper().Equals("WAIT"))
                {
                    if (flag == 0)
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void WaitByCSS(int timeout, string type, string element, string action, string desc, int flag)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(WebCapabilities.driver, TimeSpan.FromMilliseconds(timeout));
                Control.lastAction = action;
                Control.lastElement = "waiting for : " + element;

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.CssSelector(element)));

                if (action.ToUpper().Equals("WAIT"))
                {
                    if (flag == 0)
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Delay(string action, int msecond)
        {
            try
            {
                Thread.Sleep(msecond);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Write(string element, string elementtype, string value, string action, int waitSeconds, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, elementtype, element, action, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
                WebCapabilities.driver.FindElement(By.XPath(element)).Clear();
                WebCapabilities.driver.FindElement(By.XPath(element)).SendKeys(value);
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

       
        public void WriteWithJSInject(string element, string elementtype, string value, string action, int waitSeconds, string desc, string key)
        {
            try
            {
                Control.startTime = DateTime.Now;
                Wait(waitSeconds, elementtype, element, action, desc, 1);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;

                IWebElement webElement = WebCapabilities.driver.FindElement(By.XPath(element));
                IJavaScriptExecutor js = (IJavaScriptExecutor)WebCapabilities.driver;
                js.ExecuteScript(key, webElement);
                for (int x = 0; x < value.Length; x++)
                {
                    WebCapabilities.driver.FindElement(By.XPath(element)).SendKeys(value.Substring(x,1));
                }
                
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void SelectFileUpload(string element, string elementtype, string value, string action, int waitSeconds, string desc)
        {
            try
            {
                Control.startTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
                Thread.Sleep(5000);
                SendKeys.SendWait(value);
                Thread.Sleep(5000);
                SendKeys.SendWait(@"{Enter}");
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "File Path: " + value , duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }


        public void SetDateTime(string element, string elementtype, string Value, string action, int waitSeconds, string desc)
        {
            try
            {
                DateTime startTimedate = DateTime.Now;
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
                Control.lastElement = element;
                Control.lastAction = action;

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
                if (temp[0].Contains("HH") || temp[0].Contains("mm"))
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

                WebCapabilities.driver.FindElement(By.XPath(element)).Clear();
                WebCapabilities.driver.FindElement(By.XPath(element)).SendKeys(setDateTime.ToString(temp[3]));
                //WebCapabilities.driver.FindElement(By.XPath(element)).SendKeys(Keys.Enter);
                SendKeys.SendWait(@"{Enter}");

                DateTime endTimedate = DateTime.Now;
                duration = Control.DurationCalculate(startTimedate, endTimedate);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void SetDateTimeWithoutEnter(string element, string elementtype, string Value, string action, int waitSeconds, string desc)
        {
            try
            {
                string OldVar, tempVar, NewVar, checkDictionary;
                DateTime startTimedate = DateTime.Now;
                System.Globalization.DateTimeFormatInfo convert = new System.Globalization.DateTimeFormatInfo();
                IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
                string[] temp = Value.ToString().Split(new string[] { "," }, StringSplitOptions.None);
                string[] tempDateTime = temp[0].ToString().Split(new string[] { "&" }, StringSplitOptions.None);
                string[] tempValue = temp[2].ToString().Split(new string[] { "&" }, StringSplitOptions.None);
                DateTime setDateTime;
                if (temp[1].ToUpper().Equals("NOW"))
                    setDateTime = DateTime.Now;
                else if(temp[1].Contains('#'))
                {
                    OldVar = temp[1];
                    tempVar = OldVar;
                    checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                    if (checkDictionary != "default")
                    {
                        NewVar = Control.myDictionary[OldVar].ToString();
                        tempVar = tempVar.Replace(OldVar, NewVar);
                    }
                    setDateTime = DateTime.Parse(tempVar, culture, System.Globalization.DateTimeStyles.AssumeLocal);
                }
                else
                    setDateTime = DateTime.Parse(temp[1], culture, System.Globalization.DateTimeStyles.AssumeLocal);
                Control.lastElement = element;
                Control.lastAction = action;

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
                if (temp[0].Contains("HH") || temp[0].Contains("mm"))
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

                WebCapabilities.driver.FindElement(By.XPath(element)).Clear();
                WebCapabilities.driver.FindElement(By.XPath(element)).SendKeys(setDateTime.ToString(temp[3]));
                //WebCapabilities.driver.FindElement(By.XPath(element)).SendKeys(Keys.Enter);

                DateTime endTimedate = DateTime.Now;
                duration = Control.DurationCalculate(startTimedate, endTimedate);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void _xCheckResultDateTimeByKeyXPath(string action, string xPath, string resultExpectation, string type, int waitSeconds, string desc, string key)
        {
            //startTime =DateTime.Now;
            Actions mouse = new Actions(WebCapabilities.driver);
            WebDriverWait wait = new WebDriverWait(WebCapabilities.driver, TimeSpan.FromMilliseconds(10000));
           
            Control.lastElement = "Check result XPath : " + xPath;
            Control.lastAction = action;
            string teks = "";

            DateTime startTimedate = DateTime.Now;
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
            if (temp[0].Contains("HH") || temp[0].Contains("mm"))
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

            teks = WebCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute(key.ToLower()).ToString();

            string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(setDateTime.ToString(temp[3]), @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void SetTime(string element, string elementtype, string Value, string action, int waitSeconds, string desc)
        {
            try
            {
                DateTime startTimedate = DateTime.Now;
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
                string[] value = new string[3];
                Control.lastElement = element;
                Control.lastAction = action;

               if (temp[0].Contains("HH") || temp[0].Contains("mm"))
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

                value[0] = setDateTime.ToString("HH");
                value[1] = setDateTime.ToString("mm");
               
                string tempHour, tempMinute, incrementHour, incrementMinute, decreaseHour, decreaseMinute;

                //Cek Jam
                tempHour = element + "/tr[2]/td[1]";
                string tempHourCheck;
                do
                {
                    tempHourCheck = WebCapabilities.driver.FindElement(By.XPath(tempHour)).Text;
                    if (int.Parse(tempHourCheck) < int.Parse(value[0]))
                    {
                        incrementHour = element + "/tr[1]/td[1]/a";
                        WebCapabilities.driver.FindElement(By.XPath(incrementHour)).Click();
                    }
                    else if (int.Parse(tempHourCheck) > int.Parse(value[0]))
                    {
                        decreaseHour = element + "/tr[3]/td[1]/a";
                        WebCapabilities.driver.FindElement(By.XPath(decreaseHour)).Click();
                    }
                } while (int.Parse(tempHourCheck) != int.Parse(value[0])) ;

                //Cek Menit
                tempMinute = element + "/tr[2]/td[3]";
                string tempMinuteCheck;
                do
                {
                    tempMinuteCheck = WebCapabilities.driver.FindElement(By.XPath(tempMinute)).Text;
                    if (int.Parse(tempMinuteCheck) < int.Parse(value[1]))
                    {
                        incrementMinute = element + "/tr[1]/td[3]/a";
                        WebCapabilities.driver.FindElement(By.XPath(incrementMinute)).Click();
                    }
                    else if (int.Parse(tempMinuteCheck) > int.Parse(value[1]))
                    {
                        decreaseMinute = element + "/tr[3]/td[3]/a";
                        WebCapabilities.driver.FindElement(By.XPath(decreaseMinute)).Click();
                    }
                } while (tempMinuteCheck != value[1]) ;
                        
                
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void DeleteText(string element, string elementtype)
        {
            try
            {
                WebCapabilities.driver.FindElement(By.XPath(element)).Clear();
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void ClickByPath(string element, string elementtype, string action, int waitSeconds, int delayTime ,string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, elementtype, element, action, desc, 1);
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
               
                WebCapabilities.driver.FindElement(By.XPath(element)).Click();
                  
                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void ClickByCss(string element, string elementtype, string action, int waitSeconds, int delayTime, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                WaitByCSS(waitSeconds, elementtype, element, action, desc, 1);
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;

                WebCapabilities.driver.FindElement(By.CssSelector(element)).Click();

                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void DoubleClick(string element, string elementtype, string action, int waitSeconds, int delayTime, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, elementtype, element, action, desc, 1);
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;

                WebCapabilities.driver.FindElement(By.XPath(element)).Click();
                WebCapabilities.driver.FindElement(By.XPath(element)).Click();

                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void LoopClick(string element, string elementtype, string action, int waitSeconds, int delayTime, string desc, string loop)
        {
            try
            {
                int i;
                //Control.startTime = DateTime.Now;
                for (i = 0; i < int.Parse(loop); i++)
                {
                    Wait(waitSeconds, elementtype, element, action, desc, 1);
                    Delay(action, delayTime);
                    Control.lastElement = element;
                    Control.lastAction = action;
                    WebCapabilities.driver.FindElement(By.XPath(element)).Click();
                }
                endTime = DateTime.Now;
                duration = DurationCalculate(Control.startTime, endTime, (delayTime*i));
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void ClickByKey(string element, string elementtype, string action, int waitSeconds, int delayTime ,string desc, string key)
        {
            try
            {
                Thread.Sleep(5000);
                IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
                List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
                //To calculate no of rows In table.
                int flag = 0;
                int rows_count = rows_table.Capacity;
                Control.lastAction = action;
                if(elementtype.ToUpper().Equals("PATH"))
                {
                    //Loop will execute till the last row of table.
                    for (int row = 0; row < rows_count; row++)
                    {
                        //To locate columns(cells) of that specific row.
                        List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();
                        //To calculate no of columns(cells) In that specific row.
                        int columns_count = Columns_row.Capacity;
                        //Loop will execute till the last cell of that specific row.
                        for (int column = 0; column < columns_count; column++)
                        {
                            //To retrieve text from that specific cell.
                            String celtext = Columns_row[column].Text;
                            if (celtext.Contains(key))
                            {
                                element = element + "/tr[" + (row + 1) + "]/td[" + (column + 1) + "]";
                                Control.lastElement = element;
                                WebCapabilities.driver.FindElement(By.XPath(element)).Click();
                                Delay(action, delayTime);
                                flag = 1;
                                break;
                            }
                        }
                        if (flag == 1)
                            break;
                    }
                }
                endTime = DateTime.Now;
                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                element = "";
            }
            catch (Exception ex)
            {
                //Cek error session buka report terakhir
                CatchControl(ex.Message);
            }
        }

        public void Select(string element, string elementtype, string action, string value, string index, int waitSeconds, int delayTime ,string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
                switch (elementtype.ToUpper())
                {
                    case "TEXT": new SelectElement(WebCapabilities.driver.FindElement(By.XPath(element))).SelectByText(value);
                        break;
                    case "INDEX": new SelectElement(WebCapabilities.driver.FindElement(By.XPath(element))).SelectByIndex(int.Parse(index));
                        break;
                    case "VALUE": new SelectElement(WebCapabilities.driver.FindElement(By.XPath(element))).SelectByValue(value);
                        break;
                }
                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Hover(string element, string elementtype, string action, int waitSeconds,int delayTime ,string desc)
        {
            try
            {
                Actions mouse = new Actions(WebCapabilities.driver);
                //Control.startTime = DateTime.Now;
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
                mouse.MoveToElement(WebCapabilities.driver.FindElement(By.XPath(element))).Build().Perform();
                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Submit(string element, string elementtype, string action, int waitSeconds, int delayTime ,string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
                WebCapabilities.driver.FindElement(By.XPath(element)).Submit();
                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void AlertAuthentication(string element, string key, string value)
        {
            try
            {
                var AutoIT = new AutoItX3();
                AutoIT.WinWaitActive(element);
                String[] keys = key.Split(';');
                String[] values = value.Split(';');
                int a = 0;
                do
                {
                    string tempKey = keys[a].Substring(1, keys[a].Length - 1);
                    AutoIT.ControlFocus(element, "", tempKey);
                    string tempValue = values[a];
                    AutoIT.Send(tempValue);
                    a++;
                } while (a < keys.Length);

            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void AcceptAlert(string action, int delayTime)
        {
            try
            {
                Delay(action, delayTime);
                WebCapabilities.driver.SwitchTo().Alert().Accept();
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void GoTo(string baseUrl, string action, string desc)
        {
            try
            {
                int checkframe;
                Control.lastElement = baseUrl;
                Control.lastAction = action;
                WebCapabilities.driver.Navigate().GoToUrl(baseUrl);
                checkframe = WebCapabilities.driver.FindElements(By.TagName("frame")).Count;
                if (checkframe > 0)
                    WebCapabilities.driver.SwitchTo().Frame(0);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", "Not Use", desc);
            }
            catch { }
        }

        public void Back()
        {
            try
            {
                WebCapabilities.driver.Navigate().Back();
                //WebCapabilities.driver.SwitchTo().Frame(0);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Forward()
        {
            try
            {
                WebCapabilities.driver.Navigate().Forward();
                //WebCapabilities.driver.SwitchTo().Frame(0);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Stop()
        {
            try
            {
                WebCapabilities.driver.FindElement(By.TagName("body")).SendKeys("Keys.ESCAPE");
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void Refresh()
        {
            try
            {
                WebCapabilities.driver.Navigate().Refresh();
                //WebCapabilities.driver.SwitchTo().Frame(0);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void ScrollTo(string element, string elementtype, string action, int waitSeconds, int delayTime, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, elementtype, element, action, desc, 1);
                Delay(action, delayTime);
                endTime = DateTime.Now;
                Control.lastElement = element;
                Control.lastAction = action;
                var scrollElementPath =  WebCapabilities.driver.FindElement(By.XPath(element));
                Actions actionsPath = new Actions(WebCapabilities.driver);
                actionsPath.MoveToElement(scrollElementPath);
                actionsPath.Perform();
                duration = DurationCalculate(Control.startTime, endTime, delayTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void _xCheckRowCountByXPath(string action, string element, string resultExpectation, string desc)
        {
            Control.startTime = DateTime.Now;
            Control.lastAction = action;
            Control.lastElement = element;

            Thread.Sleep(5000);
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();
           
            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;
            string lastResult = "", lastStatus = "";

            string fixedStringOne = Regex.Replace(rows_count.ToString().ToLower(), @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation.ToLower(), @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = rows_count.ToString();
                Control.successFlag = 0;
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
            

            endTime = DateTime.Now;
            duration = Control.DurationCalculate(Control.startTime, endTime);
            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, lastStatus, lastResult, duration, desc);
        }

        public void _xCheckResultColorByCSS(string action, string css, string resultExpectation, string type, int waitSeconds, string desc, string key)
        {
            //startTime =DateTime.Now;
            WaitByCSS(waitSeconds, type, css, action, desc, 1);
            endTime = DateTime.Now;
            Control.lastElement = "Check result css : " + css;
            Control.lastAction = action;
            string replaceRGB = WebCapabilities.driver.FindElement(By.CssSelector(css)).GetCssValue(key.ToLower());
            int startIndex = replaceRGB.IndexOf('(');
            replaceRGB = replaceRGB.Substring(startIndex + 1, replaceRGB.Length - startIndex - 2);

            string[] splitRGB = replaceRGB.Split(',');
            Color color = Color.FromArgb(int.Parse(splitRGB[0]), int.Parse(splitRGB[1]), int.Parse(splitRGB[2]));

            string teks = "";
            teks = "#"+color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            string fixedStringOne = Regex.Replace(teks.ToLower(), @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation.ToLower(), @"\s+", String.Empty);

            bool isEqual = fixedStringOne.Contains(fixedStringTwo);
            if (isEqual)
            {
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void _xCheckResultByXPath(string action, string xPath, string resultExpectation, string type, int waitSeconds, string desc)
        {
            //startTime =DateTime.Now;
            Wait(waitSeconds, type, xPath, action, desc, 1);
            endTime = DateTime.Now;
            Control.lastElement = "Check result xpath : " + xPath; 
            Control.lastAction = action;

            string teks = "";
            teks = WebCapabilities.driver.FindElement(By.XPath(xPath)).Text.ToString();
            
            string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation, @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void _xCheckResultByKeyCSS(string action, string css, string resultExpectation, string type, int waitSeconds, string desc,string key)
        {
            //startTime =DateTime.Now;
            WaitByCSS(waitSeconds, type, css, action, desc, 1);
            endTime = DateTime.Now;
            Control.lastElement = "Check result css : " + css;
            Control.lastAction = action;

            string teks = "";
            teks = WebCapabilities.driver.FindElement(By.CssSelector(css)).GetAttribute(key.ToLower()).ToString();
            string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation, @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void _xCheckResultByKeyXPath(string action, string xPath, string resultExpectation, string type, int waitSeconds, string desc, string key)
        {
            //startTime =DateTime.Now;
            Actions mouse = new Actions(WebCapabilities.driver);
            WebDriverWait wait = new WebDriverWait(WebCapabilities.driver, TimeSpan.FromMilliseconds(10000));
            endTime = DateTime.Now;
            Control.lastElement = "Check result XPath : " + xPath;
            Control.lastAction = action;
            string teks = "";

            teks = WebCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute(key.ToLower()).ToString();

            string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
            string fixedStringTwo = Regex.Replace(resultExpectation, @"\s+", String.Empty);

            bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
            if (isEqual)
            {
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            else
            {
                Control.actualValue = teks;
                Control.successFlag = 0;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, duration, desc);
            }
        }

        public void CheckExistByCss(string action, string elementId, string desc)
        {
            try
            {
                Control.lastAction = action;
                Control.lastElement = elementId;
                //Control.startTime = DateTime.Now;
                if (WebCapabilities.driver.FindElement(By.CssSelector(elementId)).Displayed)
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    Control.successFlag = 0;
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + elementId + " not found", duration, desc);
                }
            }
            catch (Exception ex)
            {
                Control.successFlag = 0;
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + elementId + " not found", duration, desc);
            }
        }

        public void CheckExistByXpath(string action, string xPath, string desc)
        {
            try
            {
                Control.lastAction = action;
                Control.lastElement = xPath;
                //Control.startTime = DateTime.Now;
                if (WebCapabilities.driver.FindElement(By.XPath(xPath)).Displayed)
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.XPath(xPath)).Displayed.ToString();
                    Control.successFlag = 0;
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " not found", duration, desc);
                }
            }
            catch (Exception ex)
            {
                Control.successFlag = 0;
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " not found", duration, desc);
            }
        }

        public void CheckNotExistByXpath(string action, string xPath, string desc)
        {
            try
            {
                Control.lastAction = action;
                Control.lastElement = xPath;
                //Control.startTime = DateTime.Now;
                if (WebCapabilities.driver.FindElement(By.XPath(xPath)).Displayed)
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " found", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.XPath(xPath)).Displayed.ToString();

                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
            }
            catch (Exception ex)
            {
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
        }

        public void CheckMenuNotExistByXpath(string action, string xPath, string resultExpectation, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                endTime = DateTime.Now;
                Control.lastElement = "Check result xpath : " + xPath;
                Control.lastAction = action;

                string teks = "";
                teks = WebCapabilities.driver.FindElement(By.XPath(xPath)).Text.ToString();

                string fixedStringOne = Regex.Replace(teks, @"\s+", String.Empty);
                string fixedStringTwo = Regex.Replace(resultExpectation, @"\s+", String.Empty);

                bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
                if (isEqual)
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Element " + xPath + " found", duration, desc);
                }
                else
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.XPath(xPath)).Displayed.ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
            }
            catch (Exception ex)
            {
                endTime = DateTime.Now;
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
        }

        public void CheckEnabledByElementId(string action, string elementId, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, type, elementId, action, desc, 1);
                Control.lastElement = "Check result element : " + elementId;
                Control.lastAction = action;
                
                if (WebCapabilities.driver.FindElement(By.Id(elementId)).Enabled.ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.Id(elementId)).Enabled.ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void CheckEnabledByXPath(string action, string xPath, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, type, xPath, action, desc, 1);
                Control.lastElement = "Check result xpath : " + xPath;
                Control.lastAction = action;
               
                if (WebCapabilities.driver.FindElement(By.XPath(xPath)).Enabled.ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.XPath(xPath)).Enabled.ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void CheckReadOnlyByElementId(string action, string elementId, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, type, elementId, action, desc, 1);
                Control.lastElement = "Check result element : " + elementId;
                Control.lastAction = action;

                if (WebCapabilities.driver.FindElement(By.Id(elementId)).GetAttribute("readonly").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.Id(elementId)).GetAttribute("readonly").ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void CheckReadOnlyByXPath(string action, string xPath, string type, int waitSeconds, string resultExpectation, string desc)
        {
            try
            {
                //Control.startTime = DateTime.Now;
                Wait(waitSeconds, type, xPath, action, desc, 1);
                Control.lastElement = "Check result xpath : " + xPath;
                Control.lastAction = action;

                if (WebCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("readonly").ToString().ToUpper().Equals(resultExpectation.ToString().ToUpper()))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute("readonly").ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void CheckSelectedtById(string action, string elementId, string resultExpectation, string desc)
        {
            //Control.startTime = DateTime.Now;
            Control.lastElement = "Check result element : " + elementId;
            Control.lastAction = action;
            try
            {
                if (WebCapabilities.driver.FindElement(By.Id(elementId)).Selected.ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.Id(elementId)).Selected.ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void CheckSelectedtByXPath(string action, string xPath, string resultExpectation, string desc)
        {
            //Control.startTime = DateTime.Now;
            Control.lastElement = "Check result element : " + xPath;
            Control.lastAction = action;
            
            try
            {
                if (WebCapabilities.driver.FindElement(By.XPath(xPath)).Selected.ToString().ToUpper().Contains(resultExpectation.ToString().ToUpper()))
                {
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
                else
                {
                    Control.successFlag = 0;
                    endTime = DateTime.Now;
                    duration = Control.DurationCalculate(Control.startTime, endTime);
                    actualValue = WebCapabilities.driver.FindElement(By.XPath(xPath)).Selected.ToString();
                    config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + actualValue + " and expectation : " + resultExpectation, duration, desc);
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }

        }


        public void _xCheckSortedDateTime(string action, string xPath, string value, string type, int waitSeconds, string desc, string columnKey)
        {
            try
            {
                Control.startTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                if (value.ToUpper().Equals("ASCENDING"))
                {
                    if (_xIsSortedAscendingDateTime(xPath, columnKey))
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted ascending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else if (value.ToUpper().Equals("DESCENDING"))
                {
                    if (_xIsSortedDescendingDateTime(xPath, columnKey))
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted descending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted descending", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public bool _xIsSortedAscendingDateTime(string element, string columnKey)
        {
            //*[@id="pageset1-0"]/div[1]/a2is-datatable/div[2]/div/table/tbody/tr[1]/td[2]
            IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();

            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;

            //Loop will execute till the last row of table.
            int startIndex = 0;
            int endIndex = rows_count - 1;
            for (int row = 0; row < rows_count; row++)
            {
                //To locate columns(cells) of that specific row.
                List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();

                //To retrieve text from that specific cell.
                ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
                if (row > startIndex)
                {
                    if (ListOfText.ElementAt(row - startIndex - 1) != ListOfText.ElementAt(row - startIndex))
                    {
                        //Check for date with DateTimeCompare method
                        DateTime dateBefore = DateTime.Parse(ListOfText.ElementAt(row - startIndex - 1).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                        DateTime dateAfter = DateTime.Parse(ListOfText.ElementAt(row - startIndex).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);

                        //Descending greater than 0
                        //Ascending less than 0
                        //Same is 0
                        if (DateTime.Compare(dateBefore, dateAfter) > 0)//Descending check because of return FALSE
                            return false;
                    }
                }
            }
            return true;
        }

        public bool _xIsSortedDescendingDateTime(string element, string columnKey)
        {
            //*[@id="pageset1-0"]/div[1]/a2is-datatable/div[2]/div/table/tbody/tr[1]/td[2]
            IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();

            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;

            //Loop will execute till the last row of table.
            int startIndex = 0;
            int endIndex = rows_count - 1;
            for (int row = 0; row < rows_count; row++)
            {
                //To locate columns(cells) of that specific row.
                List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();

                //To retrieve text from that specific cell.
                ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
                if (row > startIndex)
                {
                    if (ListOfText.ElementAt(row - startIndex - 1) != ListOfText.ElementAt(row - startIndex))
                    {
                        //Check for date with DateTimeCompare method
                        DateTime dateBefore = DateTime.Parse(ListOfText.ElementAt(row - startIndex - 1).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                        DateTime dateAfter = DateTime.Parse(ListOfText.ElementAt(row - startIndex).ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);

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
        ///////////////////////////////

        public void _xCheckSortedDropDown(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                Control.startTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                if (value.ToUpper().Equals("ASCENDING"))
                {
                    if (_xIsSortedAscendingDropDown(xPath))
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted ascending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else if (value.ToUpper().Equals("DESCENDING"))
                {
                    if (_xIsSortedDescendingDropDown(xPath))
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted descending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted descending", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public bool _xIsSortedAscendingDropDown(string element)
        {
            Thread.Sleep(5000);
            List<IWebElement> buttonList = WebCapabilities.driver.FindElements(By.XPath(element+"//*")).ToList();
            List<string> ListOfText = new List<string>();

            //Loop will execute till the last row of table.
            int startIndex = 0;
            int endIndex = buttonList.Count - 1;
            for (int x = 0; x < buttonList.Count; x++)
            {
                string[] splitButtonText = buttonList[x].GetAttribute("outerHTML").Split('>');
                string[] split2ButtonText = splitButtonText[1].Split('<');
                ListOfText.Add(split2ButtonText[0]);

                if (x > startIndex)
                {
                    if (ListOfText.ElementAt(x - startIndex - 1) != ListOfText.ElementAt(x - startIndex))
                    {
                        //Check for string and numeric using LINQ method
                        var OrderedByDesc = ListOfText.OrderByDescending(d => d);
                        if (ListOfText.SequenceEqual(OrderedByDesc))
                            return false;
                    }
                }
            }
            return true;
        }

        public bool _xIsSortedDescendingDropDown(string element)
        {
            Thread.Sleep(5000);
            List<IWebElement> buttonList = WebCapabilities.driver.FindElements(By.XPath(element+"//*")).ToList();
            List<string> ListOfText = new List<string>();

            //Loop will execute till the last row of table.
            int startIndex = 0;
            int endIndex = buttonList.Count - 1;
            for (int x = 0; x < buttonList.Count; x++)
            {
                string[] splitButtonText = buttonList[x].GetAttribute("outerHTML").Split('>');
                string[] split2ButtonText = splitButtonText[1].Split('<');
                ListOfText.Add(split2ButtonText[0]);

                if (x > startIndex)
                {
                    if (ListOfText.ElementAt(x - startIndex - 1) != ListOfText.ElementAt(x - startIndex))
                    {
                        //Check for string and numeric using LINQ method
                        var OrderedByAsc = ListOfText.OrderBy(d => d);
                        if (ListOfText.SequenceEqual(OrderedByAsc))
                            return false;
                    }
                }
            }
            return true;
        }
        
        ///////////////////////////////
   
        public void _xCheckSorted(string action, string xPath, string value, string type, int waitSeconds, string desc, string columnKey)
        {
            try
            {
                Control.startTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                if (value.ToUpper().Equals("ASCENDING"))
                {
                    if (_xIsSortedAscending(xPath, columnKey))
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted ascending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted ascending", duration, desc);
                    }
                }
                else if (value.ToUpper().Equals("DESCENDING"))
                {
                    if (_xIsSortedDescending(xPath, columnKey))
                    {
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "List sorted descending", duration, desc);
                    }
                    else
                    {
                        Control.successFlag = 0;
                        endTime = DateTime.Now;
                        duration = Control.DurationCalculate(Control.startTime, endTime);
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "List not sorted descending", duration, desc);
                    }
                }
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public bool _xIsSortedAscending(string element, string columnKey)
        {
                //*[@id="pageset1-0"]/div[1]/a2is-datatable/div[2]/div/table/tbody/tr[1]/td[2]
                IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
                List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
                List<string> ListOfText = new List<string>();

                //To calculate no of rows In table.
                int rows_count = rows_table.Capacity;
                
                //Loop will execute till the last row of table.
                int startIndex = 0;
                int endIndex = rows_count-1;
                for (int row = 0; row < rows_count; row++)
                {
                    //To locate columns(cells) of that specific row.
                    List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();
                    
                    //To retrieve text from that specific cell.
                    ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
                    if (row > startIndex)
                    {
                        if (ListOfText.ElementAt(row - startIndex - 1) != ListOfText.ElementAt(row - startIndex))
                        {
                            //Check for string and numeric using LINQ method
                            var OrderedByDesc = ListOfText.OrderByDescending(d => d);
                            if (ListOfText.SequenceEqual(OrderedByDesc))
                                return false;
                        }
                    }
                }
                return true;
        }

        public bool _xIsSortedDescending(string element, string columnKey)
        {
            //*[@id="pageset1-0"]/div[1]/a2is-datatable/div[2]/div/table/tbody/tr[1]/td[2]
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();

            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;

            //Loop will execute till the last row of table.
            int startIndex = 0;
            int endIndex = rows_count - 1;
            for (int row = 0; row < rows_count; row++)
            {
                //To locate columns(cells) of that specific row.
                List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();
            
                //To retrieve text from that specific cell.
                ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
                if (row > startIndex)
                {
                    if (ListOfText.ElementAt(row - startIndex - 1) != ListOfText.ElementAt(row - startIndex))
                    {
                        //Check for string and numeric using LINQ method
                        var OrderedByAsc = ListOfText.OrderBy(d => d);
                        if (ListOfText.SequenceEqual(OrderedByAsc))
                            return false;
                    }
                }
            }
            return true;
        }

        ///////////////////////////////

        public void _xGetElementByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Wait(waitSeconds, type, xPath, action, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                Control.myDictionary.Add(value, WebCapabilities.driver.FindElements(By.XPath(xPath)).ToList()[0].Text);
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void _xGetElementByKeyXPath(string action, string xPath, string value, string type, int waitSeconds, string key, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Wait(waitSeconds, type, xPath, action, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                Control.myDictionary.Add(value, WebCapabilities.driver.FindElement(By.XPath(xPath)).GetAttribute(key).ToString());
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void _xPostByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                //startTime =DateTime.Now;
                Wait(waitSeconds, type, xPath, action, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                WebCapabilities.driver.FindElement(By.XPath(xPath)).SendKeys(Control.myDictionary[value].ToString());
                duration = Control.DurationCalculate(Control.startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }

        public void _xCheckColumnResult(string action, string element, string resultExpectation, string columnKey, string desc)
        {
            Control.startTime = DateTime.Now;
            Control.lastAction = action;
            Control.lastElement = element;

            Thread.Sleep(5000);
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();
            string[] splitResultExpectation = resultExpectation.ToString().Split(new string[] { ";" }, StringSplitOptions.None);
            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;
            string lastResult="", lastStatus="";
            int flag = 1;
            
            //Loop will execute till the last row of table.
            for (int row = 0; row < rows_count; row++)
            {
                //To locate columns(cells) of that specific row.
                List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();
                //To retrieve text from that specific cell.
                ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
            }
            
            for (int x = 0; x < splitResultExpectation.Length; x++)
            {
                string fixedStringOne = Regex.Replace(ListOfText[x].ToLower(), @"\s+", String.Empty);
                string fixedStringTwo = Regex.Replace(splitResultExpectation[x].ToLower(), @"\s+", String.Empty);

                bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
                if (isEqual)
                {
                    lastResult = lastResult + "Match;";
                }
                else
                {
                    lastResult = lastResult + "Unmatch;";
                    Control.successFlag = 0;
                    flag = Control.successFlag;
                }
            }

            if (flag == 1)
                lastStatus = "Success";
            else
                lastStatus = "Failed";

            endTime = DateTime.Now;
            duration = Control.DurationCalculate(Control.startTime, endTime);
            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, lastStatus, lastResult, duration, desc);
        }

        public void _xCheckResultDropDown(string action, string element, string resultExpectation, string columnKey, string desc)
        {
            Control.startTime = DateTime.Now;
            Control.lastAction = action;
            Control.lastElement = element;

            Thread.Sleep(5000);
            List<IWebElement> buttonList = WebCapabilities.driver.FindElements(By.XPath(element + "//*")).ToList();
            List<string> ListOfText = new List<string>();
            string[] splitResultExpectation = resultExpectation.ToString().Split(new string[] { ";" }, StringSplitOptions.None);
            //To calculate no of rows In table.
            string lastResult = "", lastStatus = "";
            //Loop will execute till the last row of dropdown.
            for (int x = 0; x < buttonList.Count; x++)
            {
                string[] splitButtonText = buttonList[x].GetAttribute("outerHTML").Split('>');
                string[] split2ButtonText = splitButtonText[1].Split('<');
                ListOfText.Add(split2ButtonText[0]);
            }

            for (int x = 0; x < splitResultExpectation.Length; x++)
            {
                string fixedStringOne = Regex.Replace(ListOfText[x].ToLower(), @"\s+", String.Empty);
                string fixedStringTwo = Regex.Replace(splitResultExpectation[x].ToLower(), @"\s+", String.Empty);

                bool isEqual = String.Equals(fixedStringOne, fixedStringTwo, StringComparison.OrdinalIgnoreCase);
                if (isEqual)
                {
                    lastResult = lastResult + "Match;";
                }
                else
                {
                    lastResult = lastResult + "Unmatch;";
                    Control.successFlag = 0;
                }
            }

            if (Control.successFlag == 1)
                lastStatus = "Success";
            else
                lastStatus = "Failed";

            endTime = DateTime.Now;
            duration = Control.DurationCalculate(Control.startTime, endTime);
            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, lastStatus, lastResult, duration, desc);
        }

        public void _xCheckColumnResultContain(string action, string element, string resultExpectation, string columnKey, string desc)
        {
            Control.startTime = DateTime.Now;
            Control.lastAction = action;
            Control.lastElement = element;

            Thread.Sleep(5000);
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();
            string[] splitResultExpectation = resultExpectation.ToString().Split(new string[] { ";" }, StringSplitOptions.None);
            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;
            string lastResult = "", lastStatus = "";
            //Loop will execute till the last row of table.
            for (int row = 0; row < rows_count; row++)
            {
                //To locate columns(cells) of that specific row.
                List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();
                //To retrieve text from that specific cell.
                ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
            }

            for (int x = 0; x < splitResultExpectation.Length; x++)
            {
                string fixedStringOne = Regex.Replace(ListOfText[x].ToLower(), @"\s+", String.Empty);
                string fixedStringTwo = Regex.Replace(splitResultExpectation[x].ToLower(), @"\s+", String.Empty);

                if (fixedStringOne.Contains(fixedStringTwo))
                {
                    lastResult = lastResult + "Match;";
                }
                else
                {
                    lastResult = lastResult + "Unmatch;";
                    Control.successFlag = 0;
                }
            }

            if (Control.successFlag == 1)
                lastStatus = "Success";
            else
                lastStatus = "Failed";

            endTime = DateTime.Now;
            duration = Control.DurationCalculate(Control.startTime, endTime);
            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, lastStatus, lastResult, duration, desc);
        }

        public void _xCheckColumnResultContainOnlyOneExpectation(string action, string element, string resultExpectation, string columnKey, string desc)
        {
            Control.startTime = DateTime.Now;
            Control.lastAction = action;
            Control.lastElement = element;
            
            Thread.Sleep(5000);
            IWebElement mytable = WebCapabilities.driver.FindElement(By.XPath(element));
            List<IWebElement> rows_table = mytable.FindElements(By.TagName("tr")).ToList();
            List<string> ListOfText = new List<string>();
            //To calculate no of rows In table.
            int rows_count = rows_table.Capacity;
            string lastResult = "", lastStatus = "";
            //Loop will execute till the last row of table.
            for (int row = 0; row < rows_count; row++)
            {
                //To locate columns(cells) of that specific row.
                List<IWebElement> Columns_row = rows_table[row].FindElements(By.TagName("td")).ToList();
                //To retrieve text from that specific cell.
                ListOfText.Add(Columns_row[int.Parse(columnKey)].Text);
            }

            for (int x = 0; x < rows_count; x++)
            {
                string fixedStringOne = Regex.Replace(ListOfText[x].ToLower(), @"\s+", String.Empty);
                string fixedStringTwo = Regex.Replace(resultExpectation.ToLower(), @"\s+", String.Empty);

                if (fixedStringOne.Contains(fixedStringTwo))
                {
                    lastResult = lastResult + "Match;";
                }
                else
                {
                    lastResult = lastResult + "Unmatch;";
                    Control.successFlag = 0;
                }
            }

            if (Control.successFlag == 1)
                lastStatus = "Success";
            else
                lastStatus = "Failed";

            endTime = DateTime.Now;
            duration = Control.DurationCalculate(Control.startTime, endTime);
            config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, lastStatus, lastResult, duration, desc);
        }

        public void CatchControl(string message)
        {
            if (message.Contains("not found"))
            {
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Not Found", message.ToString(), "Not Use", "");
            }
            else
            {
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Error", message.ToString(), "Not Use", "");
            }

            string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Error message : " + Control.lastFeatureAction + " | " + Control.lastElement + "|" + message;
            errorFlag = 1;
            Control.successFlag = 0;
        }

        public void SuccessControl()
        {
            if (errorFlag != 1)
            {
                string errorMessage = DateTime.Now.ToString("[dd-MM-yyyy HH:mm:ss]") + " - Success message : Success - " + Control.lastFeatureAction;
            }
        }

        public void QuitAction()
        {
            WebCapabilities.driver.Quit();
        }

        public void CloseAction()
        {
            WebCapabilities.driver.Close();
        }

        public void NoActionSupported(string action, string element)
        {
            try
            {
                Control.lastAction = action;
                Control.lastElement = element;
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Not Supported", "No action triggered", "Not Use", "");
            }
            catch (Exception ex)
            {
                CatchControl(ex.Message);
            }
        }
    }
}
