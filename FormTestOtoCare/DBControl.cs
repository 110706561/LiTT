using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Data;
using System.Threading;
using System.Data.SqlClient;
using OpenQA.Selenium;

namespace FormTestOtoCare
{
    class DBControl
    {
        //public static Dictionary<string, string> myDictionaries = new Dictionary<string, string>();
        private Configuration config = new Configuration();
        private Control _xControl = new Control();
        private DateTime startTime, endTime;
        private string duration;

        //Add By Ferie
        public void _xDBExec(string action, string sqlconn, string query, string desc)
        {
            try
            {
                /////////////////////////
                startTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = query;
                string strConn = sqlconn;
                SqlConnection sqlConn = new SqlConnection(strConn);
                SqlCommand sqlCmd = new SqlCommand();
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = query;
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
                endTime = DateTime.Now;
                /////////////////////
                duration = Control.DurationCalculate(startTime, endTime);
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }

        public void _xDBGet(string action, string sqlconn, string query, string resultVarible, string desc)
        {
            try
            {
                /////////////////////////
                startTime = DateTime.Now;
                DataSet ds = new DataSet();
                string strConn = sqlconn;
                SqlConnection sqlConn = new SqlConnection(strConn);
                SqlCommand sqlCmd = new SqlCommand();
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = query;
                SqlDataAdapter da = new SqlDataAdapter(sqlCmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                ds.Tables.Add(dt);
                sqlConn.Close();
                endTime = DateTime.Now;
                /////////////////////
                Control.lastAction = action;
                Control.lastElement = query;
                duration = Control.DurationCalculate(startTime, endTime);
                String[] resultvariables = resultVarible.Split(';');
                for (int a = 0; a < resultvariables.Length; a++)
                {
                    string tempResultVariable = resultvariables[a].Substring(1, resultvariables[a].Length - 1);
                    Control.myDictionary.Add(resultvariables[a], ds.Tables[0].Rows[0][tempResultVariable].ToString());
                }
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);

            }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }

        public void _xDBGetColumn(string action, string sqlconn, string query, string resultVarible, string desc)
        {
            try
            {
                /////////////////////////
                string tempValue = "";
                startTime = DateTime.Now;
                DataSet ds = new DataSet();
                string strConn = sqlconn;
                SqlConnection sqlConn = new SqlConnection(strConn);
                SqlCommand sqlCmd = new SqlCommand();
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = query;
                SqlDataAdapter da = new SqlDataAdapter(sqlCmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                ds.Tables.Add(dt);
                sqlConn.Close();
                endTime = DateTime.Now;
                /////////////////////
                Control.lastAction = action;
                Control.lastElement = query;
                duration = Control.DurationCalculate(startTime, endTime);
               
                String[] resultvariables = resultVarible.Split(';');
                for (int a = 0; a < resultvariables.Length; a++)
                {
                    string tempResultVariable = resultvariables[a].Substring(1, resultvariables[a].Length - 1);
                    tempValue = "";
                    for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                    { 
                        if(x + 1 == ds.Tables[0].Rows.Count)
                            tempValue = tempValue + ds.Tables[0].Rows[x][tempResultVariable].ToString();
                        else
                            tempValue = tempValue + ds.Tables[0].Rows[x][tempResultVariable].ToString() + ";";
                    }
                    Control.myDictionary.Add(resultvariables[a], tempValue);
                }
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);

            }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }

        public void _xDBPostByElementId(string action, string elementId, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xControl._xWaitUntil(action, elementId, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = elementId;
                duration = Control.DurationCalculate(startTime, endTime);
                Capabilities.driver.FindElementById(elementId).ReplaceValue(Control.myDictionary[value]);
                Control._xHideKeyboard();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
                }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }

        public void _xDBPostByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xControl._xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                duration = Control.DurationCalculate(startTime, endTime);
                Capabilities.driver.FindElementByXPath(xPath).ReplaceValue(Control.myDictionary[value]);
                Control._xHideKeyboard();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }

        public void _xDBPostNumericByXPath(string action, string xPath, string value, string type, int waitSeconds, string desc)
        {
            try
            {
                startTime = DateTime.Now;
                _xControl._xWaitUntil(action, xPath, waitSeconds, type, desc, 1);
                endTime = DateTime.Now;
                Control.lastAction = action;
                Control.lastElement = xPath;
                duration = Control.DurationCalculate(startTime, endTime);
                string textSent = Control.myDictionary[value];

                int numKeys;
                for (int x = 0; x < textSent.Length; x++)
                {
                    numKeys = int.Parse(textSent.Substring(x, 1));
                    if (numKeys == 0) { Capabilities.driver.PressKeyCode(7); }
                    else if (numKeys == 1) { Capabilities.driver.PressKeyCode(8);}
                    else if (numKeys == 2) { Capabilities.driver.PressKeyCode(9); }
                    else if (numKeys == 3) { Capabilities.driver.PressKeyCode(10); }
                    else if (numKeys == 4) { Capabilities.driver.PressKeyCode(11); }
                    else if (numKeys == 5) { Capabilities.driver.PressKeyCode(12); }
                    else if (numKeys == 6) { Capabilities.driver.PressKeyCode(13); }
                    else if (numKeys == 7) { Capabilities.driver.PressKeyCode(14); }
                    else if (numKeys == 8) { Capabilities.driver.PressKeyCode(15); }
                    else if (numKeys == 9) { Capabilities.driver.PressKeyCode(16); }
                }

                Control._xHideKeyboard();
                config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", duration, desc);
            }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }

        public void _xCheckResultVariableByXPath(string action, string value, string resultExpectation, string type, string desc)
        {
            try
            {
                Control.lastAction = action;
                Control.lastElement = "";
                if (resultExpectation != null && resultExpectation.Length > 0)
                {
                    if (Control.myDictionary[value].ToString().Contains(resultExpectation.ToString()))
                    {
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", "Not Use", desc);
                    }
                    else
                    {
                        Control.actualValue = Control.myDictionary[value].ToString();
                        Control.successFlag = 0;
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, "Not Use", desc);
                    }
                }
                else
                {
                    if (Control.myDictionary[value].ToString().Equals(resultExpectation.ToString()))
                    {
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Success", "", "Not Use", desc);
                    }
                    else
                    {
                        Control.actualValue = Control.myDictionary[value].ToString();
                        Control.successFlag = 0;
                        config._xInsertRowData(Control.dateTimeCreatedXls, Control.lastFeatureAction, Control.lastAction, Control.lastElement, "Failed", "Not match between actual : " + Control.actualValue + " and expectation : " + resultExpectation, "Not Use", desc);
                    }
                }
            }
            catch (Exception ex)
            {
                _xControl._xCatchControl(ex.Message, desc);
            }
        }
    }
}
