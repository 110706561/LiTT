using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Remote;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Data.OleDb;
using System.IO;
using System.Configuration;


namespace FormTestOtoCare
{
    public partial class FormTest : Form
    {
        public static DataSet dsConfiguration;
        public static DataSet dsScenario;
        public static DataSet dsBaseAlternateClass;
        public static DataSet dsMasterResourceId;
        public static DataSet dsDictionary;
        public static DataSet dsData, dsdataTemp;

        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";
        private string dataSourceConfiguration;
        private string dataSourceReport;
        public static string dataSourceBaseAlternateClass;
        //private static string defaultBashPath;
        private string masterResourceId,dictionary;
        public string platform;

        private int maxDataCounter,maxDataCounterTemp;
        public static string pictureFile;
        public static string deviceType;
        private string [] Result;
        private string[,] LoopResult;
        //private int waitSeconds;
        private string tempSheetName;
        private int count = 0;
        private int totalSuccess = 0;
        private int totalFailed = 0;
        private int TotalDataScenario = 0;

        private int initiOSDriver = 0;
        private int initAndroidDriver = 0;
        private int initWebDriver = 0;

        //Add By Ferie
        private DBControl _xDBControl = new DBControl();
        private Configuration xConfiguration = new Configuration();
        //Capabilities capabilitesInvoke;
        DataTable dt = new DataTable();
        DataRow dr = null;
        private int LoopingCount;

        public FormTest()
        {
            InitializeComponent();
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            dataSourceConfiguration = ini.GetValue("dataSourceConfiguration", "file");
            dataSourceReport = ini.GetValue("dataSourceReport", "file");
            pictureFile = ini.GetValue("pictureFile", "file");
            dataSourceBaseAlternateClass = ini.GetValue("baseAlternateClass", "file");
            masterResourceId = ini.GetValue("masterResourceId", "file");
            dictionary = ini.GetValue("dictionary", "file");
        }

        private void FormTest_Load(object sender, EventArgs e)
        {
            Clear();
            InitGridView();
        }

        public void BeforeAll(int a)
        {
            platform = dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper();
            if (dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper().Equals("ANDROID") && initAndroidDriver == 0)
            {
                Capabilities capabilitiesInvoke = new Capabilities();
                capabilitiesInvoke.CapabilitiesDevice();
                initAndroidDriver = 1;
            }
            else if (dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper().Equals("IOS") && initiOSDriver == 0)
            {
                iOSCapabilities capabilitiesInvoke = new iOSCapabilities();
                capabilitiesInvoke.CapabilitiesDevice();
                initiOSDriver = 1;
            }
            else if (dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper().Equals("WEB") && initWebDriver == 0)
            {
                WebCapabilities capabilitiesInvoke = new WebCapabilities();
                capabilitiesInvoke.CapabilitiesDriver();
                initWebDriver = 1;
            }
            
        }

        public void AfterAll()
        {
            if (initiOSDriver == 1)
            {
                iOSCapabilities.driver.Quit();
                initiOSDriver = 0;
            }
            else if (initAndroidDriver == 1)
            {
                Capabilities.driver.Quit();
                initAndroidDriver = 0;
            }
            else if (initWebDriver == 1)
            {
                WebCapabilities.driver.Close();
                initWebDriver = 0;
            }
        }

        private void Clear()
        {
            count = 0;
            totalSuccess = 0;
            totalFailed = 0;
            TotalDataScenario = 0;
            lblStep.Width = 100;
            lblScenario.Width = 100;
            lblScenario.Text = "-";
            lblStep.Text = "-";
            lblTotalSuccess.Text = "0";
            lblTotalFailed.Text = "0";
            lblTotalScenario.Text = "0/0";
            lblTotalStep.Text = "0/0";
            lblPercentScenario.Text = "0%";
            lblPercentageCurrentStep.Text = "0%";
            progressBar1.Value = 0;
            progressBar2.Value = 0;
        }

        private void InitGridView()
        {
            dt.Columns.Add(new DataColumn("No", typeof(string)));
            dt.Columns.Add(new DataColumn("Scenario", typeof(string)));
            dt.Columns.Add(new DataColumn("Percentage(%)", typeof(string)));
            dt.Columns.Add(new DataColumn("Status", typeof(string)));
            dataGridView1.DataSource = dt;
            dataGridView1.Columns["No"].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns["Scenario"].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns["Percentage(%)"].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns["Status"].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns["No"].Width = 50;
            dataGridView1.Columns["Scenario"].Width = 172;
            dataGridView1.Columns["Percentage(%)"].Width = 100;
            dataGridView1.Columns["Status"].Width = 120;
            dataGridView1.ForeColor = System.Drawing.Color.Black;
        }

        private void InsertDetailScenario(string lastFeatureAction)
        {
            count = count + 1;
            dr = dt.NewRow();
            dr["No"] = count;
            dr["Scenario"] = lastFeatureAction;
            dr["Percentage(%)"] = "0%";
            dr["Status"] = "On Progress";
            dt.Rows.Add(dr);
            bindingSource1.DataSource = dt;
            dataGridView1.DataSource = bindingSource1;
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
        }

        private void UpdateDetailScenario(string value, int row, int column)
        {
            if (value == "1")
            {
                value = "Success";
                totalSuccess = totalSuccess + 1;
            }
            else if (value == "0")
            {
                value = "Failed";
                totalFailed = totalFailed + 1;
            }

            dt.Rows[row-1][column] = value;
            dataGridView1.DataSource = dt;
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            dsConfiguration = xConfiguration._xReadExcelFile(dataSourceConfiguration);
            dsMasterResourceId = xConfiguration._xReadExcelFile(masterResourceId);
            dsDictionary = xConfiguration._xReadExcelFile(dictionary);
            Result = new string[dsConfiguration.Tables[0].Rows.Count];
                       
            if (btnStart.Text == "START")
            {
                if (txtLoopingCount.Text == "" || txtLoopingCount.Text == "0")
                    MessageBox.Show("Scenario Looping tidak boleh kosong atau nol!!!");
                else
                {
                    LoopingCount = int.Parse(txtLoopingCount.Text);
                    Clear();
                    Control.successFlag = 1;
                    dt.Rows.Clear();
                    dataGridView1.DataSource = dt;
                    btnStart.Text = "STOP";
                    //Hitung total data scenario looping
                    for (int x = 0; x < dsConfiguration.Tables[0].Rows.Count; x++)
                    {
                        if (dsConfiguration.Tables[0].Rows[x]["Data"].ToString() != "")
                        {
                            dsdataTemp = xConfiguration._xReadExcelFile(dsConfiguration.Tables[0].Rows[x]["Data"].ToString());
                            maxDataCounterTemp = dsdataTemp.Tables[0].Rows.Count;
                            TotalDataScenario = TotalDataScenario + maxDataCounterTemp;
                        }
                        else if (dsConfiguration.Tables[0].Rows[x]["Data"].ToString() == "")
                        {
                            TotalDataScenario = TotalDataScenario + 1;
                        }
                    }
                    LoopResult = new string[dsConfiguration.Tables[0].Rows.Count, TotalDataScenario];
                    //BeforeAll();
                    if (txtLoopingCount.Text != "" && txtLoopingCount.Text != "0" && txtLoopingCount.Text != "1")
                    {
                        for (int i = 0; i < int.Parse(txtLoopingCount.Text); i++)
                        {
                            StartTestAction();
                        }
                    }
                    else
                    {
                        StartTestAction();
                    }
                    AfterAll();
                    MessageBox.Show("Test Complete!!!");
                    Array.Clear(Result, 0, Result.Length);

                    if(TotalDataScenario>0)
                    Array.Clear(LoopResult, 0, LoopResult.Length);
                    btnStart.Text = "START";
                }
            }
            else if (btnStart.Text == "STOP")
            {
                btnStart.Text = "START";
                Clear();
                AfterAll();
                MessageBox.Show("Test Fail!!!");
            }
        }

        public void StartTestAction()
        {
            Control.dateTimeCreatedXls = dataSourceReport + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xls";
            Control.actionPath = "";
            Control.tempNum = 0;
            int a = 0;
            xConfiguration._xCreateExcelFile(Control.dateTimeCreatedXls);
            dsConfiguration = xConfiguration._xReadExcelFile(dataSourceConfiguration);

            for (; a < dsConfiguration.Tables[0].Rows.Count; a++)
            {

                BeforeAll(a);
                Control xControl = new Control();
                Control.insertText = "";

                //Assign last feature (Login, PremiumSimulation, etc)
                Control.lastFeatureAction = xConfiguration._xReplaceSpecialChars(dsConfiguration.Tables[0].Rows[a]["Scenario"].ToString());
                maxDataCounter = 1;

                if (dsConfiguration.Tables[0].Rows[a]["Data"] != null && dsConfiguration.Tables[0].Rows[a]["Data"].ToString().Length > 1)
                {
                    dsData = xConfiguration._xReadExcelFile(dsConfiguration.Tables[0].Rows[a]["Data"].ToString());
                    maxDataCounter = dsData.Tables[0].Rows.Count;
                }

                for (int d = 0; d < maxDataCounter; d++)
                {
                    Control.successFlag = 1;

                    if (d == 0)
                    {
                        //Create sheet in report
                        xConfiguration._xCreateSheetExcelFile(Control.dateTimeCreatedXls, Control.lastFeatureAction);
                        tempSheetName = Control.lastFeatureAction;
                    }
                    else
                    {
                        //Create sheet in report
                        Control.lastFeatureAction = tempSheetName + "_" + (d + 1).ToString();
                        xConfiguration._xCreateSheetExcelFile(Control.dateTimeCreatedXls, Control.lastFeatureAction);
                    }

                    //Read pathfile scenario xls
                    dsScenario = xConfiguration._xReadExcelFile(dsConfiguration.Tables[0].Rows[a]["PathFile"].ToString());

                    //Add By Ferie
                    lblScenario.Text = Control.lastFeatureAction;
                    InsertDetailScenario(Control.lastFeatureAction);

                    if (txtLoopingCount.Text == "1" && TotalDataScenario == 1)
                        lblTotalScenario.Text = (a + 1).ToString() + " / " + dsConfiguration.Tables[0].Rows.Count.ToString();
                    else if (txtLoopingCount.Text != "1" && TotalDataScenario == 1)
                        lblTotalScenario.Text = count.ToString() + " / " + ((dsConfiguration.Tables[0].Rows.Count * LoopingCount).ToString());
                    else if (txtLoopingCount.Text == "1" && TotalDataScenario > 1 || txtLoopingCount.Text != "1" && TotalDataScenario > 1)
                        lblTotalScenario.Text = count.ToString() + " / " + ((TotalDataScenario * LoopingCount).ToString());
                    Application.DoEvents();

                    for (int b = 0; b < dsScenario.Tables[0].Rows.Count; b++)
                    {
                            //Add by Ferie
                            progressBar2.Value = b + 1;
                            progressBar2.Maximum = dsScenario.Tables[0].Rows.Count;
                            lblStep.Text = dsScenario.Tables[0].Rows[b]["Action"].ToString();
                            lblTotalStep.Text = (b + 1).ToString() + " / " + dsScenario.Tables[0].Rows.Count.ToString();
                            lblPercentageCurrentStep.Text = (Math.Round(((double)progressBar2.Value / (double)progressBar2.Maximum), 2) * 100).ToString() + "%";
                            UpdateDetailScenario(lblPercentageCurrentStep.Text, count, 2);
                            Application.DoEvents();

                            if (dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper().Equals("ANDROID"))
                            {
                                AndroidAction AndroidStart = new AndroidAction();
                                AndroidStart.StartTestAction(b, d);
                            }
                            else if (dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper().Equals("IOS"))
                            {
                                iOSAction iOSStart = new iOSAction();
                                iOSStart.StartTestAction(b, d);
                            }
                            else if (dsConfiguration.Tables[0].Rows[a]["Platform"].ToString().ToUpper().Equals("WEB"))
                            {
                                WebAction WebStart = new WebAction();
                                WebStart.StartTestAction(b, d);
                            }

                        #region b dalam
                        //if (!dsScenario.Tables[0].Columns.Contains("Skip") || (dsScenario.Tables[0].Columns.Contains("Skip") && !dsScenario.Tables[0].Rows[b]["Skip"].ToString().ToUpper().Equals("YES")))
                        //{

                        //    //Pengecekkan Looping(Compare nilai pada tiap iterasi)
                        //    if (dsScenario.Tables[0].Rows[b]["Value"] != null && dsScenario.Tables[0].Rows[b]["Value"].ToString().Length > 0 && dsScenario.Tables[0].Rows[b]["Value"].ToString()[0].Equals('$'))
                        //    {
                        //        dsScenario.Tables[0].Rows[b]["Value"] = dsData.Tables[0].Rows[d][dsScenario.Tables[0].Rows[b]["Value"].ToString().Substring(1, dsScenario.Tables[0].Rows[b]["Value"].ToString().Length - 1)];
                        //    }
                        //    if (dsScenario.Tables[0].Rows[b]["ResultExpectation"] != null && dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Length > 0 && dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString()[0].Equals('$'))
                        //    {
                        //        dsScenario.Tables[0].Rows[b]["ResultExpectation"] = dsData.Tables[0].Rows[d][dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Substring(1, dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Length - 1)];
                        //    }
                        //    if (dsScenario.Tables[0].Rows[b]["Key"] != null && dsScenario.Tables[0].Rows[b]["Key"].ToString().Length > 0 && dsScenario.Tables[0].Rows[b]["Key"].ToString()[0].Equals('$'))
                        //    {
                        //        dsScenario.Tables[0].Rows[b]["Key"] = dsData.Tables[0].Rows[d][dsScenario.Tables[0].Rows[b]["Key"].ToString().Substring(1, dsScenario.Tables[0].Rows[b]["Key"].ToString().Length - 1)];
                        //    }
                        //    if (dsScenario.Tables[0].Rows[b]["ResourceId"] != null && dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Length > 0 && dsScenario.Tables[0].Rows[b]["ResourceId"].ToString()[0].Equals('$'))
                        //    {
                        //        //dsScenario.Tables[0].Rows[b]["ResourceId"] = dsMasterResourceId.Tables[0].Select("Variable = " + dsScenario.Tables[0].Rows[b]["ResourceId"].ToString()).ToString();
                        //        DataRow dr = dsMasterResourceId.Tables[0].Select("Variable = '" + dsScenario.Tables[0].Rows[b]["ResourceId"].ToString() + "'")[0];
                        //        dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["ResourceId"].ToString();
                        //    }

                        //    string tempChecking = "";
                        //    if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH") && dsScenario.Tables[0].Rows[b]["Class"].ToString().ToUpper() != null && dsScenario.Tables[0].Rows[b]["Class"].ToString().Length > 1)
                        //    {
                        //        tempChecking = xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, true);
                        //    }
                        //    if (!dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("WAIT") && !dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST") && (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH")) && (dsScenario.Tables[0].Rows[b]["Class"].ToString().Length > 1) && !Control._xCheckBaseClass(tempChecking, waitSeconds, dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                        //    {
                        //        int c, flagAlternate = 0;
                        //        dsBaseAlternateClass = xConfiguration._xReadAlternateClass(dataSourceBaseAlternateClass, dsScenario.Tables[0].Rows[b]["Class"].ToString());
                        //        if (dsBaseAlternateClass.Tables[0].Rows.Count > 0)
                        //        {
                        //            string tempClass = dsScenario.Tables[0].Rows[b]["Class"].ToString();
                        //            for (c = 0; c < dsBaseAlternateClass.Tables[0].Rows.Count; c++)
                        //            {
                        //                dsScenario.Tables[0].Rows[b]["Class"] = dsBaseAlternateClass.Tables[0].Rows[c]["AlternateClass"].ToString();
                        //                tempChecking = xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, true);
                        //                if (Control._xCheckBaseClass(tempChecking, waitSeconds, dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                        //                {
                        //                    flagAlternate = 1;
                        //                    break;
                        //                }
                        //            }
                        //            if (c == dsBaseAlternateClass.Tables[0].Rows.Count && !dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST"))
                        //            {
                        //                dsScenario.Tables[0].Rows[b]["Class"] = tempClass;
                        //                xControl._xNoClassSupported(dsScenario.Tables[0].Rows[b]["Class"].ToString());
                        //            }
                        //            else if (flagAlternate != 1 && !dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST"))
                        //            {
                        //                xControl._xNoClassSupported(dsScenario.Tables[0].Rows[b]["Class"].ToString());
                        //            }
                        //        }
                        //        else
                        //        {
                        //            xControl._xNoClassOrElement(dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        //        }
                        //    }


                        //    string[] elementClass = dsScenario.Tables[0].Rows[b]["Class"].ToString().Split(Control.delimiterChar);
                        //    switch (dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper())
                        //    {
                        //        //definition of each action by its type
                        //        case "WRITE":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xWriteByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["Value"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xWriteByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Value"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            break;

                        //        case "CLICK":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xClickByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xClickByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            break;

                        //        case "SWIPE": xControl._xSwipe(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        //            break;
                        //        case "SWIPEUP": xControl._xSwipeUp(dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //            break;
                        //        case "SWIPEDOWN": xControl._xSwipeDown(dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //            break;
                        //        case "SWIPERIGHT": xControl._xSwipeRight(dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //            break;
                        //        case "SWIPELEFT": xControl._xSwipeLeft(dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //            break;

                        //        case "DATESWIPEUP": xControl._xDateSwipeUp(dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        //            break;

                        //        case "DATESWIPEDOWN": xControl._xDateSwipeDown(dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        //            break;

                        //        case "SCREENSHOT": xControl._xScreenShot(dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        //            break;

                        //        case "WAIT":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //                xControl._xWaitUntil(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"]), dsScenario.Tables[0].Rows[b]["Type"].ToString());
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //                xControl._xWaitUntil(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"]), dsScenario.Tables[0].Rows[b]["Type"].ToString());
                        //            break;

                        //        case "DELAY": xControl._xDelay(dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"]));
                        //            break;

                        //        case "CHECKRESULT":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xCheckResultById(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), elementClass[2].ToString());
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xCheckResultByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), elementClass[2].ToString());
                        //                break;
                        //            }
                        //            break;

                        //        case "PRESSKEYENTER": xControl._xPressKeyCode(dsScenario.Tables[0].Rows[b]["Action"].ToString(), 66);
                        //            break;

                        //        case "BUILDPATH": Control.actionPath += xControl._xComposeString(dsScenario.Tables[0].Rows[b], b);
                        //            break;

                        //        case "CHECKEXIST":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xCheckExistByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xCheckExistByXpath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false));
                        //                break;
                        //            }
                        //            break;

                        //        case "GET":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xGetByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xGetElementByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            break;

                        //        case "POST":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xPostByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["Value"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xPostByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Value"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            break;

                        //        case "BACK": xControl._xBackDriver(dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        //            break;

                        //        case "SWIPENOTIFICATION": xControl._xSwipeNotificationBar(dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        //            break;

                        //        case "CHECKTOAST": xControl._xCheckToast(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString());
                        //            break;

                        //        //Forbidden only few devices supported, bad effect is disconecting adb driver to appium server.
                        //        case "NETWORK": xControl._xNetworkChange(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        //            break;

                        //        case "DELETETEXT":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xDeleteTextByElementId(dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xDeleteTextByXPath(xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //                break;
                        //            }
                        //            break;

                        //        case "COMPAREIMAGE": xControl._xScreenShot(dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        //            break;

                        //        case "CHECKCOLOR": xControl._xScreenShot(dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        //            break;

                        //        case "ZOOM":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xZoomByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xZoomByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false));
                        //            }
                        //            break;

                        //        case "LONGCLICK":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xLongPressByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), 1000);
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xLongPressByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), 1000);
                        //            }
                        //            break;

                        //        case "SORT":
                        //            xControl._xCheckSorted(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        //            break;

                        //        case "CHECKACTIVEDIRECTORY":
                        //            xControl._xCheckUserAD(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        //            break;

                        //        case "CHECKSCREENORIENTATION":
                        //            xControl._xCheckPortrait(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString());
                        //            break;

                        //        case "RATE":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xRateByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xRateByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), Convert.ToInt32(dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        //            }
                        //            break;

                        //        case "CHECKNOTEXIST":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xCheckNotExistByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xCheckNotExistByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false));
                        //                break;
                        //            }
                        //            break;

                        //        //Add By Ferie
                        //        case "DBGET":
                        //            _xDBControl._xDBGet(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), dsScenario.Tables[0].Rows[b]["SQL"].ToString(), dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString());
                        //            break;

                        //        case "DBPOST":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                _xDBControl._xDBPostByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["Value"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                _xDBControl._xDBPostByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Value"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        //                break;
                        //            }
                        //            break;

                        //        case "DBEXECUTE":
                        //            _xDBControl._xDBExec(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), dsScenario.Tables[0].Rows[b]["SQL"].ToString());
                        //            break;

                        //        case "CHECKENABLED":
                        //            if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        //            {
                        //                xControl._xCheckEnabledByElementId(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString());
                        //                break;
                        //            }
                        //            else if (dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        //            {
                        //                xControl._xCheckEnabledByXPath(dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(dsScenario.Tables[0].Rows[b], b, false), dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString());
                        //                break;
                        //            }
                        //            break;

                        //        default: xControl._xNoActionSupported(dsScenario.Tables[0].Rows[b]["Action"].ToString(), dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        //            break;
                        //    }
                        //}
                        #endregion

                    }

                    if (maxDataCounter > 1)
                    {
                        if (Control.successFlag == 1)
                        {
                            LoopResult[a, d] = LoopResult[a, d] + "Success; ";
                        }
                        else
                        {
                            LoopResult[a, d] = LoopResult[a, d] + "Failed; ";
                        }
                        xConfiguration._xUpdateDataResult(dsConfiguration.Tables[0].Rows[a]["Data"].ToString(), dsData, d, LoopResult[a, d]);
                    }

                    //Add By Ferie
                    UpdateDetailScenario(Control.successFlag.ToString(), count, 3);
                    if (totalFailed >= 0 && totalSuccess >= 0)
                    {
                        lblTotalFailed.Text = totalFailed.ToString();
                        lblTotalSuccess.Text = totalSuccess.ToString();
                        Application.DoEvents();
                    }

                    if (txtLoopingCount.Text == "1" && TotalDataScenario > 1 || txtLoopingCount.Text != "1" && TotalDataScenario > 1)
                    {
                        progressBar1.Value = count;
                        progressBar1.Maximum = (TotalDataScenario * LoopingCount);
                    }
                    lblPercentScenario.Text = (Math.Round(((double)progressBar1.Value / (double)progressBar1.Maximum), 2) * 100).ToString() + "%";
                    Application.DoEvents();
                    Control.myDictionary.Clear();

                }//Batas Loop "d"

                if (Control.successFlag == 1)
                {
                    Result[a] = Result[a] + "Success; ";
                }
                else
                {
                    Result[a] = Result[a] + "Failed; ";
                }
                Control.successFlag = 1;

                //Add by Ferie
                if (txtLoopingCount.Text == "1" && TotalDataScenario == 1)
                {
                    progressBar1.Value = a+1;
                    progressBar1.Maximum = dsConfiguration.Tables[0].Rows.Count;
                }
                else if (txtLoopingCount.Text != "1" && TotalDataScenario == 1)
                {
                    progressBar1.Value = count;
                    progressBar1.Maximum = dsConfiguration.Tables[0].Rows.Count*LoopingCount;
                }
                
                lblPercentScenario.Text = (Math.Round(((double)progressBar1.Value / (double)progressBar1.Maximum), 2) * 100).ToString() + "%";
                Application.DoEvents();

                #region CSV
                /* CSV
                    foreach (string line in testCase)
                    {
                        string[] testCommand = line.Split(',');
                        switch (testCommand[1])
                        {
                            case "Write": xControl._xWriteByElementId(testCommand[4].ToString(), testCommand[5].ToString());
                                break;
                            case "Click": xControl._xClickByElementId(testCommand[4].ToString());
                                break;
                            case "Screenshot": xControl._xScreenShot();
                                break;
                        }
                    }
                //xControl._xWriteByElementId("com.aab.mobilesalesnative:id/etUser", "abcd@asuransiastra.com");
                //xControl._xWriteByElementId("com.aab.mobilesalesnative:id/etPassword", "qwerty1234");
                //xControl._xClickByElementId("com.aab.mobilesalesnative:id/btnLogin");
                xControl._xScreenShot();
                xControl._xWaitUntil("//android.widget.ImageView[contains(@resource-id, 'android:id/up')]", 500);
                 * */
                #endregion

                xConfiguration._xUpdate(dataSourceConfiguration, "Scenario", dsConfiguration.Tables[0].Rows[a]["No"].ToString(), dsConfiguration.Tables[0].Rows[a]["Scenario"].ToString(), Result[a]);
                dsScenario.Clear();
            }
            dsConfiguration.Clear();
        }

        private void txtLoopingCount_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) )
            {
                e.Handled = true;
            }
        }
    }
}
