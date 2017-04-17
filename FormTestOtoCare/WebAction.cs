using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace FormTestOtoCare
{
    class WebAction
    {
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";

        public static string unaclass = "";

        private static int waitSeconds, delayTime;

        public static int flagAlternateMaster = 1;

        public WebAction()
        {
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            waitSeconds = Convert.ToInt32(ini.GetValue("timeout", "file"));
            delayTime = Convert.ToInt32(ini.GetValue("delaytime", "WebCapabilities"));
        }

        public void StartTestAction(int b, int d)
        {
            DBControl xDBControl = new DBControl();
            WebControl xControl = new WebControl();
            //Control.successFlag = 1;
            flagAlternateMaster = 1;
            if (!FormTest.dsScenario.Tables[0].Columns.Contains("Skip") || (FormTest.dsScenario.Tables[0].Columns.Contains("Skip") && !FormTest.dsScenario.Tables[0].Rows[b]["Skip"].ToString().ToUpper().Equals("YES")))
            {
                Control.startTime = DateTime.Now;
                //Pengecekkan Looping(Compare nilai pada tiap iterasi)
                if (FormTest.dsScenario.Tables[0].Rows[b]["Value"] != null && FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()[0].Equals('$'))
                {
                    FormTest.dsScenario.Tables[0].Rows[b]["Value"] = FormTest.dsData.Tables[0].Rows[d][FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString().Substring(1, FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString().Length - 1)];
                }
                if (FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] != null && FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString()[0].Equals('$'))
                {
                    FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] = FormTest.dsData.Tables[0].Rows[d][FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Substring(1, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Length - 1)];
                }
                if (FormTest.dsScenario.Tables[0].Rows[b]["Key"] != null && FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString()[0].Equals('$'))
                {
                    FormTest.dsScenario.Tables[0].Rows[b]["Key"] = FormTest.dsData.Tables[0].Rows[d][FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Substring(1, FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Length - 1)];
                }

                //Tambahin logic untuk cek result expectation with parameter
                //Pengecekkan query dengan parameter result variabel
                if (FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString().Contains("#"))
                {
                    int count = FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString().Split('#').Length - 1;
                    string tempSQL = FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString();
                    for (int r = 0; r < count; r++)
                    {
                        string OldVar = "", NewVar = "";
                        int startIndex = tempSQL.IndexOf('#');
                        if (tempSQL[startIndex - 1].ToString() == "'")
                        {
                            int endIndex = tempSQL.IndexOf("'", startIndex);
                            OldVar = tempSQL.Substring(startIndex, endIndex - startIndex);
                            //Pengecekkan apakah key terdapat di dictionary atau tidak
                            string checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                            if (checkDictionary != "default")
                            {
                                NewVar = Control.myDictionary[OldVar].ToString();
                                tempSQL = tempSQL.Replace(OldVar, NewVar);
                            }
                        }
                        else if (tempSQL[startIndex - 1].ToString() == " ")
                        {
                            int endIndex = 0;
                            if (!tempSQL.Substring(startIndex, tempSQL.IndexOf(" ", startIndex) - startIndex).Contains(")"))
                                endIndex = tempSQL.IndexOf(" ", startIndex);
                            else
                                endIndex = tempSQL.IndexOf(")", startIndex);
                            OldVar = tempSQL.Substring(startIndex, endIndex - startIndex);
                            //Pengecekkan apakah key terdapat di dictionary atau tidak
                            string checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                            if (checkDictionary != "default")
                            {
                                NewVar = Control.myDictionary[OldVar].ToString();
                                tempSQL = tempSQL.Replace(OldVar, NewVar);
                            }
                        }
                        else if (tempSQL[startIndex - 1].ToString() == "(")
                        {
                            int endIndex = tempSQL.IndexOf(" ", startIndex);
                            OldVar = tempSQL.Substring(startIndex, endIndex - startIndex);
                            //Pengecekkan apakah key terdapat di dictionary atau tidak
                            string checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                            if (checkDictionary != "default")
                            {
                                NewVar = Control.myDictionary[OldVar].ToString();
                                tempSQL = tempSQL.Replace(OldVar, NewVar);
                            }
                        }
                    }
                    FormTest.dsScenario.Tables[0].Rows[b]["SQL"] = tempSQL;
                }

                //Pengecekkan ResultExpectation dengan parameter 
                if (FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Contains("#"))
                {
                    int count = FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Split('#').Length - 1;
                    string tempExpectation = FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString();
                    for (int r = 0; r < count; r++)
                    {
                        string OldVar = "", NewVar = "";
                        int startIndex = tempExpectation.IndexOf('#');
                        if (startIndex > 0)
                        {
                            if (tempExpectation[startIndex - 1].ToString() == "'")
                            {
                                int endIndex = tempExpectation.IndexOf("'", startIndex);
                                OldVar = tempExpectation.Substring(startIndex, endIndex - startIndex);
                                //Pengecekkan apakah key terdapat di dictionary atau tidak
                                string checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                                if (checkDictionary != "default")
                                {
                                    NewVar = Control.myDictionary[OldVar].ToString().Replace(" ", "");
                                    tempExpectation = tempExpectation.Replace(OldVar, NewVar);
                                }
                            }
                            else if (tempExpectation[startIndex - 1].ToString() == " ")
                            {
                                int endIndex = 0;
                                if (!tempExpectation.Substring(startIndex, tempExpectation.IndexOf(" ", startIndex) - startIndex).Contains("."))
                                    endIndex = tempExpectation.IndexOf(" ", startIndex);
                                else
                                    endIndex = tempExpectation.IndexOf(".", startIndex);
                                OldVar = tempExpectation.Substring(startIndex, endIndex - startIndex);

                                //Pengecekkan apakah key terdapat di dictionary atau tidak
                                string checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                                if (checkDictionary != "default")
                                {
                                    NewVar = Control.myDictionary[OldVar].ToString().Replace(" ", "");
                                    tempExpectation = tempExpectation.Replace(OldVar, NewVar);

                                }
                            }
                        }
                        else
                        {
                            OldVar = tempExpectation;
                            string checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                            if (checkDictionary != "default")
                            {
                                NewVar = Control.myDictionary[OldVar].ToString().Replace(" ", "");
                                tempExpectation = tempExpectation.Replace(OldVar, NewVar);
                            }
                        }
                    }
                    FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] = tempExpectation;
                }


                if (FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Contains("#"))
                {
                    string OldVar="", NewVar="", tempValue="", checkDictionary="";

                    OldVar = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString();
                    tempValue = OldVar;
                    checkDictionary = Control.myDictionary.ContainsKey(OldVar) ? Control.myDictionary[OldVar] : "default";
                    if (checkDictionary != "default")
                    {
                        NewVar = Control.myDictionary[OldVar].ToString().Replace(" ", "");
                        tempValue = tempValue.Replace(OldVar, NewVar);
                    }
                    FormTest.dsScenario.Tables[0].Rows[b]["Key"] = tempValue;
                }

                if (FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] != null && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString()[0].Equals('$'))
                {
                    DataRow dr = FormTest.dsMasterResourceId.Tables[0].Select("Variable = '" + FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString() + "'")[0];
                    FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["WebResourceId"].ToString();
                    //string tempXPath = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString();
                    //FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Replace("#Key", "'0'");
                    if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().Equals(""))
                        FormTest.dsScenario.Tables[0].Rows[b]["Type"] = "PATH";
                    else
                        FormTest.dsScenario.Tables[0].Rows[b]["Type"] = FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper();
                    FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Replace("[Key]", "[" + FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString() + "]");
                }

                switch (FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper())
                {
                    //definition of each action by its type
                    case "GOTO":
                        xControl.GoTo(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "ALERTAUTHENTICATION":
                        xControl.AlertAuthentication(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        break;
                    case "ACCEPTALERT":
                        xControl.AcceptAlert(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), delayTime);
                        break;
                    case "WRITE":
                        xControl.Write(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "WRITEWITHJSINJECT":
                        xControl.WriteWithJSInject(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        break;
                    case "SELECTFILEUPLOAD":
                        xControl.SelectFileUpload(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SETDATETIME":
                        xControl.SetDateTime(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SETDATETIMEWITHOUTENTER":
                        xControl.SetDateTimeWithoutEnter(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SETTIME":
                        xControl.SetTime(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "DELETETEXT":
                        xControl.DeleteText(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString());
                        break;
                    case "CLICK":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl.ClickByPath(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("CSS"))
                        {
                            xControl.ClickByCss(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;
                    case "DOUBLECLICK":
                        xControl.DoubleClick(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "LOOPCLICK":
                        xControl.LoopClick(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        break;
                    case "CLICKBYKEY":
                        xControl.ClickByKey(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        break;
                    case "HOVER":
                        xControl.Hover(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SELECT":
                        xControl.Select(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SUBMIT":
                        xControl.Submit(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "DELAY":
                        xControl.Delay(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), int.Parse(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()));
                        break;
                    case "WAIT":
                        xControl.Wait(int.Parse(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), 0);
                        break;
                    case "GET":
                        xControl._xGetElementByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "GETBYKEY":
                        xControl._xGetElementByKeyXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "POST":
                        xControl._xPostByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "BACK":
                        xControl.Back();
                        break;
                    case "FORWARD":
                        xControl.Forward();
                        break;
                    case "STOP":
                        xControl.Stop();
                        break;
                    case "REFRESH":
                        xControl.Refresh();
                        break;
                    case "SCROLLTO":
                        xControl.ScrollTo(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), waitSeconds, delayTime, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SORT":
                        xControl._xCheckSorted(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        break;
                    case "SORTDATETIME":
                        xControl._xCheckSortedDateTime(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        break;
                    case "SORTDROPDOWN":
                        xControl._xCheckSortedDropDown(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKCOLUMNRESULT":
                        xControl._xCheckColumnResult(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKRESULTDROPDOWN":
                        xControl._xCheckResultDropDown(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKCOLUMNMULTILIKE":
                        xControl._xCheckColumnResultContain(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKCOLUMNLIKE":
                        xControl._xCheckColumnResultContainOnlyOneExpectation(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKRESULT": 
                        xControl._xCheckResultByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKROWCOUNT":
                        xControl._xCheckRowCountByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKCOLOR":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("CSS"))
                        {
                            xControl._xCheckResultColorByCSS(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                            break;
                        }
                            break;
                    case "CHECKRESULTBYKEY":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckResultByKeyXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("CSS"))
                        {
                            xControl._xCheckResultByKeyCSS(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                            break;
                        }
                        break;
                    case "CHECKRESULTDATETIME":
                            xControl._xCheckResultDateTimeByKeyXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                            break;
                    case "CHECKNOTEXIST":
                        xControl.CheckNotExistByXpath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKMENUNOTEXIST":
                        xControl.CheckMenuNotExistByXpath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKEXIST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl.CheckExistByXpath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(),FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("CSS"))
                        {
                            xControl.CheckExistByCss(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;
                    case "CHECKENABLED":
                        /*if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl.CheckEnabledByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(),FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {*/
                            xControl.CheckEnabledByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(),FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        //}
                       // break;
                    case "CHECKSELECTED":
                        /*if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl.CheckSelectedtById(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(),FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {*/
                            xControl.CheckSelectedtByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(),FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        //}
                        //break;
                    case "CHECKREADONLY":
                        /*if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl.CheckReadOnlyByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {*/
                            xControl.CheckReadOnlyByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        //}
                        //break;
                    case "DBGET":
                        xDBControl._xDBGet(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "DBGETCOLUMN":
                        xDBControl._xDBGetColumn(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "DBPOST":
                        /*if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xDBControl._xDBPostByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {*/
                            xDBControl._xDBPostByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        //}
                        //break;

                    case "DBEXECUTE":
                        xDBControl._xDBExec(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    default: xControl.NoActionSupported(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        break;
                }
            }
        }
    }
}
