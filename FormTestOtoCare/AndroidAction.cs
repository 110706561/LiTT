using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormTestOtoCare
{
    class AndroidAction
    {
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";
        
        private static string dataSourceBaseAlternateClass,language;
        
        public static string unaclass = "";

        private static int waitSeconds,delaySwitchPort;

        public static int flagAlternateMaster = 1;

        public AndroidAction()
        {
            //ConfigurationManager
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            //dataSourceConfiguration = ini.GetValue("dataSourceConfiguration", "file");
            //dataSourceReport = ini.GetValue("dataSourceReport", "file");
            //pictureFile = ini.GetValue("pictureFile", "file");
            dataSourceBaseAlternateClass = ini.GetValue("baseAlternateClass", "file");
            //masterResourceId = ini.GetValue("masterResourceId", "file");
            waitSeconds = Convert.ToInt32(ini.GetValue("timeout", "file"));
            delaySwitchPort = Convert.ToInt32(ini.GetValue("delaySwitchPort", "AndroidCapabilities"));
            language = ini.GetValue("language", "file");
        }

        public void StartTestAction(int b, int d)
        {
            Control xControl = new Control();
            Configuration xConfiguration = new Configuration();
            DBControl xDBControl = new DBControl();
            flagAlternateMaster = 1;
            string[] elementClass;

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
                if (FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] != null && FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString()[0].Equals('#'))
                {
                    FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] = Control.myDictionary[FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString()].ToString();
                }
                if (FormTest.dsScenario.Tables[0].Rows[b]["Key"] != null && FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString()[0].Equals('$'))
                {
                    FormTest.dsScenario.Tables[0].Rows[b]["Key"] = FormTest.dsData.Tables[0].Rows[d][FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Substring(1, FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Length - 1)];
                }

                if (FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Contains("#"))
                {
                    string [] tempConcat = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Split(new string[] { "'" }, StringSplitOptions.None);
                    if (tempConcat.Length == 3)
                    {
                        for (int s = 0; s < tempConcat.Length; s++)
                        {
                            if (tempConcat[s].Contains("#"))
                                tempConcat[s] = Control.myDictionary[tempConcat[s]].ToString();
                        }
                        FormTest.dsScenario.Tables[0].Rows[b]["Key"] = tempConcat[0] + "'" + tempConcat[1] + tempConcat[2] + "'";
                    }
                }

                if (FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Contains("~"))
                {
                    DataRow[] drDictionary = null;
                    string[] tempKey = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Split(new string[] { "'" }, StringSplitOptions.None);
                    drDictionary = FormTest.dsDictionary.Tables[1].Select("Variable = '" + tempKey[1] + "'");
                    if (drDictionary.Count() > 0)
                    {
                        if(language.ToUpper().Equals("ENGLISH"))
                            FormTest.dsScenario.Tables[0].Rows[b]["Key"] = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Replace(tempKey[1], drDictionary[0]["EnglishText"].ToString());
                        else if (language.ToUpper().Equals("INDONESIAN"))
                            FormTest.dsScenario.Tables[0].Rows[b]["Key"] = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Replace(tempKey[1], drDictionary[0]["IndonesianText"].ToString());
                    }
                }

                if (FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString().Contains("~"))
                {
                    DataRow[] drDictionary = null;
                    string tempResultExpectation = FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString();
                    drDictionary = FormTest.dsDictionary.Tables[1].Select("Variable = '" + tempResultExpectation + "'");
                    if (drDictionary.Count() > 0)
                    {
                        if (language.ToUpper().Equals("ENGLISH"))
                            FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] = drDictionary[0]["EnglishText"].ToString();
                        else if (language.ToUpper().Equals("INDONESIAN"))
                            FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"] = drDictionary[0]["IndonesianText"].ToString();
                    }
                }
                //Pengecekkan query dengan parameter result variabel
                if (FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString().Contains("#"))
                {
                    int count = FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString().Split('#').Length - 1;
                    string tempSQL = FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString();
                    for (int r = 0; r < count; r++)
                    {
                        string OldVar="",NewVar="";
                        int startIndex = tempSQL.IndexOf('#');
                        if (tempSQL[startIndex - 1].ToString() == "'")
                        {
                            int endIndex = tempSQL.IndexOf("'", startIndex);
                            OldVar = tempSQL.Substring(startIndex, endIndex-startIndex);
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
                            if (checkDictionary!="default")
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

                if (!FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ANDROID") && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] != null && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString()[0].Equals('$') && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SETDATETIME"))
                {
                    flagAlternateMaster = 0;
                    DataRow dr = FormTest.dsMasterResourceId.Tables[0].Select("Variable = '" + FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString() + "'")[0];
                    if (dr["AndroidResourceId"].ToString().Contains("~"))
                    {
                        DataRow[] drDictionary = null;
                        string[] splitAndroidResourceIdKey = dr["AndroidResourceId"].ToString().Split(new string[] { "'" }, StringSplitOptions.None);
                        drDictionary = FormTest.dsDictionary.Tables[0].Select("Variable = '" + splitAndroidResourceIdKey[1] + "'");
                        if (drDictionary.Count() > 0)
                        {
                            if (language.ToUpper().Equals("ENGLISH"))
                                FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["AndroidResourceId"].ToString().Replace(splitAndroidResourceIdKey[1], drDictionary[0]["EnglishText"].ToString());
                            else if (language.ToUpper().Equals("INDONESIAN"))
                                FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["AndroidResourceId"].ToString().Replace(splitAndroidResourceIdKey[1], drDictionary[0]["IndonesianText"].ToString());
                        }
                    }
                    else
                    {
                        FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["AndroidResourceId"].ToString();
                    }

                    string tempXPath = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString();
                    FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Replace("#Key", "'0'");
                    FormTest.dsScenario.Tables[0].Rows[b]["Type"] = "PATH";
                    string tempUna = "";
                    unaclass = "";

                    if (!FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SORT"))
                    {
                        string[] anu = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Split(new string[] { "//" }, StringSplitOptions.None);
                        for (int x = 1; x < anu.Length; x++)
                        {
                            anu[x] = anu[x].Replace("[Key]", "[" + FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString() + "]");
                            string una = "//" + anu[x];
                            int pFrom = una.IndexOf("//") + "//".Length;
                            int pTo = una.LastIndexOf("[");
                            tempUna = una.Substring(pFrom, pTo - pFrom);
                            string tempCheckingPath = "";
                            //////////////
                            int c, flagAlternate = 0;
                            if (!Control._xCheckBaseClass(unaclass + una, waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                            {
                                FormTest.dsBaseAlternateClass = xConfiguration._xReadAlternateClass(dataSourceBaseAlternateClass, tempUna);
                                if (FormTest.dsBaseAlternateClass.Tables[0].Rows.Count > 0)
                                {
                                    string tempClass = tempUna;
                                    for (c = 0; c < FormTest.dsBaseAlternateClass.Tables[0].Rows.Count; c++)
                                    {
                                        tempUna = FormTest.dsBaseAlternateClass.Tables[0].Rows[c]["AlternateClass"].ToString();
                                        anu[x] = anu[x].Replace(tempClass, tempUna);
                                        tempCheckingPath = unaclass + "//" + anu[x];
                                        if (Control._xCheckBaseClass(tempCheckingPath, waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                                        {
                                            flagAlternate = 1;
                                            break;
                                        }
                                    }

                                    unaclass = tempCheckingPath;
                                    if (c == FormTest.dsBaseAlternateClass.Tables[0].Rows.Count && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST") || flagAlternate != 1)
                                    {
                                        tempUna = tempClass;
                                        xControl._xNoClassSupported(tempUna);
                                        unaclass = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString();
                                    }
                                }
                                else
                                {

                                    if (!FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SETDATETIME") ||
                                        !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("DISMISS") ||
                                        !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("THREADWAIT") ||
                                        !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SWITCHPORT"))
                                    {
                                        xControl._xNoClassOrElement(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString());
                                    }
                                    unaclass = FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString();
                                }
                            }
                            else
                            {
                                unaclass = unaclass + "//" + anu[x];
                            }
                        }
                    }
                        elementClass = tempUna.Split(Control.delimiterChar);
                    
                    if (FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SETDATETIME")||
                        FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("DISMISS")||
                        FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("THREADWAIT")||
                        FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SORT"))
                    {
                        unaclass = tempXPath;
                    }
                }
                else //if(!FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("SETDATETIME")||)
                {
                    string tempChecking = "";
                    if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH") && FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().ToUpper() != null && FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().Length > 1)
                    {
                        tempChecking = xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, true);
                    }

                    if (!FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("WAIT") && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST") && (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH")) && (FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().Length > 1) && !Control._xCheckBaseClass(tempChecking, waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                    {
                        int c, flagAlternate = 0;
                        FormTest.dsBaseAlternateClass = xConfiguration._xReadAlternateClass(dataSourceBaseAlternateClass, FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString());
                        if (FormTest.dsBaseAlternateClass.Tables[0].Rows.Count > 0)
                        {
                            string tempClass = FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString();
                            for (c = 0; c < FormTest.dsBaseAlternateClass.Tables[0].Rows.Count; c++)
                            {
                                FormTest.dsScenario.Tables[0].Rows[b]["Class"] = FormTest.dsBaseAlternateClass.Tables[0].Rows[c]["AlternateClass"].ToString();
                                tempChecking = xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, true);
                                if (Control._xCheckBaseClass(tempChecking, waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                                {
                                    flagAlternate = 1;
                                    break;
                                }
                            }
                            if (c == FormTest.dsBaseAlternateClass.Tables[0].Rows.Count && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST") || flagAlternate != 1)
                            {
                                FormTest.dsScenario.Tables[0].Rows[b]["Class"] = tempClass;
                                xControl._xNoClassSupported(FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString());
                            }

                        }
                        else
                        {
                            xControl._xNoClassOrElement(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString());
                        }
                    }
                        elementClass = FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().Split(Control.delimiterChar);
                 }

                switch (FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper())
                {
                    //definition of each action by its type
                    case "WRITE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xWriteByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xWriteByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "WRITENUMERIC":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xWriteNumericByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "WRITENUMERICWITHOUTCLEAR":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xWriteNumericWithoutClearByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "WRITEEMPTYSPACE":
                        xControl._xWriteEmptySpaceByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CLICK":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xClickByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xClickByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;
                    case "DISMISS":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xDismissByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xDismissByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds,FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;
                    case "SWIPE": xControl._xSwipe(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPEUP": xControl._xSwipeUp(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPEDOWN": xControl._xSwipeDown(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPELEFT": xControl._xSwipeRight(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPERIGHT": xControl._xSwipeLeft(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPEHORIZONTALBAR": xControl._xSwipeHorizontalBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPEVERTICALBAR": xControl._xSwipeVerticalBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "DATESWIPEUP": xControl._xDateSwipeUp(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        break;

                    case "DATESWIPEDOWN": xControl._xDateSwipeDown(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        break;

                    case "SCREENSHOT": xControl._xScreenShot(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "WAIT":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                            xControl._xWaitUntil(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"]), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(),0);
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                            xControl._xWaitUntil(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"]), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(),0);
                        break;

                    case "DELAY": xControl._xDelay(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"]));
                        break;

                    case "CHECKRESULT":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckResultById(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), elementClass[2].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckResultByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), elementClass[2].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "CHECKRESULTEMPTY":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckResultEmptyById(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), elementClass[2].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckResultEmptyByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), elementClass[2].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "CHECKRESULTVARIABLE":
                        xDBControl._xCheckResultVariableByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(),  FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKRESULTDATETIME":
                        xControl._xCheckResultDateTime(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "PRESSKEYENTER": 
                        xControl._xPressKeyCode(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), 66);
                        break;

                    case "BUILDPATH": 
                        Control.actionPath += xControl._xComposeString(FormTest.dsScenario.Tables[0].Rows[b], b);
                        break;

                    case "CHECKEXIST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckExistByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckExistByXpath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "GET":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xGetByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xGetElementByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;
                    case "GETWITHSUBSTRING":
                        xControl._xGetElementByXPathSubstring(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "POST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xPostByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xPostByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "BACK": xControl._xBackDriver(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPENOTIFICATION": xControl._xSwipeNotificationBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKTOAST": xControl._xCheckToast(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString());
                        break;

                    //Forbidden only few devices supported, bad effect is disconecting adb driver to appium server.
                    case "NETWORK": xControl._xNetworkChange(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString());
                        break;

                    case "DELETETEXT":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xDeleteTextByElementId(FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xDeleteTextByXPath(xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "DELETENUMERIC":
                            xControl._xDeleteNumeric(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKCOLOR": xControl._xCheckColorByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "ZOOM":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xZoomByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xZoomByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "LONGCLICK":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xLongPressByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), 1000, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xLongPressByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), 1000, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "SORT":
                        xControl._xCheckSorted(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKACTIVEDIRECTORY":
                        xControl._xCheckUserAD(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKSCREENORIENTATION":
                        xControl._xCheckPortrait(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "RATE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xRateByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xRateByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds);
                        }
                        break;

                    case "CHECKNOTEXIST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckNotExistByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckNotExistByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    //Add By Ferie
                    case "DBGET":
                        xDBControl._xDBGet(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "DBPOST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xDBControl._xDBPostByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xDBControl._xDBPostByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "DBPOSTNUMERIC":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xDBControl._xDBPostNumericByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "DBEXECUTE":
                        xDBControl._xDBExec(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKENABLED":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckEnabledByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckEnabledByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "CHECKCHECKED":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckCheckedByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckCheckedByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;

                    case "CHECKFOCUSABLE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("ELEMENT"))
                        {
                            xControl._xCheckFocusableByElementId(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        else if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckFocusableByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                            break;
                        }
                        break;
                    case "UPDATEDATETIME":
                        xControl._xUpdateDateTimeByXPath(xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "UPDATEDATETIMEBYINTERVAL":
                        xControl._xUpdateDateTimeWithIntervalByXPath(xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SETDATETIME":
                        //xControl._xSetDateTime(xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        xControl._xSetDateTime(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CLEARTEXT":
                        xControl._xClearTextByXPath(xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "WRITEWITHOUTHIDE":
                        xControl._xWriteWithoutHide(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWITCHPORT":
                        xControl.SwitchPort(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(),delaySwitchPort);
                        break;
                    case "THREADWAIT":
                        xControl.ThreadWait(xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, false), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    default: xControl._xNoActionSupported(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        break;
                }
        }
 
    }
}
}
