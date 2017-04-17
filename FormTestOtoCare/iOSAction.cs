using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace FormTestOtoCare
{
    class iOSAction
    {
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";

        private static string dataSourceBaseAlternateClass, language;

        private static int waitSeconds;

        public iOSAction()
        {
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            dataSourceBaseAlternateClass = ini.GetValue("baseAlternateClass", "file");
            waitSeconds = Convert.ToInt32(ini.GetValue("timeout", "file"));
            language = ini.GetValue("language", "file");
        }

        public void StartTestAction(int b, int d)
        {
            iOSControl xControl = new iOSControl();
            DBControl xDBControl = new DBControl();
            Configuration xConfig = new Configuration();

            if (!FormTest.dsScenario.Tables[0].Columns.Contains("Skip") || (FormTest.dsScenario.Tables[0].Columns.Contains("Skip") && !FormTest.dsScenario.Tables[0].Rows[b]["Skip"].ToString().ToUpper().Equals("YES")))
            {
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
                if (FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] != null && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString()[0].Equals('$'))
                {
                    DataRow dr = FormTest.dsMasterResourceId.Tables[0].Select("Variable = '" + FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString() + "'")[0];
                    //Pengecekkan ResourceID yang mengandung [Key]
                    if (dr["IosResourceId"].ToString().Contains("[Key]"))
                    {
                        int n;
                        bool isNumeric = int.TryParse(FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), out n);

                        if (isNumeric == true)
                        {
                            string tempResourceId = dr["IosResourceId"].ToString();
                            dr["IosResourceId"] = tempResourceId;
                            tempResourceId = tempResourceId.Replace("[Key]", "[" + FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString() + "]");
                            FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = tempResourceId;
                        }

                        if (FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Contains("#"))
                        {
                            string[] tempConcat = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Split(new string[] { "'" }, StringSplitOptions.None);
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

                        if (FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().ToUpper().Contains("@TEXT"))
                        {
                            string tempKey = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString();
                            string tempResourceId = dr["IosResourceId"].ToString();
                            FormTest.dsScenario.Tables[0].Rows[b]["Key"] = FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString().Replace("@text", "@value");
                            dr["IosResourceId"] = dr["IosResourceId"].ToString().Replace("[Key]", "[" + FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString() + "]");

                            //Pengecekkan jika attribute @value tidak valid maka diubah menjadi @name
                            try
                            {
                                if (iOSCapabilities.driver.FindElement(By.XPath(dr["IosResourceId"].ToString())).GetAttribute("value").ToString()=="" ||
                                    iOSCapabilities.driver.FindElement(By.XPath(dr["IosResourceId"].ToString())).GetAttribute("value").ToString().Length <=0 ||
                                    iOSCapabilities.driver.FindElement(By.XPath(dr["IosResourceId"].ToString())).GetAttribute("value").ToString()==null)
                                {
                                    dr["IosResourceId"] = tempResourceId;
                                    tempKey = tempKey.ToString().Replace("@text", "@name");
                                    tempResourceId = tempResourceId.Replace("[Key]", "[" + tempKey + "]");
                                    FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = tempResourceId;
                                }
                                else
                                {
                                    FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["IosResourceId"];
                                    dr["IosResourceId"] = tempResourceId;
                                }
                            }
                            catch (Exception ex)
                            {
                                dr["IosResourceId"] = tempResourceId;
                                tempKey = tempKey.Replace("@text", "@name");
                                tempResourceId = tempResourceId.Replace("[Key]", "[" + tempKey + "]");
                                FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = tempResourceId;
                            }
                        }
                    }
                    else
                    {
                        FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["IosResourceId"].ToString();
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

                if (!FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("IOS") && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] != null && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString().Length > 0 && FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString()[0].Equals('$'))
                {
                    //flagAlternateMaster = 0;
                    DataRow dr = FormTest.dsMasterResourceId.Tables[0].Select("Variable = '" + FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString() + "'")[0];
                    if (dr["IosResourceId"].ToString().Contains("~"))
                    {
                        DataRow[] drDictionary = null;
                        string[] splitIosResourceIdKey = dr["IosResourceId"].ToString().Split(new string[] { "'" }, StringSplitOptions.None);
                        drDictionary = FormTest.dsDictionary.Tables[0].Select("Variable = '" + splitIosResourceIdKey[1] + "'");
                        if (drDictionary.Count() > 0)
                        {
                            if (language.ToUpper().Equals("ENGLISH"))
                                FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["IosResourceId"].ToString().Replace(splitIosResourceIdKey[1], drDictionary[0]["EnglishText"].ToString());
                            else if (language.ToUpper().Equals("INDONESIAN"))
                                FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["IosResourceId"].ToString().Replace(splitIosResourceIdKey[1], drDictionary[0]["IndonesianText"].ToString());
                        }
                    }
                    else
                    {
                        FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"] = dr["IosResourceId"].ToString();
                    }
                }

                #region Base and Alternate Class
                //string tempChecking = "";
                //if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH") && FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().ToUpper() != null && FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().Length > 1)
                //{
                //    tempChecking = xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, true);
                //}
                //if (!FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("WAIT") && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST") && (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH")) && (FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().Length > 1) && !Control._xCheckBaseClass(tempChecking, waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                //{
                //    int c, flagAlternate = 0;
                //    FormTest.dsBaseAlternateClass = xConfiguration._xReadAlternateClass(dataSourceBaseAlternateClass, FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString());
                //    if (FormTest.dsBaseAlternateClass.Tables[0].Rows.Count > 0)
                //    {
                //        string tempClass = FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString();
                //        for (c = 0; c < FormTest.dsBaseAlternateClass.Tables[0].Rows.Count; c++)
                //        {
                //            FormTest.dsScenario.Tables[0].Rows[b]["Class"] = FormTest.dsBaseAlternateClass.Tables[0].Rows[c]["AlternateClass"].ToString();
                //            tempChecking = xControl._xCheckBuildPath(FormTest.dsScenario.Tables[0].Rows[b], b, true);
                //            if (Control._xCheckBaseClass(tempChecking, waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString()))
                //            {
                //                flagAlternate = 1;
                //                break;
                //            }
                //        }
                //        if (c == FormTest.dsBaseAlternateClass.Tables[0].Rows.Count && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST"))
                //        {
                //            FormTest.dsScenario.Tables[0].Rows[b]["Class"] = tempClass;
                //            xControl._xNoClassSupported(FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString());
                //        }
                //        else if (flagAlternate != 1 && !FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper().Equals("CHECKNOTEXIST"))
                //        {
                //            xControl._xNoClassSupported(FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString());
                //        }
                //    }
                //    else
                //    {
                //        xControl._xNoClassOrElement(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString());
                //    }
                //}
                #endregion

                string[] elementClass = FormTest.dsScenario.Tables[0].Rows[b]["Class"].ToString().Split(Control.delimiterChar);
                switch (FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString().ToUpper())
                {
                    //definition of each action by its type
                    case "CLICK":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xClick(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "TAP":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xTap(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "WRITE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xWrite(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "DELETETEXT":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xDeleteText(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "SORT":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckSorted(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        }
                        break;

                    case "SORTDATE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckSortedDate(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString());
                        }
                        break;
                    #region delta action

                    case "WRITEEMPTYSPACE":
                        xControl._xWriteEmptySpaceByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "WRITEWITHOUTHIDE":
                        xControl._xWriteWithoutHide(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "GETWITHSUBSTRING":
                        xControl._xGetElementByXPathSubstring(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKENABLED":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckEnabledByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "CHECKEXIST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckExistByXpath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "CHECKNOTEXIST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckNotExistByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "CHECKNOTEXISTBYSIZE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xCheckNotExistBySize(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "CHECKRESULT":
                        xControl._xCheckResultByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "WRITEWITHCHECKRESULT":
                        xControl._xWriteWithCheckResultByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKRESULTDATETIME":
                        xControl._xCheckResultDateTime(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKRESULTEMPTY":
                        xControl._xCheckResultEmptyByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKRESULTVARIABLE":
                        xDBControl._xCheckResultVariableByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SETDATETIME":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xSetDateTime(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "SCREENSHOT": xControl._xScreenShot(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPEDOWNNOTIFICATION": xControl._xSwipeDownNotificationBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPEUPNOTIFICATION": xControl._xSwipeUpNotificationBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPE": xControl._xSwipe(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPEUP": xControl._xSwipeUp(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPEDOWN": xControl._xSwipeDown(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPEHORIZONTALBAR": xControl._xSwipeHorizontalBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPEVERTICALBAR": xControl._xSwipeVerticalBar(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "SWIPESIDEMENU": xControl._xSwipeSideMenu(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;
                    case "CHECKCHECKED": xControl._xCheckCheckedByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "DELAY": xControl._xDelay(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"]));
                        break;

                    case "WAIT": xControl._xWaitUntil(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"]), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), 0);
                        break;

                    case "ZOOM": xControl._xZoomByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "CHECKSCREENORIENTATION":
                        xControl._xCheckPortrait(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultExpectation"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPERIGHT": xControl._xSwipeRight(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "SWIPELEFT": xControl._xSwipeLeft(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), Convert.ToInt32(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString()), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "RESETAPP": xControl._xResetAppByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "DISMISS": xControl._xDismissByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    #endregion
                    case "DBGET":
                        xDBControl._xDBGet(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "DBPOST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xDBPostByXPath(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "DBEXECUTE":
                        xDBControl._xDBExec(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ConnectionString"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["SQL"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        break;

                    case "GET":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xGet(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "GETSIZE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xGetSize(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResultVariable"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "POST":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xPost(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "DELETELISTWITHSWIPE":
                        if (FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString().ToUpper().Equals("PATH"))
                        {
                            xControl._xSwipeListToDelete(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Type"].ToString(), waitSeconds, FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString());
                        }
                        break;

                    case "SWITCHPORT":
                        xControl.SwitchPort(FormTest.dsScenario.Tables[0].Rows[b]["Value"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Key"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["Description"].ToString(), 5000);
                        break;

                    default: xControl._xNoActionSupported(FormTest.dsScenario.Tables[0].Rows[b]["Action"].ToString(), FormTest.dsScenario.Tables[0].Rows[b]["ResourceId"].ToString());
                        break;
                }
            }
        }
    }
}
