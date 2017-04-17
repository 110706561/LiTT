using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;

namespace FormTestOtoCare
{
    class Configuration
    {
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";

        public static string _xGetConnectionString(string dataSource, string imex)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            props["Data Source"] = dataSource;
            props["Extended Properties"] = "'Excel 12.0;HDR=Yes;IMEX=" + imex.ToString() + ";MAXSCANROWS=0'";
            //props["IMEX"] = "1";
            //props["Mode"] = "ReadWrite";

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Data Source"] = dataSource;
            //props["Extended Properties"] = "'Excel 8.0;HDR=Yes;IMEX=" + imex.ToString() + "'";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        public void _xCreateExcelFile(string dataSource)
        {
            // Create the file using the FileInfo object
            var file = new FileInfo(dataSource);

            Ini ini = new Ini(iniFileLocation);
            ini.Load();

            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(ini.GetValue("applicationName", "file"));
                package.Save();
            }
        }

        #region WriteExcel

        public void WriteExcelFile(string dataSource)
        {
            string connectionString = _xGetConnectionString(dataSource, "0");

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = "CREATE TABLE [table1] (id INT, name VARCHAR, datecol DATE );";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(1,'AAAA','2014-01-01');";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(2, 'BBBB','2014-01-03');";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(3, 'CCCC','2014-01-03');";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE [table1] SET name = 'DDDD' WHERE id = 3;";
                cmd.ExecuteNonQuery();

                conn.Close();
            }
        }

        #endregion

        public void _xCreateSheetExcelFile(string dataSource, string sheetName)
        {
            string connectionString = _xGetConnectionString(dataSource, "0");
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = "CREATE TABLE [" + sheetName + "] (Actions LONGTEXT, Element LONGTEXT, Status LONGTEXT, Result LONGTEXT, Duration LONGTEXT, Description LONGTEXT);";
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        public void _xInsertRowData(string dataSource, string sheetName, string Action, string Elemen, string Status, string Result, string duration, string desc)
        {
            string connectionString = _xGetConnectionString(dataSource, "0");
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.Parameters.Add("@0", OleDbType.LongVarChar).Value = Action.Length < 255 ? Action : Action.Substring(0, 254);
                cmd.Parameters.Add("@1", OleDbType.LongVarChar).Value = Elemen.Length < 255 ? Elemen : Elemen.Substring(0, 254);
                cmd.Parameters.Add("@2", OleDbType.LongVarChar).Value = Status.Length < 255 ? Status : Status.Substring(0, 254);
                cmd.Parameters.Add("@3", OleDbType.LongVarChar).Value = Result.Length < 255 ? Result : Result.Substring(0, 254);
                cmd.Parameters.Add("@4", OleDbType.LongVarChar).Value = duration.Length < 255 ? duration : duration.Substring(0, 254);
                cmd.Parameters.Add("@5", OleDbType.LongVarChar).Value = desc.Length < 255 ? desc : desc.Substring(0, 254);
                cmd.CommandText = "INSERT INTO [" + sheetName + "$] (Actions, Element, Status, Result, Duration, Description) VALUES(@0, @1, @2, @3, @4, @5);";

                cmd.ExecuteNonQuery();
                conn.Close();

            }
            #region comment
            //Action = Action.Length < 255 ? Action : Action.Substring(0, 254);
            //Elemen = Elemen.Length < 255 ? Elemen : Elemen.Substring(0, 254);
            //Status = Status.Length < 255 ? Status : Status.Substring(0, 254);
            //Result = Result.Length < 255 ? Result : Result.Substring(0, 254);
            //Control.insertText += "INSERT INTO " + sheetName + "(Actions, Element, Status, Result) VALUES('" + Action + "', '" + Elemen + "', '" + Status + "', '" + Result + "');";
            #endregion
        }

        public DataSet _xReadExcelFile(string dataSource)
        {

            DataSet ds = new DataSet();
            string connectionString = _xGetConnectionString(dataSource, "1");

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    if (!sheetName.EndsWith("$"))
                        continue;
                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    ds.Tables.Add(dt);
                }
                cmd = null;
                conn.Close();
            }
            return ds;
        }

        public DataSet _xReadAlternateClass(string dataSource, string baseClass)
        {
            DataSet ds = new DataSet();
            string connectionString = _xGetConnectionString(dataSource, "1");

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    if (sheetName.Equals("Class$"))
                    {
                        // Get all rows from the Sheet
                        cmd.CommandText = "SELECT * FROM [" + sheetName + "] WHERE [BaseClass] = '" + baseClass + "'";

                        DataTable dt = new DataTable();
                        dt.TableName = sheetName;
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(dt);
                        ds.Tables.Add(dt);
                    }
                }
                cmd = null;
                conn.Close();
            }
            return ds;
        }

        public void Save(string excelFile)
        {

            // Create the file using the FileInfo object
            var file = new FileInfo(excelFile);
            using (var package = new ExcelPackage(file))
            {
                #region comment
                //// add a new worksheet to the empty workbook
                //ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sales list - " + DateTime.Now.ToShortDateString());

                //// --------- Data and styling goes here -------------- //
                //worksheet.Cells[1, 1].Value = "Company name";
                //worksheet.Cells[1, 2].Value = "Address";
                //worksheet.Cells[1, 3].Value = "Status (unstyled)";

                //worksheet.Cells[2, 1].Value = "Vehicle registration plate";
                //worksheet.Cells[2, 2].Value = "Vehicle brand";
                #endregion
                package.Save();
            }
        }

        public void _xUpdate(string dataSource, string sheetName, string no, string lastFeature, string Result /*, int success, int maxData*/)
        {
            #region comment
            //using (ExcelPackage package = new ExcelPackage(new FileInfo(dataSource)))
            //{
            //    var worksheet = package.Workbook.Worksheets["Scenario"];
            //    if(success == 1)
            //        worksheet.Cells[5, Convert.ToInt32(No)].Value = "Success";
            //    else
            //        worksheet.Cells[5, Convert.ToInt32(No)].Value = "Failed";

            //    package.Save();
            //}
            #endregion
            string connectionString = _xGetConnectionString(dataSource, "0");
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                #region comment
                //if (success == 1)
                //{
                //    //cmd.CommandText = "SELECT * FROM [Scenario$]";
                //    cmd.CommandText = string.Format("UPDATE [Scenario$] SET Result = 'Success' WHERE [Scenario] = @lastfeature AND [No] = @no");
                //}
                //else
                //{
                //    //cmd.CommandText = "SELECT * FROM [Scenario$]";
                //    cmd.CommandText = string.Format("UPDATE [Scenario$] SET Result = 'Failed' WHERE [Scenario] = @lastfeature AND [No] = @no");
                //}
                #endregion
                cmd.CommandText = string.Format("UPDATE [Scenario$] SET Result = @result WHERE [Scenario] = @lastfeature AND [No] = @no");

                cmd.Parameters.AddWithValue("@result", SqlDbType.NText);
                cmd.Parameters["@result"].Value = Result;
                cmd.Parameters.AddWithValue("@lastfeature", SqlDbType.VarChar);
                cmd.Parameters["@lastfeature"].Value = lastFeature;
                cmd.Parameters.AddWithValue("@no", SqlDbType.VarChar);
                cmd.Parameters["@no"].Value = no;
                cmd.ExecuteNonQuery();
                Control.successFlag = 1;
                conn.Close();
            }
        }

        public void _xUpdateDataResult(string dataSource, DataSet ds, int b, string result)
        {
            string connectionString = _xGetConnectionString(dataSource, "0");
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                cmd.CommandText = "UPDATE [Data$] SET Result = @result WHERE [No] = @no";
                cmd.Parameters.AddWithValue("@result", SqlDbType.VarChar);
                cmd.Parameters["@result"].Value = result;
                cmd.Parameters.AddWithValue("@no", SqlDbType.VarChar);
                cmd.Parameters["@no"].Value = ds.Tables[0].Rows[b]["No"].ToString();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        public string _xReplaceSpecialChars(string input)
        {
            char[] specialChar = new char[] { ')', '(', '&', ' ', '\'', '"', '\'', '@', '`', '#', '%', '>', '<', '!', '[', ']', '*', '$', ';', ':', '?', '^', '{', '}', '+', '-', '=', '~', '\\' };
            int pos = -1;
            while ((pos = input.IndexOfAny(specialChar)) != -1)
                input = input.Substring(0, pos) + '_' + input.Substring(pos + 1);

            return input;
        }
    }
}
