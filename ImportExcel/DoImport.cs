//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for DoImport.cs
//Function: Import sheets from Excel files into the i_staging tabel and then call p_process_staging_table to process the data.
//Uses ACE Oledb.12.0.
//Created Jan 2018.
//------------------------------------------------
using System;
using System.IO;
using Shared.Settings;
using Shared.Utilities;
using System.Data.OleDb;
using System.Data;
namespace ImportExcel
{
    class DoImport : AnEventBase
    {
        override public string TheCondition(ref string strUnused)
        {
            try
            {
            }
            catch (Exception ex)
            {
                Logging.Log("A general error occurred in Main: " + ex);
            }
            return "Always True";
        }


        public bool TestResult()
        {
            return true;
         }

        private void MoveTheFile(FileInfo file)
        {
            if (File.Exists(TheSettings.Instance.ExcelImportedFolder + file.Name))
                File.Delete(TheSettings.Instance.ExcelImportedFolder + file.Name);
            file.MoveTo(TheSettings.Instance.ExcelImportedFolder + file.Name);
            Logging.Log("Moving file " + file.Name + " to " + TheSettings.Instance.ExcelImportedFolder);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strReel"></param>
        /// <param name="strUnused"></param>
        /// <returns></returns>
        override public bool TheAction(string strReel, string strUnused)
        {
            int MAXFIELDSPERROW = 100;
            int MAXSHEETS = 100;

            if (!Directory.Exists(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder)))
                Directory.CreateDirectory(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder));

            var directoryInfo = new DirectoryInfo(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder));
            foreach (var file in directoryInfo.GetFiles("*.xls*"))
            {
                try
                {
                    FileStream fs = new FileStream(file.FullName, FileMode.OpenOrCreate, FileAccess.Read);
                    if (fs.CanRead)
                    {
                        fs.Close();
                        Logging.Log("Processing file " + file.FullName);
                        string sSheetName = null;
                        //string sLastSheetName = "";
                        string sConnection = null;
                        DataTable dtTablesList = default(DataTable);
                        OleDbCommand oleExcelCommand = default(OleDbCommand);
                        OleDbDataReader oleExcelReader = default(OleDbDataReader);
                        OleDbConnection oleExcelConnection = default(OleDbConnection);
                        sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file.FullName + ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";
                        oleExcelConnection = new OleDbConnection(sConnection);
                        oleExcelConnection.Open();
                        oleExcelCommand = oleExcelConnection.CreateCommand();
                        oleExcelCommand.CommandType = CommandType.Text;
                        dtTablesList = oleExcelConnection.GetSchema("Tables");
                        int tables = 0;
                        int real_tables = 0;

                        DbManager.CreateStagingTable(MAXFIELDSPERROW);

                        while (tables < dtTablesList.Rows.Count && real_tables < MAXSHEETS)
                        {
                            sSheetName = dtTablesList.Rows[tables]["TABLE_NAME"].ToString();
                            Logging.Log("Checking whether to process sheet " + sSheetName.ToString() + " from file " + file.Name);
                            //Logging.Log("Processing sheet " + sSheetName.ToString() + " " + dtTablesList.Rows[tables]["TABLE_TYPE"].ToString() + " from file "+ file.Name);

                            //These are undefined.
                            //Console.WriteLine(dtTablesList.Rows[tables]["TABLE_GUID"].ToString());
                            //    Console.WriteLine( dtTablesList.Rows[tables]["TABLE_CATALOG"].ToString());
                            //    Console.WriteLine(dtTablesList.Rows[tables]["TABLE_SCHEMA"].ToString());
                            //     Console.WriteLine(dtTablesList.Rows[tables]["DESCRIPTION"].ToString());

                            if (sSheetName.EndsWith("$") || sSheetName.EndsWith("$'"))
                            {
                                if (!string.IsNullOrEmpty(sSheetName))
                                {
                                    Logging.Log("Selecting data from sheet " + sSheetName.ToString() + " from file " + file.Name);
                                    oleExcelCommand.CommandText = "Select * From [" + sSheetName + "]";
                                    oleExcelReader = oleExcelCommand.ExecuteReader();
                                    //   nOutputRow = 0;
                                    int rows = 0;
                                    while (oleExcelReader.Read())
                                    {
                                        int i = 0;
                                        //sSheetName includes '' except if it is the original name for example Sheet1. Could agree that a sheet has to have a name if it is to be included in the import.
                                        if (sSheetName.Contains("'") == false) sSheetName = "'" + sSheetName + "'";
                                        string Values = "SELECT " + sSheetName + ", " + rows.ToString();
                                        while (i < oleExcelReader.FieldCount)
                                        {
                                            if (i < MAXFIELDSPERROW)
                                            {
                                                string sField = oleExcelReader[i].ToString();
                                                sField = sField.Replace("'", "'\'");
                                                if (sField.Length > 100) sField = sField.Substring(0, 100);

                                                Values = Values + ", '" + sField + "'";
                                                if (i==5) Console.WriteLine(oleExcelReader[i].ToString());
                                            }
                                            i++;
                                        }

                                        //if there were less fields than MAXFIELDSPERROW fill out.
                                        while (i < MAXFIELDSPERROW)
                                        {
                                            Values = Values + ", '\'";
                                            i++;
                                        }

                                        //insert the row into the database table.
                                        DbManager.InsertStagingRow(MAXFIELDSPERROW, Values);
                                        rows++;
                                        //Console.WriteLine(r.ToString());
                                    }
                                    oleExcelReader.Close();
                                    Logging.Log("Selected " + rows.ToString() + " rows of data from sheet " + sSheetName.ToString() + " from file " + file.Name);
                                }
                                else
                                {
                                    Logging.Log("Not processing sheet " + sSheetName.ToString());
                                }

                                real_tables++;
                            }
                            else
                            {
                                Logging.Log("Not processing sheet " + sSheetName.ToString());
                            }
                            tables++;
                        }
                        dtTablesList.Clear();
                        dtTablesList.Dispose();
                        oleExcelConnection.Close();
                        Logging.Log("Processing data from file " + file.Name);
                        int iCount = DbManager.ProcessStagingTable(file.Name);
                        Logging.Log("Processed " + iCount.ToString() + " records from file " + file.Name);
                        try
                        {
                            MoveTheFile(file);
                            Logging.Log(file.Name + " has been Imported.");
                        }
                        catch (Exception ex)
                        {
                            Logging.Log("Problem moving file in TheAction. " + ex.Message);
                        }
                    }

                    else fs.Close();

                }
                catch {
                    Console.WriteLine(file.FullName + " has been found but cannot be opened yet.");

                    //20180222
                    MoveTheFile(file);
                    }
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strReel"></param>
        /// <param name="strUnused"></param>
        /// <returns></returns>
        override public string FormatApplicationLog(string strReel, string strUnused)
        {
            return "";
        }
    }

}


