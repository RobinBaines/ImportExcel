//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for DoImportTextFile.cs
//Function: Import text files into the i_staging tabel and then call p_process_staging_table to process the data. 
//20190704 DoImportTextFile: The file name is put in the sheet field and that is max 50 characters.
//20191118 DoImportTextFile: A problem was caused by an incorrect line at the end of a file.
//Avoid an exception by checking whether the line was correctly split.
//Created Jan 2018.
//------------------------------------------------
using System;
using System.IO;
using Shared.Settings;
using Shared.Utilities;
namespace ImportExcel
{
    class DoImportTextFile : AnEventBase
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strReel"></param>
        /// <param name="strUnused"></param>
        /// <returns></returns>
         public bool TheActionInsert(string strReel, string strUnused)
        {
            int MAXFIELDSPERROW = 100;
            if (!Directory.Exists(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder)))
                Directory.CreateDirectory(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder));

            var directoryInfo = new DirectoryInfo(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder));
            foreach (var file in directoryInfo.GetFiles("*.csv"))
            {
                try
                {
                    StreamReader fs = new StreamReader(file.FullName);
                    String line;
                    bool bFirstTime = true;
                    int rows=0;
                    int recordLength = 0;
                    string [] split;
                    CATDataSet.i_stagingDataTable dtStaging = new CATDataSet.i_stagingDataTable();

                    while((line = fs.ReadLine()) != null)  
                    {
                     split = line.Split(new Char[] { ';', '\t'});
                     if (bFirstTime)
                     {
                         bFirstTime=false;
                          Logging.Log("Processing file " + file.FullName);
                         recordLength=split.Length;
                         DbManager.CreateStagingTable(MAXFIELDSPERROW);
                     }

                     //20190704 The file name is put in the sheet field and that is max 50 characters.
                     string Values = "SELECT '" + file.Name + "', " + rows.ToString();

                     if ( file.Name.Length > 50)
                        {
                        //Values =  file.Name.Substring(0, 50) + ", " + rows.ToString();
                        Values = "SELECT '" + file.Name.Substring(0, 50) + "', " + rows.ToString();
                        }

                        int i=0;
                        while (i < recordLength)
                                        {
                                        //20191118 A problem was caused by an incorrect line at the end of a file. 
                                        //The processing of the incorrect line should be handled from the p_process_staging_table.
                                        if (i < split.Length)
                                            {
                                               Values = Values + ", '" + split[i].Replace("'", "'\'") + "'";
                                                
                                            }
                                        else
                                        Values = Values + ",'' ";
                                            i++;
                                        }
                        //if there were less fields than MAXFIELDSPERROW fill out.
                        while (i < MAXFIELDSPERROW)
                        {
                           Values = Values + ", '\'";

                            i++;
                        }
                        DbManager.InsertStagingRow(MAXFIELDSPERROW, Values);
                      
                    rows++;
                    }  
                    fs.Close();



                        Logging.Log("Selected " + rows.ToString() + " rows of data from file " + file.Name);
                        Logging.Log("Processing data from file " + file.Name);
                        int iCount = DbManager.ProcessStagingTable(file.Name);
                        Logging.Log("Processed " + iCount.ToString() + " records from file " + file.Name);
                        try
                        {
                            if (File.Exists(TheSettings.Instance.ExcelImportedFolder + file.Name))
                                File.Delete(TheSettings.Instance.ExcelImportedFolder + file.Name);
                            file.MoveTo(TheSettings.Instance.ExcelImportedFolder + file.Name);
                            Logging.Log("Moving file " + file.Name + " to " + TheSettings.Instance.ExcelImportedFolder);
                            Logging.Log(file.Name + " has been Imported.");
                        }
                        catch (Exception ex)
                        {
                            Logging.Log("Problem moving file in TheAction. " + ex.Message);
                        }
                    }
                catch { Console.WriteLine(file.FullName + " has been found but cannot be opened yet."); }
            }
            return true;
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
            if (!Directory.Exists(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder)))
                Directory.CreateDirectory(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder));

            var directoryInfo = new DirectoryInfo(Path.GetDirectoryName(TheSettings.Instance.ExcelImportFolder));
            foreach (var file in directoryInfo.GetFiles("*.csv"))
            {
                try
                {
                    StreamReader fs = new StreamReader(file.FullName);
                    String line;
                    bool bFirstTime = true;
                    int rows = 0;
                    int recordLength = 0;
                    string[] split;
                    CATDataSet.i_stagingDataTable dtStaging = new CATDataSet.i_stagingDataTable();

                    while ((line = fs.ReadLine()) != null)
                    {
                        split = line.Split(new Char[] { ';', '\t' });
                        if (bFirstTime)
                        {
                            bFirstTime = false;
                            Logging.Log("Processing file " + file.FullName);
                            recordLength = split.Length;
                            DbManager.CreateStagingTable(MAXFIELDSPERROW);
                        }

                        //20190704 The file name is put in the sheet field and that is max 50 characters.
                        string Values = file.Name + "," + rows.ToString();
                        if (file.Name.Length > 50)
                        {
                            Values = file.Name.Substring(0, 50) + "," + rows.ToString();
                        }

                        int i = 0;
                        while (i < recordLength)
                        {
                            //20191118 A problem was caused by an incorrect line at the end of a  file.
                            //The processing of the incorrect line should be handled from the p_process_staging_table.
                            if (i < split.Length)
                            {
                                Values = Values + "," + split[i].Replace("'", "'\'");
                            }
                            else
                                Values = Values + ",";
                            i++;
                        }
                        //if there were less fields than MAXFIELDSPERROW fill out.
                        while (i < MAXFIELDSPERROW)
                        {
                                 Values = Values + ",";
                            i++;
                        }

                        //Making a larger sql statement just slows things down.
                        //Making one connection at the start of insert has no effect.

                        DbManager.InsertStagingDataTableRow(dtStaging, MAXFIELDSPERROW, Values);
                        rows++;
                    }

                    fs.Close();
                    Logging.Log("Selected " + rows.ToString() + " rows of data from file " + file.Name);
                    Logging.Log(file.Name + " Starting Bulk Insert.");
                    try
                    {
                        DbManager.BulkInsert(dtStaging, "i_staging");
                        Logging.Log(file.Name + " Completed Bulk Insert.");
                    }
                    catch (Exception ex)
                    {
                        Logging.Log("Problem during Bulk Insert. " + ex.Message);
                    }

                                      
                    Logging.Log("Processing data from file " + file.Name);
                    int iCount = DbManager.ProcessStagingTable(file.Name);
                    Logging.Log("Processed " + iCount.ToString() + " records from file " + file.Name);
                    try
                    {
                        if (File.Exists(TheSettings.Instance.ExcelImportedFolder + file.Name))
                            File.Delete(TheSettings.Instance.ExcelImportedFolder + file.Name);
                        file.MoveTo(TheSettings.Instance.ExcelImportedFolder + file.Name);
                        Logging.Log("Moving file " + file.Name + " to " + TheSettings.Instance.ExcelImportedFolder);
                        Logging.Log(file.Name + " has been Imported.");
                    }
                    catch (Exception ex)
                    {
                        Logging.Log("Problem moving file in TheAction. " + ex.Message);
                    }
                }
                catch { Logging.Log(file.FullName + " has been found but cannot be opened yet."); }
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
