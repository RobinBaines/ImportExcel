//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for ImportExcel.cs
//Function: 
//Created Jan 2018.
//Notes: 
//Modifications:
//20181009 Increased timeout on sql command in public static int ProcessStagingTable(string FileName) to 6000 seconds after a timeout on a file with 850000 records.
//20190704 DbManager.ProcessStagingTable. For output and return value parameters, the value is set on completion of the SqlCommand and after the SqlDataReader is closed.
//20190704 DoImportTextFile: The file name is put in the sheet field and that is max 50 characters.
//20191118 DoImportTextFile: A problem was caused by an incorrect line at the end of a file. 
//------------------------------------------------
using System;
using System.ServiceProcess;
using System.Threading;
using Shared.Settings;
using Shared.Utilities;
namespace ImportExcel
{
    public partial class ImportExcel : ServiceBase
    {
        private static bool _shouldRun = true;
        public static ImportExcel Instance { get; set; }
        public ImportExcel()
        {
            Instance = this;
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // Start fromout a thread otherwise the service will never start
            StartMailServiceThread();
        }

        protected override void OnStop()
        {
            Logging.Log("Service stopped");
            Stop();
        }

        public void Stop()
        {
            _shouldRun = false;
        }

        public void StartMailServiceThread()
        {
            // Start fromout a thread otherwise the service will never start and always is in the 'starting' mode
            var th = new Thread(StartService) { IsBackground = true };
            th.Start();
        }

        private static void StartService()
        {
            try
            {
                Logging.Log("VERSION 1.1. Service started!");

                // Get settings from xml
                Settings.GetServiceSettings();

                Logging.Log("Settings:");
                foreach (var property in (TheSettings.Instance.GetType()).GetProperties())
                {
                    Logging.Log(property.Name + ": " + property.GetValue(TheSettings.Instance, null));
                }

                //20130417 Set the current directory to the service install directory.
                System.IO.Directory.SetCurrentDirectory(System.AppDomain.CurrentDomain.BaseDirectory);

                CheckForEvent();
            }

            catch (Exception ex)
            {
                Logging.Log("A general error occurred in Main: " + ex);
            }
        }

       
        private static void CheckForEvent()
        {
            try
            {
                DoImport DoImport = new DoImport();
                DoImportTextFile DoImportTextFile = new DoImportTextFile();
                int iCount = 0;
                int dow = DateTime.Now.DayOfYear - 1; //DayOfYear = The day of the year, expressed as a value between 1 and 366.
                while (_shouldRun)
                {
                    bool blnDoNothing = TheSettings.Instance.DoNothing();
                    if (blnDoNothing == false)
                    {
                        if (++iCount > 10)
                        {
                            Logging.Log("Service: Active");
                            iCount = 0;
                        }
                        else Console.WriteLine("Service: Active");

                        /////////////////////////////////////////////////////////////////////////////////////////////////////
                        //1. DoImport from Excel
                        /////////////////////////////////////////////////////////////////////////////////////////////////////
                        try
                        {
                            DoImport.DoTheEvent();
                        }
                        catch (Exception ex)
                        {
                            Logging.Log("ERROR: A error occurred in DoImport: " + ex);
                        }

                        /////////////////////////////////////////////////////////////////////////////////////////////////////
                        //2. DoImportTextFile from Text file
                        /////////////////////////////////////////////////////////////////////////////////////////////////////
                        try
                        {
                            DoImportTextFile.DoTheEvent();
                        }
                        catch (Exception ex)
                        {
                            Logging.Log("ERROR: A error occurred in DoImportTextFile: " + ex);
                        }

                        /////////////////////////////////////////////////////////////////////////////////////////////////////
                        //3. Once per day event.
                        /////////////////////////////////////////////////////////////////////////////////////////////////////
                        try
                        {

                            if (dow != DateTime.Now.DayOfYear)
                            {
                                dow = DateTime.Now.DayOfYear;

                                //and start a log file with a new name.
                                Logging.ReNameLogFile();
                                Logging.RemoveOldLogFiles(TheSettings.Instance.RemoveOldLogfilesAfterDays);
                            }
                        }
                        catch (Exception exception)
                        {
                            Logging.Log(string.Format(@"ERROR: Error deleting old log file. {0}", exception));
                        }
                    }
                    else
                        Logging.Log("Service: Doing Nothing.");
                    
                    // Wait for the next time to check
                    Thread.Sleep(TheSettings.Instance.CheckEverySeconds * 1000);
                }
            }
            catch (Exception ex)
            {
                Logging.Log("ERROR: A general error occurred in CheckForEvent: " + ex);
            }
        }
    }
}
