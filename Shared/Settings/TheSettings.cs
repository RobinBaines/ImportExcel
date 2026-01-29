//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for TheSettings.cs
//Function: Read and write the settings file.
//Created Jan 2018.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
using System.IO;
using System.Reflection;
using Shared.Utilities;

namespace Shared.Settings
{
    /// <summary>
    /// Class that stores the settings of the mail application
    /// </summary>
    public class TheSettings
    {
        private static TheSettings _myInstance;
       
        private int? _removeOldLogfilesAfterDays;
        private int? _checkEverySeconds;
        private string _connectionString;

      
        private string _ExcelFolder;
        private string _ExcelImportedFolder;

        private string _DoNothingStart;
        private string _DoNothingEnd;

        /// <summary>
        /// RemoveOldLogfilesAfterDays
        /// </summary>
        public int RemoveOldLogfilesAfterDays
        {
            get
            {
                if (_removeOldLogfilesAfterDays == null)
                {
                    return 42;
                }
                return (int)_removeOldLogfilesAfterDays;
            }
            set { _removeOldLogfilesAfterDays = value; }
        }

        /// <summary>
        /// CheckEverySeconds
        /// </summary>
        public int CheckEverySeconds
        {
            get
            {
                if (_checkEverySeconds == null)
                {
                    return 10;
                }
                return (int)_checkEverySeconds;
            }
            set { _checkEverySeconds = value; }
        }


        /// <summary>
        /// The folder where the excel files which need to be imported are placed.
        /// </summary>
        public string ExcelImportFolder
        {
            get { return _ExcelFolder ?? @"c:\Projects\ExcelImport\"; }
            set { _ExcelFolder = value; }
        }

                /// <summary>
        /// The folder where the imported excel files are stored.
        /// </summary>
        public string ExcelImportedFolder
        {
            get { return _ExcelImportedFolder ?? @"c:\Projects\ExcelImport\Imported\"; }
            set { _ExcelImportedFolder = value; }
        }
        
           

        /// <summary>
        /// Connectionstring to the database
        /// </summary>
        public string ConnectionString
        {
            get { return _connectionString ?? @"Data Source=xxxxxx;Initial Catalog=yyy;Integrated Security=True"; }
            set { _connectionString = value; }
        }

        //-------------------------------------------------------------------------------------
        //return true if the service should Do Nothing.
        public bool DoNothing()
        {
            bool blnDoNothing = false;
            if (DoNothingStart != DoNothingEnd)
            {
                if (DoNothingStart != "" && DoNothingEnd != "")
                {
                    //The settings are valid times.
                    TimeSpan start = TimeSpan.Parse(DoNothingStart);
                    TimeSpan end = TimeSpan.Parse(DoNothingEnd);
                    TimeSpan now = DateTime.Now.TimeOfDay;
                    if (end > start)
                    {
                        if ((now > start) && (now < end))
                        {
                            blnDoNothing = true;
                        }
                    }
                    else
                    {
                        if ((now > start) || (now < end))
                        {
                            blnDoNothing = true;
                        }
                    }
                }
            }
            return blnDoNothing;
        }

        //return true if the string is not a valid time hh:mm.
        private bool NotATime(string strTime)
        {
            try
            {
                TimeSpan start = TimeSpan.Parse(strTime);
                return false;
            }
            catch
            {
            }
            return true;
        }

        /// <summary>
        /// Do Nothing Start
        /// </summary>
        /// 
        public string DoNothingStart
        {
            get
            {
                if (_DoNothingStart == null)
                {
                    return "07:00";
                }
                if (NotATime(_DoNothingStart))
                    return "";
                return _DoNothingStart;
            }
            set { _DoNothingStart = value; }
        }

        /// <summary>
        /// Do Nothing End
        /// </summary>
        /// 
        public string DoNothingEnd
        {
            get
            {
                if (_DoNothingEnd == null)
                {
                    return "07:00";
                }
                if (NotATime(_DoNothingEnd))
                    return "";
                return _DoNothingEnd;
            }
            set { _DoNothingEnd = value; }
        }
   
            /// <summary>
        /// Instance for reading and saving
        /// </summary>
        public static TheSettings Instance
        {
            get
            {
                if (_myInstance == null)
                {
                    _myInstance = new TheSettings();
                }
                return (_myInstance);
            }
            set { _myInstance = value; }
        }
    }

    /// <summary>
    /// Class that saves and reads the settings
    /// </summary>
    public class Settings
    {
        /// <summary>
        /// File where the settings are maintained
        /// </summary>
        public static readonly string SettingsFile = Path.Combine(Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath), "Settings", "Settings.xml");

        /// <summary>
        /// Serializes the given instance to an XML file
        /// </summary>
        /// <param name="settingsFile">filename to save the settings</param>
        /// <param name="instanceToSave">Instance to save</param>
        public static void SaveSettings(string settingsFile, object instanceToSave)
        {
            SerializeXml.ToFile(settingsFile, instanceToSave);
        }

        /// <summary>
        /// Gets the mailsettings from the XML file
        /// </summary>
        public static void GetServiceSettings()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    TheSettings.Instance = (TheSettings) SerializeXml.ReadObject(SettingsFile, typeof (TheSettings));
                    SaveSettings(SettingsFile, TheSettings.Instance);
                }
                else
                    SaveSettings(SettingsFile, TheSettings.Instance);
            }
            catch (Exception exception)
            {
                Logging.Log(string.Format("An error occurred while getting the mail settings: {0}", exception));
            }
        }
    }
}
