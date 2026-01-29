//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for Logging.cs
//Function: 
//Created Jan 2018.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
using System.IO;
using System.Reflection;

namespace Shared.Utilities
{
    public class Logging
    {
        private static string Logfile = ReNameLogFile();
        
        private static readonly Object WriteLock = new object();

        public static void Log(string line)
        {
            try
            {
                // Lock needed because there is written from all different kind of threads
                lock (WriteLock)
                {
                    if (!Directory.Exists(Path.GetDirectoryName(Logfile)))
                        Directory.CreateDirectory(Path.GetDirectoryName(Logfile));

                    var message = string.Format(@"{0} - {1}", DateTime.Now, line);

                    // create a writer and open the file
                    TextWriter tw = new StreamWriter(Logfile, true);
                    
                    // write a line of text to the file
                    tw.WriteLine(message);
                    
                    // close the stream
                    tw.Close();
                    Console.WriteLine(message);
                }
            }
            catch (Exception exception)
            {
                // The only messagebox is shown here. If we can't write to the lofile we can't see what's happening
                Console.WriteLine(String.Format("An error occurred while trying to write to the logfile: {0}", exception));
            }
        }


        public static string ReNameLogFile()
        {
            //20130424 Create the log file name each time so that we get a new log file at the beginning of the day.
            Logfile = Path.Combine(Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath), "Logfiles",
                                                    "Logfile_" +
                                                    DateTime.Now.ToString("ddMMyyyy") + ".txt");
            return Logfile;
        }


        public static void RemoveOldLogFiles(int days)
        {
            try
            {
                var directoryInfo = new DirectoryInfo(Path.GetDirectoryName(Logfile));
                foreach (var file in directoryInfo.GetFiles("Logfile_*"))
                {
                    if (file.Name.StartsWith("Logfile_"))
                    {
                        // Remove when lastwritetime is older than 42 days
                        if (file.LastWriteTime < DateTime.Now.AddDays(-days))
                        {
                            try
                            {
                                file.Delete();
                            }
                            catch (Exception exc)
                            {
                                Log("Cannot delete logfile: " + file.FullName + ". Reason: " + exc);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Log(string.Format("An error occurred while removing the old logfiles : {0}", exception));
            }
        }
    }
}
