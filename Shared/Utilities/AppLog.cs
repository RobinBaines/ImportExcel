//------------------------------------------------
//Name: Module for AppLog.cs
//Function: 
//Copyright Robin Baines 2017. All rights reserved.
//Created Nov 2017.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;
using System.Threading;
using System.ComponentModel;
using Microsoft.VisualBasic;
using System.Net;
using Shared.Settings;

namespace Shared.Utilities
{
   public class AppLog
    {
    //Get the Function/Sub name where the log text originated.
         private static string GetSource()
         {
             //StackTrace st = new StackTrace();
             //StackFrame sf = st.GetFrame(1);

             StackTrace st = new StackTrace(1, true);
                StackFrame [] stFrames = st.GetFrames();
             foreach(StackFrame sf in stFrames )
            {
                 if (sf.GetMethod().Name != "UpdateAppLog")
                     return sf.GetMethod().Name;
             }

             return "";
         }

    }
}
