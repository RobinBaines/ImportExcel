//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for AnEventBase.cs
//Function: Base class of the AnEvent interface.
//Created Jan 2018.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
using Shared.Utilities;
namespace ImportExcel
{
    abstract class AnEventBase: AnEvent
    {
        public bool DoTheEvent()
        {
            bool ret = false;
            string strSecond_String="";
            string strRet = TheCondition(ref strSecond_String);
            if (strRet.Length > 0)
            {
                ret = TheAction(strRet, strSecond_String);
                string strMessage = FormatApplicationLog(strRet, strSecond_String);
                if (strMessage.Length > 0)
                    Logging.Log(strMessage);
            }
            return ret;
        }
        abstract public  string TheCondition(ref string strSecond_String);
        abstract public bool TheAction(string strKey, string strSecond_String);
        abstract public string FormatApplicationLog(string strKey, string strSecond_String);
        
    }
}
