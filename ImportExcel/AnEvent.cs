//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for AnEvent.cs
//Function: Define the AnEvent interface.
//Created Jan 2018.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
namespace ImportExcel
{
    interface AnEvent
    {
         bool DoTheEvent();
         string TheCondition(ref string strSecond_String);
         bool TheAction(string strKey, string strSecond_String);
    }
}
