//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for DoImportTextFile.cs
//Function: Import text files into the i_staging tabel and then call p_process_staging_table to process the data. 
//Created Jan 2018.
//------------------------------------------------
//cd \projects\applications\ImportExcel\ImportExcel\bin\Debug\
//installutil ImportExcel.exe
//installutil /u ImportExcel.exe
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
namespace ImportExcel
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }
    }
}
