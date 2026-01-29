 ImportExcel is a service which is used to import Excel (xslx) and csv data, into an SQL Server database.

IMPORTING A CSV FILE. 
     The file is read from the ExcelImportFolder variable defined in the settings file, see below.
     The data fields: 
        ◦ Must be delimited with a ; character or a tab character 
        ◦ The filename must use the .csv extension. (Csv files created with Excel may use a comma (,) as field delimiter depending on the regional settings.) 
        ◦ The file may have up to 100 columns each with a maximum of 100 characters. 
        ◦ It may have ‘unlimited’ number of rows. 
     An SQL Database table called dbo.i_staging is dropped if it exists and is then created with fields sheet, arrow and 100 fields with type NVARCHAR(100).
     The rows from the ascii file are inserted into the table and SQL database procedure dbo.p_process_staging_table, which must exist, is called with parameter FileName.
     The file is moved from the ExcelImportFolder folder to the ExcelImportedFolder  also defined in the settings file, see below.

The p_process_staging_table must do the work of checking the filename, the data and updating relevant tables. 

IMPORTING AN EXCEL FILE.

     The file is read from the ExcelImportFolder variable defined in the settings file, see below.
     The data fields: 
        ◦ The filename should use the .xlsx or xls extension. 
        ◦ The file may have any number of sheets 
        ◦ Each sheet may have up to 100 columns each with a maximum of 100 characters. 
        ◦ A sheet may have ‘unlimited’ number of rows. 
And, as for a csv file: 
     A table called dbo.i_staging is dropped if it exists and is then created with fields sheet, arrow and 100 fields with type NVARCHAR(100).
     The rows from the first Excel sheet are inserted into the table and SQL procedure dbo.p_process_staging_table is called with parameter FileName.
     This is repeated for all the sheets in the Excel file. An excel file may contain several sheets some of which may not be visible to the user. The .i_staging.sheet field may be used in p_process_staging_table to check the sheet name.
     The file is moved from the ExcelImportFolder folder to the ExcelImportedFolder  also defined in the settings file, see below.

The OLEDB 12.0 driver (Provider=Microsoft.ACE.OLEDB.12.0) is used to access the data in the excel file. The driver 
must therefore be installed on the computer if xslx or xls files are being imported. 
This driver is available in the AccessDatabaseEngine.exe download from Microsoft. 
This download is for the 32 bit version of the driver and the ImportExcel service is compiled for x86 only. 
The Visual Studio ImportExcel project is therefore configured to create an x86 Platform Target.


RUNNING IMPORTEXCEL
The executable may be run in console mode: 
 ImportExcel -console

It may also be installed as a service using the Windows InstallUtil.exe. Delayed automatic start waiting for the SQL driver works the best.

ImportExcel reads settings from the \Settings\settings.xml file. If it does not exist this file is created the first time the utility is run. An example is shown below.

<?xml version="1.0" encoding="utf-8"?>
<TheSettings xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <RemoveOldLogfilesAfterDays>42</RemoveOldLogfilesAfterDays>
  <CheckEverySeconds>10</CheckEverySeconds>
  <ExcelImportFolder>c:\Projects\ImportExcel\Import\</ExcelImportFolder>
  <ExcelImportedFolder>c:\Projects\ImportExcel\Imported\</ExcelImportedFolder>
  <ConnectionString>Data Source=xxx;Initial Catalog=XXX;Integrated Security=True</ConnectionString>
  <DoNothingStart>01:00</DoNothingStart>
  <DoNothingEnd>04:00</DoNothingEnd>
</TheSettings>

The service will just idle without functioning in the period between DoNothingStart and DoNothingEnd This is useful if database maintenance functions need to run at night.
The meaning of the rest of the variable names should be clear.
 
The utility logs results in \LogFiles starting a new log at the start of each day. 
The log files are deleted automatically after the number of days defined in the settings file (RemoveOldLogfilesAfterDays).
The service does not delete files from the ExcelImportedFolder folder.

The service will need a login with rights to drop and create tables and to execute BulkInsert commands on the database.
Rights to write data to the log files are also required.

Correct execution can be checked either by reading the log files or by checking that a file with extension csv or xlsx is moved from the ExcelImportFolder to the ExcelImportedFolder.

An easy way to check the processing in the database is to create a version of p_process_staging_table with no processing and to view the i_staging table.