//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for DBManager.cs
//Function: interfaces with the database.
//Created Jan 2018.
//Notes: 
//Modifications: 
//20180904 Timer extended to 33.33 minutes instead of 10 mins.
//20181009 Increased timeout on sql command in public static int ProcessStagingTable(string FileName) to 6000 seconds
//after a timeout on a file with 850000 records.
//20190704 For output and return value parameters, the value is set on completion of the SqlCommand and
//after the SqlDataReader is closed.
//------------------------------------------------
using System;
using System.Data;
using System.Data.SqlClient;
using Shared.Settings;
using Shared.Utilities;
namespace ImportExcel
{
    public class DbManager
    {
        public static readonly string ConnectionString = TheSettings.Instance.ConnectionString;
        #region "Staging"

  
        public static void CreateStagingTable(int fields)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {

                try
                {
                    //
                    // Open the SqlConnection.
                    //
                    con.Open();
                    //
                    // The following code uses an SqlCommand based on the SqlConnection.
                    //
                    string strCreate = "CREATE TABLE dbo.i_staging(sheet NVARCHAR(50) NOT NULL, arow INT NOT NULL";
                    int i = 0;
                    while (i < fields)
                    {
                        strCreate = strCreate + ", Field" + i.ToString() + " NVARCHAR(100)";
                        i++;
                    }
                    strCreate = strCreate + ");";

                    try
                    {
                        //fails the first time because table is not present.
                        using (SqlCommand command = new SqlCommand("DROP TABLE dbo.i_staging", con))
                            command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        //Logging.Log("Problem in CreateStagingTable. " + ex.Message);
                    }

                    using (SqlCommand command = new SqlCommand(strCreate, con))
                        command.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    Logging.Log("Problem in CreateStagingTable. " + ex.Message);
                }
            }
        }

        public static void InsertStagingDataTableRow(CATDataSet.i_stagingDataTable dtStaging, int fields, string strValues)
        {
            string[] split;
            split = strValues.Split(new Char[] { ',', '\t' });
            DataRow newRow = dtStaging.NewRow();

            newRow["sheet"] = split[0];
            newRow["arow"] = split[1];
            int i = 0;
            while (i < fields)
            {
                newRow["Field" + i.ToString()] = split[i + 2]; 
                i++;
            }
            dtStaging.Rows.Add(newRow);
        }
        
        public static void InsertStagingRow(int fields, string strValues)
        {

            using (SqlConnection con = new SqlConnection(ConnectionString))
            {

                string sqlQuery="";
                try
                {
                    con.Open();
                    string strCreate = "INSERT INTO dbo.i_staging(sheet, arow";
                    int i = 0;
                    while (i < fields)
                    {
                        strCreate = strCreate + ", Field" + i.ToString();
                        i++;
                    }
                    strCreate = strCreate + ")";

                    sqlQuery = strCreate + strValues + ";";

                    using (SqlCommand cmd = new SqlCommand(sqlQuery, con))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Logging.Log("Problem in InsertStagingRow. " + sqlQuery + ". " + ex.Message);
                }
            }
        }

        //.
        //20181009 Increased timeout on sql command in public static int ProcessStagingTable(string FileName) to 6000 seconds
        //after a timeout on a file with 850000 records.
        //20190704 For output and return value parameters, the value is set on completion of the SqlCommand and
        //after the SqlDataReader is closed.
        public static int ProcessStagingTable(string FileName)
        {
            int iCount = 0;
            try
            {
                using (SqlConnection con = new SqlConnection(ConnectionString))
                {
                    try
                    {
                        con.Open();
                        if (FileName.Length > 240) FileName = FileName.Substring(0, 240 - 1);
                        using (SqlCommand cmd = new SqlCommand("p_process_staging_table", con))
                        {
                            //20180904 Is now 6000 seconds = 100 minutes instead of 30 seconds.
                            cmd.CommandTimeout = 6000;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@FileName", SqlDbType.NVarChar);
                            cmd.Parameters["@FileName"].Value = FileName;

                            SqlParameter theCount = new SqlParameter();
                            theCount.ParameterName = "@theCount";
                            theCount.SqlDbType = SqlDbType.Int;
                            theCount.Direction = ParameterDirection.Output;
                            cmd.Parameters.Add(theCount);

                            System.Data.SqlClient.SqlParameter retParam;
                            retParam = cmd.Parameters.Add("@ReturnValue", SqlDbType.Int);
                            retParam.Direction = ParameterDirection.ReturnValue;
                            System.Data.SqlClient.SqlDataReader reader;
                            reader = cmd.ExecuteReader();

                            //20190704 For output and return value parameters, the value is set on completion of the SqlCommand and after the SqlDataReader is closed.
                            reader.Close();
                            iCount = (int)theCount.Value;
                            reader.Dispose();
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        Logging.Log("1. Problem in ProcessStagingTable. " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Log("2. Problem in ProcessStagingTable. " + ex.Message);
            }
            return iCount;
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="table"></param>
        /// <param name="destinationTable"></param>
        public static void BulkInsert(DataTable table, string destinationTable)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                con.Open();
                var retries = 0;
                do
                {
                    try
                    {
                        {
                            var bCopy = new SqlBulkCopy(con)
                            {
                                BulkCopyTimeout = 3600,
                                DestinationTableName = destinationTable,

                            };
                            bCopy.WriteToServer(table);
                        }
                        return;
                    }
                    catch (Exception ex)
                    {
                        if (retries >= 2)
                        {
                            return;
                        }
                        retries++;
                    }
                } while (true);
            }
        }

    }  
}


