using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;



namespace dwCommons
{

    /// <summary>
    /// File name: DataAccess.vb
    /// Class name: DBAccess
    /// Description: Provides a group of worker functions that allow easy access to the database.
    ///              The purpose is to reduce complexity and number of lines of code within the application.
    ///              Provides Ole connection strings
    /// </summary>
    /// <remarks> 
    /// Note: * The function "IsUserAuthorized" should be removed or modified when using this class in 
    ///         a different application then SGA-DMS.
    ///       * The "_rootDir parameter" can be changed to point to any directory the Web application has read and write access to. 
    ///       * Only the Microsoft Access 2003 files must exist before the Ole connection can be made.  No tables in the file need to exist.
    /// </remarks>
    public class DBAccess
    {

        public DBAccess()
        {
            errorMessage = "None";
        }

        public String errorMessage;
        private static string _rootDir = System.Environment.CurrentDirectory;

        private SqlConnection _sqlCon;
        private void SetConnection(string conStr)
        {
            try
            {
                _sqlCon.Dispose();

            }
            catch (Exception ex)
            {
            }
            _sqlCon = new SqlConnection(conStr);
        }

        /// <summary>
        ///   Provides easy access to a single value in the database using a sqlConnection to a MS SQL Server database.
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>The first column, first row value from the select query or "Nothing"</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public object GetSingleDBValue(string conStr, string qryStr)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;

                _sqlCon.Open();
                object retVal = sqlCmd.ExecuteScalar();

                if (retVal == System.DBNull.Value)
                {
                    retVal = null;
                }
                _sqlCon.Close();
                _sqlCon.Dispose();

                return retVal;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }

        /// <summary>
        ///   Provides easy access to a single value in the database using a sqlConnection to a MS SQL Server database.
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>The first column, first row value from the select query or "Nothing"</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public object GetSingleDBValue(ref string conStr, string qryStr, string param)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr + " @param", _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@param", param);

                _sqlCon.Open();
                object retVal = sqlCmd.ExecuteScalar();

                if (retVal == DBNull.Value)
                {
                    retVal = null;
                }
                _sqlCon.Close();
                _sqlCon.Dispose();

                return retVal;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }

        public object GetSingleDBValue(ref string conStr, string qryStr, int timeOutSecs)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;
                sqlCmd.CommandTimeout = timeOutSecs;

                _sqlCon.Open();
                object retVal = sqlCmd.ExecuteScalar();

                if (retVal == DBNull.Value)
                {
                    retVal = null;
                }
                _sqlCon.Close();
                _sqlCon.Dispose();

                return retVal;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }



        /// <summary>
        ///   Provides a SqlDataReader from a single line of code 
        ///   Note: the SqlDataReader should be closed or disposed of when you are done with it. 
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>SqlDataReader results from the select query</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public SqlDataReader GetDataReader(ref string conStr, string qryStr)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;

                _sqlCon.Open();
                return sqlCmd.ExecuteReader();
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }

        /// <summary>
        ///   Provides a SqlDataAdapter from a single line of code (same as GetDataReader(), but provides an "Adapter" instead)
        ///   Note: the SqlDataAdapter should be closed or disposed of when you are done with it.  
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>SqlDataAdapter results from the select query</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public SqlDataAdapter GetDataAdapter(ref string conStr, string qryStr)
        {
            return GetDataAdapter(ref conStr, qryStr);
        }



        public SqlDataAdapter GetDataAdapter(ref string conStr, string qryStr, string updateqryStr)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlCmd;

                if (!string.IsNullOrEmpty(updateqryStr))
                {
                    SqlCommand upsqlCmd = new SqlCommand(updateqryStr, _sqlCon);
                    upsqlCmd.CommandType = System.Data.CommandType.Text;

                    da.UpdateCommand = upsqlCmd;

                }
                return da;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }

        /// <summary>
        ///   Provides a SqlDataAdapter from a single line of code (same as GetDataReader(), but provides an "Adapter" instead)
        ///   Note: the DataAdapter should be closed or disposed of when you are done with it.  
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <param name="timeOutSecs">Select query time out seconds</param>
        /// <returns>SqlDataAdapter results from the select query</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public SqlDataAdapter GetDataAdapter(ref string conStr, string qryStr, int timeOutSecs)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr);
                sqlCmd.CommandType = System.Data.CommandType.Text;
                sqlCmd.CommandTimeout = timeOutSecs;
                sqlCmd.Connection = _sqlCon;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlCmd;
                da.SelectCommand.CommandTimeout = timeOutSecs;

                return da;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }


        /// <summary>
        ///   Provides a Data Table from a single line of code (same as GetDataReader(), but provides a "DataTable" instead)
        ///   Note: the SqlDataAdapter should be closed or disposed of when you are done with it.  
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>DataTable results from the select query</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public DataTable GetDataTable(string conStr, string qryStr)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlCmd;

                DataTable dtable = new DataTable();
                dtable.Clear();
                da.Fill(dtable);

                return dtable;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }

        }

        /// <summary>
        ///   Provides a DataTable from a single line of code (same as GetDataReader(), but provides a "DataTable" instead)
        ///   Note: the DataTable should be closed or disposed of when you are done with it.  
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <param name="timeOutSecs">Select query time out seconds</param>
        /// <returns>DataTable results from the select query</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public DataTable GetDataTable(string conStr, string qryStr, int timeOutSecs)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr);
                sqlCmd.CommandType = System.Data.CommandType.Text;
                sqlCmd.CommandTimeout = timeOutSecs;
                sqlCmd.Connection = _sqlCon;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlCmd;
                da.SelectCommand.CommandTimeout = timeOutSecs;

                DataTable dtable = new DataTable();
                dtable.Clear();
                da.Fill(dtable);

                return dtable;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }


        /// <summary>
        ///   Provides a DataSet from a single line of code (same as GetDataReader(), but provides an "DataSet" instead)
        ///   Note: the SqlDataAdapter should be closed or disposed of when you are done with it.  
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>DataSet results from the select query</returns>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public DataSet GetDataSet(string conStr, string qryStr)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
                sqlCmd.CommandType = System.Data.CommandType.Text;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlCmd;

                DataSet ds = new DataSet();
                ds.Clear();
                da.Fill(ds);

                return ds;
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }


        /// <summary>
        ///   Provides a quick Sql with no return value from a single line of code 
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public void ExecuteNoReturnQuery(string conStr, string qryStr)
        {
            SetConnection(conStr);
            SqlCommand sqlCmd = new SqlCommand(qryStr, _sqlCon);
            sqlCmd.CommandType = System.Data.CommandType.Text;
            try
            {
                _sqlCon.Open();
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return;
            }
        }

        /// <summary>
        ///   Provides a quick Sql with no return value from a single line of code 
        /// </summary>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <param name="timeOutSecs">Select query time out seconds</param>
        /// <remarks>For a sqlConnection to a MS SQL Server database (Odbc, Ole and others will not work.</remarks>
        public void ExecuteNoReturnQuery(string conStr, string qryStr, int timeOutSecs)
        {
            try
            {
                SetConnection(conStr);
                SqlCommand sqlCmd = new SqlCommand(qryStr);
                sqlCmd.CommandType = System.Data.CommandType.Text;
                sqlCmd.CommandTimeout = timeOutSecs;
                sqlCmd.Connection = _sqlCon;

                _sqlCon.Open();
                sqlCmd.ExecuteNonQuery();

                _sqlCon.Close();
                _sqlCon.Dispose();
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return;
            }
        }


        /// <summary>
        ///   Provides a OleDbDataReader from a single line of code 
        ///   Note: the OleDbDataReader should be closed or disposed of when you are done with it. 
        /// </summary>
        /// <param name="oleConStr">Connection String</param>
        /// <param name="qryStr">Select query text that will be used</param>
        /// <returns>OleDbDataReader results from the select query</returns>
        /// <remarks>For a OleDbConnection to a MS SQL Server database (Odbc, SQL and others will not work.</remarks>
        public OleDbDataReader GetOleDataReader(string oleConStr, string qryStr)
        {
            OleDbConnection oleCon = new OleDbConnection(oleConStr);
            oleCon.Open();
            OleDbCommand oleDbCmd = new OleDbCommand(qryStr, oleCon);

            return oleDbCmd.ExecuteReader();
        }

        /// <summary>
        ///  Provides an Ole connection string to a dbf file.  
        /// </summary>
        /// <param name="dir">directory on top of the root directory where the file is, or will be.</param>
        /// <returns>A fully functioning Ole connection string to a dbf file.</returns>
        /// <remarks> Note: the file name is just the table name with ".dbf" appended to it.</remarks>
        public string BuildOleDbfConnectionString(string dir)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _rootDir + "\\" + dir + "\\" + ";Extended Properties=dBASE IV;User ID=Admin;Password=;";
        }


        public string BuildOleExcelConnectionString(string dir, ref string fileName)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _rootDir + "\\" + dir + "\\" + fileName + ".xls" + ";Jet OLEDB:Engine Type=23;Extended Properties=Excel 8.0";
        }


        public string BuildOleAccessConnectionString(string dir, ref string fileName)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _rootDir + "\\" + dir + "\\" + fileName + ".mdb";
        }

        public void CopyDbfFile(string sourceFile, string destFile)
        {
            string dir = _rootDir + "\\";
            File.Delete(dir + destFile);
            if (File.Exists(dir + sourceFile) == true)
            {
                File.Copy(dir + sourceFile, dir + destFile);
            }
        }

        public string FormatIdAsString(ref Int32 inVal)
        {
            if (inVal == 0)
            {
                return "NULL";
            }
            return inVal.ToString();

        }

        public string FormatStr(ref string param)
        {
            return (param.Length == 0 ? "NULL" : "'" + param.Replace("'", "''") + "'");
        }

        public string FormatStr(ref Nullable<decimal> inVal)
        {
            if ((inVal == null))
            {
                return "NULL";
            }
            return inVal.ToString();
        }



    }


}