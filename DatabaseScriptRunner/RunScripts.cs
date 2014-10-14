using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using dwCommons;

namespace DatabaseScriptRunner
{
    class RunScripts
    {

        private string _conStr;
        private string _sqlScriptDir;
        private DBAccess _dba;
        private DirectoryAccess _da;

        public String errorMessage;

        public RunScripts()
        {

            System.Console.WriteLine(" This program collects a directory full of SQL Scripts\n and runs them against a database.");
            _dba = new DBAccess();
            errorMessage = "None";

            _sqlScriptDir = ConfigurationManager.AppSettings["sqlScriptDirectory"];

            _conStr = ConfigurationManager.AppSettings["connectionString"];


            _da = new DirectoryAccess(_sqlScriptDir);
            if (_da.errorMessage != "None")
            {
                errorMessage = _da.errorMessage;
                return;
            }

        }

        public void CollectScripts()
        {
            System.Console.WriteLine(" Processing - connection String: " + _conStr + "\n SQL Script Directory: " + _sqlScriptDir);

            // Process the list of files found in the directory and display results to the console through a data table.
            int i = 0;
            foreach (string fileName in _da.Files)
            {
                string shortFileName = Path.GetFileName(fileName).ToString();
                Console.WriteLine("Proccessing: " + shortFileName + " ...");

                string sqlStr = _da.ReadFileToString(fileName);
                string[] stringSeparators = new string[] { "GO", "go" };

                string[] sqlStrings = sqlStr.Split(stringSeparators, StringSplitOptions.None);

                foreach (string s in sqlStrings)
                {
                    Console.WriteLine(s);
                    if (!shortFileName.Contains("results"))
                    {
                        _dba.ExecuteNoReturnQuery(_conStr, s, 600);
                        if (_dba.errorMessage != "None")
                        {
                            errorMessage = _dba.errorMessage;
                            return;
                        }
                    }
                    else
                    {
                        RunQueryResultsToConsole(s);
                        if (_dba.errorMessage != "None")
                        {
                            errorMessage = _dba.errorMessage;
                            return;
                        }
                    }

                }
                i++;
            }

        }

        private void RunQueryResultsToConsole(String sqlStr)
        {
            DataTable dt = _dba.GetDataTable(_conStr, sqlStr);
            if (_dba.errorMessage != "None")
            {
                errorMessage = _dba.errorMessage;
                return;
            }

            System.Console.WriteLine(" Results of query: " + dt);

            int n = dt.Columns.Count;
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < n; i++)
                {
                    System.Console.Write(dr[i].ToString() + ", ");
                }
                System.Console.Write("\n");
            }
        }



    }
}
