using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseScriptRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            RunScripts RS = new RunScripts();

            if (RS.errorMessage != "None")
            {
                ShowErrorMessage(System.AppDomain.CurrentDomain.FriendlyName, RS.errorMessage);
            }
            else
            {
                RS.CollectScripts();
                if (RS.errorMessage != "None")
                {
                    ShowErrorMessage(System.AppDomain.CurrentDomain.FriendlyName, RS.errorMessage);
                }
            }
            System.Console.WriteLine("\n Done ");
            //System.Console.WriteLine("\n Press any key to exit ...");
            //System.Console.Read();

        }
        static public void ShowErrorMessage(string executableName, string errorMessage)
        {
            System.Console.WriteLine(" Error! " + errorMessage + "\n\n" +
                                     " Usage: " + executableName + "\n" +
                                     "    This program is designed to collect a group of SQL Scripts\n     from a directory and run them.\n" +
                                     "    The program requires no arguments, but these two \"values\" in the\n    \"App.confg\" XML file must be set correctly.\n" +
                                     "      - sqlScriptDirectory: The full path of the directory\n         where the SQL Scripts are that will be run against the database\n" +
                                     "        Note: The scripts in the directory will be run in alphabetical order)\n" +
                                     "              All queries in a file, with the keyword \"results\" in its name,\n" + 
                                     "                    will return the resulting data line by line to the console\n" +
                                     "              All queries in other files will not return a result set.\n" +
                                     "      - connectionString: The connection string of the database\n         that the scripts will run on, example:\n" +
                                     "        \"Data Source=<server name>;Initial Catalog=<database name>;User ID=<user name>;Password=<password in plain text>\"\n");
        }
    }

}
