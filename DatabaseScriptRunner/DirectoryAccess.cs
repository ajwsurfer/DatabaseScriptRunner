using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseScriptRunner
{
    class DirectoryAccess
    {
        private String _directory;
        // private List<String> _files;
        public IOrderedEnumerable<string> Files;
        public String errorMessage;

        public DirectoryAccess(String directory)
        {
            errorMessage = "None";
            _directory = directory;
            try
            {
                Files = Directory.EnumerateFiles(_directory, "*.sql").OrderBy(filename => filename);
            }
            catch (DirectoryNotFoundException)
            {
                errorMessage = "DirectoryNotFoundException:\n   Please set the \"sqlScriptDirectory\" value correctly in the \"App.config\" file";
                return;
            }

        }

        public string ReadFileToString(String fileName)
        {
            try
            {
                using (StreamReader sr = new StreamReader(fileName))
                {
                    return sr.ReadToEnd();
                }
            }
            catch (Exception e)
            {
                errorMessage = e.Message.ToString();
                return null;
            }
        }

    }
}
