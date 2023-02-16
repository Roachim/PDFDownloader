using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFDownloader.Classes
{
    //class for locating the folder and files using a folder path
    public static class FoldAndFileLocator
    {
        // method - insert path to excel files
        //Purpose - Clean the path so it can be read safely
        public static string CleanPath(string filepath)
        {
            if (string.IsNullOrEmpty(filepath))
            {
            }

            return filepath;
        }
    }
}
