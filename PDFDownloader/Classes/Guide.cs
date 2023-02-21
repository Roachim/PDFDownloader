using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFDownloader.Classes
{
    public static class Guide //Hey! Listen! The guide tells you where to go.
    {
        private static string _folderLocation;
        private static string _pdfLocation;

        public static string FolderLocation 
        {
            get { return _folderLocation; } 
          
            set { _folderLocation = value; } 
        }
        public static string PDFLocation
        {
            get { return _folderLocation + @"\DownloadFolder\"; }
        }



    }
}
