using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFDownloader.Classes
{
    public static class Guide //Hey! Listen! The guide tells you where to go.
    {
        /// <summary>
        /// Returns the location of the file placement.
        /// If given a string; will return location + file
        /// </summary>
        /// <param name="fileName">Name of file</param>
        /// <returns></returns>
        public static string PdfLocation()
        {
            return @"C:\Users\KOM\Desktop\Opgaver\PDF downloader\PDFDownloader\PDFDownloader\bin\Debug\net6.0\";
        }
    }
}
