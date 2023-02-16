using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFDownloader.Classes //DOES NOT WORK!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
{
    //class for documenting the result of which files that has been downloaded and which links did not work.
    public class NoteTaker
    {
        private string fileLocation = @"C:\Users\KOM\Desktop\Opgaver\PDF downloader\PDFDownloader\PDFDownloader\bin\Debug\net6.0\";
        private string filename = "DownloadStatus.txt";
        private string file = @"C:\Users\KOM\Desktop\Opgaver\PDF downloader\PDFDownloader\PDFDownloader\bin\Debug\net6.0\DownloadStatus.txt"; //initialized in the constructor



        private StreamWriter _writer;

        private static readonly Lazy<NoteTaker> _noteTaker
            = new Lazy<NoteTaker>(() => new NoteTaker());
        //Method - Create a new note; something.txt

        public static NoteTaker Instance
        {
            get
            {
                return _noteTaker.Value;
            }
        }
        protected NoteTaker() 
        {
            file = fileLocation + filename;

            if (System.IO.File.Exists(file))
            {
                System.IO.File.Delete(file);

            }
        }

        private StreamWriter GetWriter()
        {
            if (_writer == null)
            {
                _writer = new StreamWriter(file);
            }
            return _writer;
        }

        public void Write(string text)
        {
            var writer = GetWriter();
            Console.WriteLine(writer);
            writer.WriteLine(text);
            Console.WriteLine("trying to write " + text);
        }
        //Method - Add new entry to note

        //IDEA: make this a singleton, instantiate once when needed. Call to write when needed
    }
}
