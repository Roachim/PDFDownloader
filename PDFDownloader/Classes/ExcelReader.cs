using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;   //A COM reference to handle the excel file
using System.Runtime.InteropServices;
using static System.Net.WebRequestMethods;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Diagnostics;
using System.ComponentModel;

namespace PDFDownloader.Classes
{
    //Class for containing the proccesses related to reading the excel Metadata and GRI, including the links
    public static class ExcelReader
    {
        private static SemaphoreSlim semaphore;
        private static int padding;                 //Padding is used to better follow threading principles
        //Method - Get the Metadata Excel file

        //read excel data; add data to a list so we only access it once

        /// <summary>
        /// Main method, the only one to be used outside the class itself.
        /// Read excel ark, sorts links in excel into tasks, runs through the tasks.
        /// Downloads PDF's and provides a .txt with the results
        /// </summary>
        /// <param name="filepathAndFile">A string text of the path to the file, inclusding name of file</param>
        /// <returns>Void</returns>
        public static async Task ReadExcel(string filepathAndFile)
        {
            int maxThreads = 100;   //maximum number of threads allowed for the semaphore
            semaphore = new SemaphoreSlim(0, maxThreads);    //Semaphore; tasks allowed at once
            padding = 0;
            bool useSemaphores = true; //unused bool

            //HTTP start client
            using var client = new HttpClient();

            client.DefaultRequestHeaders.Accept.Add(
                new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/pdf")); //limit accepted headers to pdf

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filepathAndFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            string DownloadFolder = Guide.PDFLocation;

            if (System.IO.Directory.Exists(DownloadFolder)) //Delete and create a download folder for files to be downloaded
            {
                System.IO.Directory.Delete(DownloadFolder, true);   //True to delete everything within the folder as well
            }

            System.IO.Directory.CreateDirectory(DownloadFolder);


            string PDFStatustext = Guide.PDFLocation + @"DownloadStatus.txt";
            if (System.IO.File.Exists(PDFStatustext))
            {
                System.IO.File.Delete(PDFStatustext);
            }
            

            using StreamWriter textFileStream = System.IO.File.CreateText(PDFStatustext);   //wrtier used for the .txt file that records result

            string HTTP = string.Empty;
            string HTTP2 = string.Empty;
            string filename = string.Empty;
            int tempRow = 20;   //Used in testing

            int rows = xlRange.Rows.Count;      // Setting counters outside the loop speeds it up
            int cols = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            List<Task> tasks = new List<Task>();

            for (int i = 2; i <= tempRow; i++)     //i=2. reason: we skip the first, which is 1. excel does not start at 0.
            {
                for (int j = 1; j <= cols; j++)
                {
                    if(j != 1 && j != 38 && j != 89)    //we skip columns we don't need. Only these(1, 38 and 89) are important
                    {
                        continue;
                    }

                    if(j == 1) { filename = xlRange.Cells[i, j].Value2.ToString() + ".pdf"; }
                    if(j == 38) { HTTP = xlRange.Cells[i, j].Value2.ToString(); }
                    if(j == 39) { HTTP2 = xlRange.Cells[i, j].Value2.ToString(); }
                }
                Console.WriteLine("Adding new task: " + filename);
                tasks.Add(Task.Run(() => DownloadPDF(client, filename, HTTP, HTTP2, textFileStream, semaphore)));

            }
            Thread.Sleep(500);

            //Console.WriteLine(tasks.Count().ToString());        //21057

            semaphore.Release(maxThreads);

            await Task.WhenAll(tasks);    


            //lastly Cleanup - This is important: To prevent lingering processes from holding the file access writes to the workbook
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        /// <summary>
        /// Task to be run for every pdf. Downloades pdf via links. Writes the result to a .txt and can be limited with semaphores
        /// </summary>
        /// <param name="client">http client</param>
        /// <param name="pdfName">Name for the downloaded file</param>
        /// <param name="http">First link to try</param>
        /// <param name="http2">Second link to try</param>
        /// <param name="textFileStream">A StreamWriter to write to a .txt</param>
        /// <param name="semaphore">The loaded semaphore</param>
        /// <returns></returns>
        public static async Task DownloadPDF(HttpClient client, string pdfName, string http, string http2, StreamWriter textFileStream, SemaphoreSlim semaphore)
        {
            await semaphore.WaitAsync(); //await and waitAsync the semaphore, or suffer the consequences...
            
            //Interlocked.Add(ref padding, 100);

            bool fileDownloaded = true;

            fileDownloaded = await CheckLinkStatus(client, pdfName, http, textFileStream, fileDownloaded);

            //Second link to try
            if (!fileDownloaded)
            {
                //second http -------------------------------------------------------------------------------------------
                fileDownloaded = await CheckLinkStatus(client, pdfName, http2, textFileStream, fileDownloaded);
            }
            if (!fileDownloaded)
            {
                textFileStream.WriteLine(pdfName + " = not downloaded");
            }

            semaphore.Release(); // program MUST reach this line of code
        }

        /// <summary>
        /// Checks the status of a link, then, if ok, attempts download with 'UseDownloadLink' function.
        /// </summary>
        /// <param name="client">http client</param>
        /// <param name="pdfName">Choose a name for the pdf</param>
        /// <param name="http">The http link. The provided download link</param>
        /// <param name="textFileStream">A streamWriter to be used for the .txt file</param>
        /// <param name="fileDownloaded">current state of the file downloaded</param>
        /// <returns>False if pdf could not be downloaded; True if it succeeded.</returns>
        private static async Task<bool> CheckLinkStatus(HttpClient client, string pdfName, string http, StreamWriter textFileStream, bool fileDownloaded)
        {
            try
            {
                if (Uri.IsWellFormedUriString(http, UriKind.Absolute)) //see if URI is permissable; use other link if not
                {
                    var response = await client.GetAsync(http);

                    if (response.StatusCode == System.Net.HttpStatusCode.OK) // use other link if not ok
                    {
                        fileDownloaded = await UseDownloadLink(client, http, pdfName, textFileStream); //use other link if wasnt pdf
                    }
                    else { fileDownloaded = false; }
                }
                else { fileDownloaded = false; }
            }
            catch (Exception ex) //whatever happens with the error
            { 
                Console.WriteLine(pdfName + " = " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                fileDownloaded = false;
            }

            return fileDownloaded;
        }

        /// <summary>
        /// Async method for using a download link. Writes to a textfile if download succeeded.
        /// </summary>
        /// <param name="client">The httpClient used</param>
        /// <param name="http">Download link</param>
        /// <param name="pdfName">Name of to be downloaded</param>
        /// <param name="textFileStream">The StreamWriter used to write to the text file</param>
        /// <returns>Bool; false if it succeded; True if it didn't</returns>
        private async static Task<bool> UseDownloadLink(HttpClient client, string http, string pdfName, StreamWriter textFileStream) //revise the true and false
        {
            Console.WriteLine("Attmepting to download " + pdfName);
            using (var s = await client.GetStreamAsync(http))
            {
                
                using (var fs = new FileStream(Guide.PDFLocation + pdfName, FileMode.OpenOrCreate))
                {
                    await s.CopyToAsync(fs);
                    Console.WriteLine(pdfName +" downloaded");
                }
                //check if pdf or not
                if (!IsPdf(Guide.PDFLocation + pdfName))
                {
                    System.IO.File.Delete(Guide.PDFLocation + pdfName);
                    return false;
                }
                else
                {
                    textFileStream.WriteLine(pdfName + " = Downloaded");
                    return true;
                }
            }
        }

        /// <summary>
        /// Method for checking whether a file, already downloaded, is pdf or not.
        /// </summary>
        /// <param name="path">Location of pdf, including name of pdf</param>
        /// <returns>bool; false if input is not pdf. True if input is pdf</returns>
        public static bool IsPdf(string path) //determine whether or not a file is a pdf, does not work for 0 kb pdf files for some reason
        {
            var pdfString = "%PDF-";
            var pdfBytes = Encoding.ASCII.GetBytes(pdfString);
            var len = pdfBytes.Length;
            var buf = new byte[len];
            var remaining = len;
            var pos = 0;
            using (var f = System.IO.File.OpenRead(path))
            {
                while (remaining > 0)
                {
                    var amtRead = f.Read(buf, pos, remaining);
                    if (amtRead == 0) return false; //why not work?
                    if (amtRead < 5) return false; //also, why not work?
                    remaining -= amtRead;
                    pos += amtRead;
                }
            }
            return pdfBytes.SequenceEqual(buf);
        }
    }
}






