using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using static System.Net.WebRequestMethods;
using Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace PDFDownloader.Classes
{
    //Class for containing the proccesses related to reading the excel Metadata and GRI, including the links
    public static class ExcelReader
    {
        private static SemaphoreSlim semaphore;
        private static int padding;
        //Method - Get the Metadata Excel file
        public static string Metadata(string filepath)
        {
            string meta = @"\Metadata2006_2016.xlsx";
            string returnfile = filepath + meta;

            return returnfile;
        }

        //Method - Get the GRI excel file

        public static string GRI(string filepath) 
        {
            string meta = @"\GRI_2017_2020.xlsx";
            string returnfile = filepath + meta;

            return returnfile;
        }

        //read excel data; add data to a list so we only access it once
        public static async Task ReadExcel(string filepathAndFile)
        {
            int maxThreads = 100;
            semaphore = new SemaphoreSlim(0, maxThreads);    //Semaphore; tasks allowed at once
            padding = 0;

            //HTTP start client
            using var client = new HttpClient();

            client.DefaultRequestHeaders.Accept.Add(
                new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/pdf")); //limit accepted headers to pdf

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filepathAndFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;



            //create a text file; write for every download; name + donwloaded or could not be downloaded
            string PDFStatustext = Guide.PdfLocation() + @"DownloadStatus.txt";
            if (System.IO.File.Exists(PDFStatustext))
            {
                System.IO.File.Delete(PDFStatustext);
                
            }
            

            using StreamWriter textFileStream = System.IO.File.CreateText(PDFStatustext);

            string HTTP = string.Empty;
            string HTTP2 = string.Empty;
            string filename = string.Empty;
            int tempRow = 100;
            int rows = xlRange.Rows.Count;      // Setting counters outside the loop speeds it up
            int cols = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            List<Task> tasks = new List<Task>();

            for (int i = 2; i <= rows; i++) //25 = columns
            {
                for (int j = 1; j <= cols; j++)
                {
                    if(j != 1 && j != 38 && j != 89)
                    {
                        continue;
                    }

                    

                    //add useful things here!   
                    //important things; BRN-number, PDF-link. There are 2 links. AL and AM: Two dictionaries?
                    //alternetive: Simply download pdf here and take note immediately.

                    if(j == 1) { filename = xlRange.Cells[i, j].Value2.ToString() + ".pdf"; }
                    if(j == 38) { HTTP = xlRange.Cells[i, j].Value2.ToString(); }
                    if(j == 39) { HTTP2 = xlRange.Cells[i, j].Value2.ToString(); }



                }
                //download only if link 1 or 2 is legit
                if (!HTTP.StartsWith("http") && HTTP != string.Empty) { HTTP = "http://" + HTTP; }
                if (!HTTP2.StartsWith("http") && HTTP2 != string.Empty) { HTTP2 = "http://" + HTTP2; }
                if (!HTTP.StartsWith("http") && !HTTP2.StartsWith("http")) { continue; }

                

                if (HTTP != string.Empty && filename != string.Empty)
                {
                    Console.WriteLine("Adding new task: " + filename);
                    tasks.Add(Task.Run(() => DownloadPDF(client, filename, HTTP, HTTP2, textFileStream, semaphore)));
                    Console.WriteLine("Post task adding: " + filename);
                }

            }
            Thread.Sleep(500);

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


        //Method - Download PDF
        public static async Task DownloadPDF(HttpClient client, string pdfName, string http, string http2, StreamWriter textFileStream, SemaphoreSlim semaphore)
        {
            //Console.WriteLine("Task added: " + pdfName);
            await semaphore.WaitAsync(); //await and waitAsync the semaphore, or suffer the consequences...

            Console.WriteLine("Attmepting to download " + pdfName);
            try //try to download the file using the first http link
            {
                Interlocked.Add(ref padding, 100);
                using (var s = client.GetStreamAsync(http))
                {
                    using (var fs = new FileStream(pdfName, FileMode.OpenOrCreate))
                    {
                        await s.Result.CopyToAsync(fs);
                        //await s.CopyToAsync(fs);

                    }
                    //check if pdf or not.
                    if (!IsPdf(Guide.PdfLocation() + pdfName))
                    {
                        System.IO.File.Delete(Guide.PdfLocation() + pdfName);
                        textFileStream.WriteLine(pdfName + " = could not be downloaded");

                    }
                    else
                    {
                        textFileStream.WriteLine(pdfName + " = Downloaded");
                    }
                }
                
            }
            catch (Exception ex) //if fails
            {
                if (http2 != string.Empty) //try to use the second link if it is there
                {
                    try
                    {
                        //Console.WriteLine("PDF {0} enters the semaphore.", pdfName);
                        using (var s = await client.GetStreamAsync(http2))
                        {
                            using (var fs = new FileStream(pdfName, FileMode.OpenOrCreate))
                            {
                                //s.Result.CopyTo(fs);
                                //await s.Result.CopyToAsync(fs);
                                await s.CopyToAsync(fs);

                            }
                            //check if pdf or not
                            if (!IsPdf(Guide.PdfLocation() + pdfName))
                            {
                                System.IO.File.Delete(Guide.PdfLocation() + pdfName);
                                textFileStream.WriteLine(pdfName + " = could not be downloaded");

                            }
                            else
                            {
                                textFileStream.WriteLine(pdfName + " = Downloaded");
                            }
                        }
                        textFileStream.WriteLine(pdfName + " = Downloaded");
                    }
                    catch (Exception)
                    {
                        if (System.IO.File.Exists(Guide.PdfLocation() + pdfName))
                        {
                            System.IO.File.Delete(Guide.PdfLocation() + pdfName);
                        }
                        textFileStream.WriteLine(pdfName + " = could not be downloaded");
                    }
                }
                else
                {
                    if (System.IO.File.Exists(Guide.PdfLocation() + pdfName))
                    {
                        System.IO.File.Delete(Guide.PdfLocation() + pdfName);
                    }
                    textFileStream.WriteLine(pdfName + " = could not be downloaded");
                }
            }
            semaphore.Release();
        }


        public static bool IsPdf(string path) //determine whether or not a file is a pdf
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
                    if (amtRead == 0) return false;
                    remaining -= amtRead;
                    pos += amtRead;
                }
            }
            return pdfBytes.SequenceEqual(buf);
        }
    }
}







//-------------------------------------------don't look-------------------------------------------------
//Experimental code for downloading and trying 2 different links in a try catch

//if (HTTP != string.Empty && filename != string.Empty)
//{
//    //DownloadPDF(HTTP, filename+".pdf");
//    try //try to download the file using the first http link
//    {
//        using (var s = client.GetStreamAsync(HTTP))
//        {
//            using (var fs = new FileStream(filename, FileMode.OpenOrCreate))
//            {
//                s.Result.CopyTo(fs);
//            }
//        }
//    }
//    catch (Exception ex) //if fails
//    {
//        if (HTTP2 != string.Empty) //try to use the second link if it is there
//        {
//            try
//            {
//                using (var s = client.GetStreamAsync(HTTP2))
//                {
//                    using (var fs = new FileStream(filename, FileMode.OpenOrCreate))
//                    {
//                        s.Result.CopyTo(fs);
//                    }
//                }
//            }
//            catch (Exception) { }

//        }
//        else
//        {

//        }
//    }


//}



//-----code for seeing if client is getting a pdf format from http
//var pdfType = "application/pdf";

//if (client.ResponseHeaders["Content-Type"].Contains(someType))
//{
//    // this was a "download link"
//}



//Lines for the loop with download pdf
//new line
//if (j == 1)
//    Console.Write("\r\n");

////write the value to the console
//if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
//    Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");