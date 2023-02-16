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

namespace PDFDownloader.Classes
{
    //Class for containing the proccesses related to reading the excel Metadata and GRI, including the links
    public static class ExcelReader
    {
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
        public static void ReadExcel(string filepathAndFile)
        {
            //HTTP start client
            using var client = new HttpClient();

            client.DefaultRequestHeaders.Accept.Add(
                new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/pdf")); //limit accepted headers to pdf

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filepathAndFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //2 dictionaries for holding the links?


            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            int rows = xlRange.Rows.Count;      // Setting counters outside the loop speeds it up
            int cols = xlRange.Columns.Count;

            string HTTP = "AL";
            string HTTP2 = "AM";
            string filename = "name";

            bool worked = false;

            //create a text file; write for every download; name + donwloaded or could not be downloaded
            string PDFStatustext = Guide.PdfLocation() + @"DownloadStatus.txt";
            if (System.IO.File.Exists(PDFStatustext))
            {
                System.IO.File.Delete(PDFStatustext);
                
            }
            using StreamWriter textFileStream = System.IO.File.CreateText(PDFStatustext);

            for (int i = 2; i <= 20; i++)
            {
                HTTP = string.Empty;
                HTTP2 = string.Empty;
                filename = string.Empty;
                for (int j = 1; j <= cols; j++)
                {
                    if(j != 1 && j != 38 && j != 89)
                    {
                        continue;
                    }

                    //new line
                    //if (j == 1)
                    //    Console.Write("\r\n");

                    ////write the value to the console
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    //    Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                    //add useful things here!   
                    //important things; BRN-number, PDF-link. There are 2 links. AL and AM: Two dictionaries?
                    //alternetive: Simply download pdf here and take note immediately.

                    if(j == 1) { filename = xlRange.Cells[i, j].Value2.ToString() + ".pdf"; }
                    if(j == 38) { HTTP = xlRange.Cells[i, j].Value2.ToString(); }
                    if(j == 39) { HTTP2 = xlRange.Cells[i, j].Value2.ToString(); }



                }
                //download only if link 1 or 2 is legit
                if (!HTTP.StartsWith("http")) { HTTP = "http://" + HTTP; }
                if (!HTTP2.StartsWith("http")) { HTTP2 = "http://" + HTTP2; }
                if (!HTTP.StartsWith("http") && !HTTP2.StartsWith("http")) { continue; }

                if (HTTP != string.Empty && filename != string.Empty)
                {
                    Console.WriteLine("attmepting to download " + filename.ToString());
                    worked = DownloadPDF(client, filename, HTTP, HTTP2);
                }
                if (worked)
                {
                    textFileStream.WriteLine(filename + " = Downloaded");
                }
                else
                {
                    textFileStream.WriteLine(filename + " = could not be downloaded");
                }


            }


            //lastly Cleanup - This is important: To prevent lingering processes from holding the file access writes to the workbook
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

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

        //Method - check if AL link works
        public static bool ALLink()
        {
            return false;
        }

        //Method - Check if AM link works; Use if AL doesn't work
        public static bool MLLink()
        {
            return false;
        }

        //Method - Download PDF
        public static bool DownloadPDF(HttpClient client, string pdfName, string http, string http2 )
        {
            //string fileSpace = @"C:\Users\KOM\Desktop\Opgaver\PDF downloader\PDFDownloader\PDFDownloader\bin\Debug\net6.0\";
            try //try to download the file using the first http link
            {
                using (var s = client.GetStreamAsync(http))
                {
                    using (var fs = new FileStream(pdfName, FileMode.OpenOrCreate))
                    {
                        s.Result.CopyTo(fs);
                    }
                }
                return true;
            }
            catch (Exception ex) //if fails
            {
                if (http2 != string.Empty) //try to use the second link if it is there
                {
                    try
                    {
                        using (var s = client.GetStreamAsync(http2))
                        {
                            using (var fs = new FileStream(pdfName, FileMode.OpenOrCreate))
                            {

                                s.Result.CopyTo(fs);
                            }
                        }
                        return true;
                    }
                    catch (Exception) 
                    {
                        if (System.IO.File.Exists(Guide.PdfLocation() + pdfName))
                        {
                            System.IO.File.Delete(Guide.PdfLocation() + pdfName);

                        }
                        return false; 
                    }

                }
                else
                {
                    if (System.IO.File.Exists(Guide.PdfLocation() + pdfName))
                    {
                        System.IO.File.Delete(Guide.PdfLocation() + pdfName);

                    }
                    return false;
                }
            }
        }


        //
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