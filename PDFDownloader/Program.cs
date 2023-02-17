// See https://aka.ms/new-console-template for more information
using PDFDownloader.Classes;

//Console.WriteLine("Hello, World!");

await ExcelReader.ReadExcel(@"C:\Users\KOM\Desktop\Opgaver\PDF downloader\GRI_2017_2020.xlsx");

Console.WriteLine("\r\n" + "Download done!");
Console.ReadLine();
