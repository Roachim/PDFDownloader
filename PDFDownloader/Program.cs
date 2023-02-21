// See https://aka.ms/new-console-template for more information
using PDFDownloader.Classes;

//Console.WriteLine("Hello, World!");
Console.WriteLine("WARNING!!! Any \"DownloadFolder\", named as such, present at the location will be replaced!!");
Console.WriteLine("If a DownloadFolder is not present at the location, one will be created :)" + "\r\n");
Console.WriteLine("Please ctrl+c - ctrl+v the path of the folder containing the GRI_2017_2020.xlsx file");

string exampleFile = @"C:\Users\KOM\Desktop\Opgaver\PDF downloader\GRI_2017_2020.xlsx"; //          -----------example line

string FolderPath =
Console.ReadLine().ToString();

Guide.FolderLocation = FolderPath;

string xlFile = FolderPath + @"\GRI_2017_2020.xlsx"; //REMEMBER to catch possible mistakes made by a user


await ExcelReader.ReadExcel(xlFile);

Console.WriteLine("\r\n");
Console.WriteLine("\r\n" + "Download done!"); 
Console.WriteLine("\r\n" + "Download done!");
Console.WriteLine("\r\n" + "Download done!");
Console.WriteLine("\r\n" + "You can safely close the program!");
Console.ReadLine();
