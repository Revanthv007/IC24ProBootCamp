using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Net.Http.Headers;

namespace Revanth_RSC
{
    internal class BaseProcessor
    {
        protected static string rootFolderPath { get; set; }
        protected static string rSCConnectionString { get; set; }
        protected static string ExcelConnectionString { get; set; }

        static BaseProcessor()
        {
            rootFolderPath = ConfigurationManager.AppSettings["RootFolder"];
            rSCConnectionString = ConfigurationManager.AppSettings["RSCConnectionString"].ToString();
            ExcelConnectionString = ConfigurationManager.AppSettings["ExcelConnectionString"];
        }
        protected static void MoveFile(string FilePath, FileProcessStatus fileProcessStatus)
        {
            string storeDirPath = Path.GetDirectoryName(FilePath);
            string fileName = Path.GetFileName(FilePath);
            string destinationFolderName = string.Empty;
            switch(fileProcessStatus)
            {
                case FileProcessStatus.Processed :destinationFolderName = "Processed";break;
                case FileProcessStatus.Archive: destinationFolderName = "Archive"; break;
            }
            try
            {
                File.Move(FilePath, Path.Combine(storeDirPath, destinationFolderName, fileName));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.WriteLine($"Unable to move the file sourcefile:{FilePath},destinationpath:{Path.Combine(storeDirPath, destinationFolderName, fileName)}");
            }
        }
    }
}
