using Revanth_RSC.Models;
using Revanth_RSC.ProcessingFile;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;

namespace Revanth_RSC
{
    internal class BackendJob:BaseProcessor
    {
        static void Main(string[] args)
        {
            string[] directoryPaths = Directory.GetDirectories(rootFolderPath);
            List<StoresModels> stores = GetAllStoreCodes();
            foreach (var storeDirectoryPath in directoryPaths)
            {
                string storeDirectorName = Path.GetFileName(storeDirectoryPath);
                var store = stores.Find(x => x.StoreCode == storeDirectorName);

                if (store == null)
                {
                    continue;
                }

                string[] storeFolderFiles = Directory.GetFiles(storeDirectoryPath);

                try
                {
                    string storeFilePath = GetStoreFilepath(storeDirectoryPath, FileTypes.Stores);
                    new StoreFileProcess(storeFilePath).Process();

                }
                catch (Exception ex) 
                {
                    Console.WriteLine($"Error:There is no store file {store.StoreCode} ");
                }
                try
                {
                    string stockFilePath = GetStoreFilepath(storeDirectoryPath, FileTypes.Stock);
                    new StockFileProcess(stockFilePath, store.StoreId).Process();

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error:There is no Stock file {store.StoreCode} ");
                }


                try
                {
                    string employeeFilePath = GetStoreFilepath(storeDirectoryPath, FileTypes.Employee);
                    new EmployeeFileProcess(employeeFilePath, store.StoreId).Process();

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error:There is no EmployeeFile file {store.StoreCode} ");
                }
                try
                {
                    string customerFilePath = GetStoreFilepath(storeDirectoryPath, FileTypes.Customer);
                    new CustomerFileProcess(customerFilePath, store.StoreId).Process();

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error:There is no Customer file {store.StoreCode} ");
                }


            }

        }


        private static string GetStoreFilepath(string storeDirectoryPath, FileTypes fileType)
        {
            string[] storeLevelFiles = Directory.GetFiles(storeDirectoryPath);
            string startwithvalue = string.Empty;
            switch (fileType)
            {
                case FileTypes.Stores:
                    startwithvalue = "stores_"; break;

                case FileTypes.Stock:
                    startwithvalue = "stock_"; break;

                case FileTypes.Employee:
                    startwithvalue = "employee_"; break;

                case FileTypes.Customer:
                    startwithvalue = "employee_"; break;

            }

            foreach (string file in storeLevelFiles)
            {
                if (Path.GetFileNameWithoutExtension(file).Trim().ToLower().StartsWith(startwithvalue))
                {
                    return file;
                }
            }
            return null;
        }

  

        private static List<StoresModels> GetAllStoreCodes()
        {
            List<StoresModels> stores = new List<StoresModels>();
            using (SqlConnection con = new SqlConnection(rSCConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT StoreIdPk,StoreCode,StoreName,Location,ManagerName,ContactNumber FROM Stores", con))
                {
                    con.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            StoresModels model = new StoresModels();
                            model.StoreId = Convert.ToInt32(reader["StoreIdPk"]);
                            model.StoreCode = Convert.ToString(reader["StoreCode"]);
                            model.StoreName = Convert.ToString(reader["Storename"]);
                            model.Location = Convert.ToString(reader["Location"]);
                            model.ManagerName = Convert.ToString(reader["ManagerName"]);
                            model.ContactNumber = Convert.ToString(reader["ContactNumber"]);
                            stores.Add(model);
                        }
                    }
                }
                con.Close();
            }
            return stores;
        }
    }
}
