using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Revanth_RSC.Models;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;

namespace Revanth_RSC.ProcessingFile
{
    internal class StoreFileProcess :BaseProcessor
    {
        private string StoreFilePath { get; set; }
        private string FailReason { get; set; }
        private bool isValid {  get; set; }
        private string[] filecontentLines { get; set; }
        private string[] Data { get; set; }
        private StoresModels store { get; set; }
        private string DirName { get { return Path.GetFileName(Path.GetDirectoryName(StoreFilePath)); } }
        private string DirPath { get { return Path.GetDirectoryName(StoreFilePath); } }
        private int numberOfRowsEffetcted { get; set; }
        public StoreFileProcess(string storeFilePath)
        {
            StoreFilePath = storeFilePath;
        }

        public void Process()
        {
            ReadFile();
            Validate();
            PushDataToDB();
            Console.WriteLine( FailReason);
            MoveFile(DirPath, StoreFilePath, numberOfRowsEffetcted);
            
        }

        

        private void ReadFile()
        {
            try
            {
                filecontentLines = File.ReadAllLines(StoreFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to read content of filename:" + StoreFilePath);
                return;
            }

        }
        private void Validate()
        {
            if (filecontentLines.Length < 1)
            {
                FailReason = "Log the error: Invalid file";
            }
           
            else if (filecontentLines.Length == 1)
            {
                FailReason = "Log the warning: No data present in the file";
            }
            else if (filecontentLines.Length > 2)
            {
                FailReason = "Log the error: Invalid file; has multiple store records.";
            }
            else
            {

                Data = filecontentLines[1].Split(',');
                if (Data.Length != 6)
                {
                    FailReason = "Log the error: Invalid data; Not matching with noOffieds expected i.e., 6.";
                }
                else if (Data[1].ToLower() != DirName.ToLower())
                {
                    FailReason = "Log the error: Invalid data; store code not matching with current storecode.";
                }
            }
            if (!string.IsNullOrEmpty(FailReason))
            {
                
                return;
            }
            isValid = true;

        }
        private void PushDataToDB()
        {
            if (!isValid)
            {
                return;
            }
            PrepareDataModel();
            using (SqlConnection con = new SqlConnection(rSCConnectionString))
            {
                string query = DBConstants.STOREUPDATE;
                con.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = query;
                    cmd.Connection = con;
                    cmd.Parameters.Add("@StoreName", DbType.String).Value = store.StoreName;
                    cmd.Parameters.Add("@StoreCode", DbType.String).Value = store.StoreCode;
                    cmd.Parameters.Add("@Location", DbType.String).Value = store.Location;
                    cmd.Parameters.Add("@Manager", DbType.String).Value = store.ManagerName;
                    cmd.Parameters.Add("@ContactNumber", DbType.String).Value = store.ContactNumber;
                    numberOfRowsEffetcted = cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
        }

        private void PrepareDataModel()
        {
            store = new StoresModels();
            store.StoreCode = Data[1];
            store.StoreName = Data[2];
            store.Location = Data[3];
            store.ManagerName = Data[4];
            store.ContactNumber = Data[5];
        }
    }
}
