using Revanth_RSC.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revanth_RSC.ProcessingFile
{
    internal class StockFileProcess :BaseProcessor
    {
        private string stockFilePath { get; set; }
        private string FailReason { get; set; }
        private bool isValid { get; set; }
        private string[] FileContent { get; set; }
        private List<StockModel> stocks { get; set; }
        public int storeid { get; set; }
        private string DirName { get { return Path.GetFileName(Path.GetDirectoryName(stockFilePath)); } }
        public string DirPath { get { return Path.GetDirectoryName(stockFilePath);  } }
        private int numberOfRowsEffetcted { get; set; }
        public StockFileProcess(string StockFilePath,int StoreId)
        {
            stockFilePath = StockFilePath;
            storeid= StoreId;
        }

        public void Process()
        {
            ReadFile();
            Validate();
            PushDataToDB();
            Console.WriteLine(FailReason);
            MoveFile(DirPath,stockFilePath,numberOfRowsEffetcted);

        }

        

        private void ReadFile()
        {
            try
            {
                FileContent = File.ReadAllLines(stockFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to read content of filename:" + stockFilePath);
                return;
            }

        }
        private void Validate()
        {
            
            if (FileContent.Length < 1)
            {
                FailReason = "Log the error: Invalid file";
            }
        
            else if (FileContent.Length == 1)
            {
                FailReason = "Log the warning: No data present in the file";
            }
            else
            {
                for (int i = 1; i < FileContent.Length; i++)
                {
                    string[] data = FileContent[i].Split(';');
                    if (data.Length != 6)
                    {
                        FailReason = "Log the error: Invalid data; Not matching with noOffieds expected i.e., 6.";
                        break;
                    }

                    if (data[1].ToLower() != DirName.ToLower())
                    {
                        FailReason = "Log the error: Invalid data; store code not matching with current storecode.;";
                    }

                    if (!decimal.TryParse(data[3], out decimal result))
                    {
                        FailReason += $"Error: invalid quantity avaiable data; value:{data[3]};recordNumber:{i};";
                    }
                    if (!DateTime.TryParse(data[4], out DateTime dt))
                    {
                        FailReason += $"Error: invalid date; value:{data[4]};recordNumber:{i};";
                    }
                    if (!decimal.TryParse(data[5], out decimal price))
                    {
                        FailReason += $"Error: invalid price; value:{data[5]};recordNumber:{i} ;";
                    }
                    if (!string.IsNullOrEmpty(FailReason))
                        break;
                }
            }
            if (!string.IsNullOrEmpty(FailReason))
            {
                isValid = false;
                
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
            foreach (var data in stocks)
            {

                using (SqlConnection con = new SqlConnection(rSCConnectionString))
                {

                    using (SqlCommand cmd = new SqlCommand(DBConstants.INSTERPRODUCT, con))
                    {
                        con.Open();
                        cmd.Parameters.Add("@ProductCode", SqlDbType.VarChar).Value = data.ProductCode;
                        cmd.Parameters.Add("@ProductName", SqlDbType.VarChar).Value = data.ProductName;
                        cmd.Parameters.Add("@PricePerUnit", SqlDbType.Decimal).Value = data.PricePerUnit;
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                    using (SqlCommand cmd = new SqlCommand(DBConstants.INSERTSTOCK, con))
                    {
                        con.Open();
                        cmd.Parameters.Add("@Date", SqlDbType.Date).Value = data.Date;
                        cmd.Parameters.Add("@ProductCode", SqlDbType.VarChar).Value = data.ProductCode;
                        cmd.Parameters.Add("@StoreCode", SqlDbType.VarChar).Value = data.StoreCode;
                        cmd.Parameters.Add("@Quantity", SqlDbType.Decimal).Value = data.QuantityAvailable;
                        cmd.Parameters.Add("@PricePerUnit",SqlDbType.Decimal).Value=data.PricePerUnit;

                        numberOfRowsEffetcted += cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            
        }

        private void PrepareDataModel()
        {
           stocks = new List<StockModel>();
            for (int i = 1; i < FileContent.Length; i++)
            {
                try { 
                    string [] StockFileData =FileContent[i].Split(';');
                    //ProductCode;StoreCode;ProductName;QuantityAvailable;Date;PricePerUnit
                    StockModel StockData = new StockModel();
                    StockData.ProductCode = StockFileData[0];
                    StockData.StoreCode = StockFileData[1];
                    StockData.ProductName = StockFileData[2];
                    StockData.QuantityAvailable = StockFileData[3];
                    StockData.Date = StockFileData[4];
                    StockData.PricePerUnit = StockFileData[5];
                    stocks.Add(StockData);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error in Stock File in Store = " + DirName);
                }
            }

        }
    }
}

