using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revanth_RSC.ProcessingFile
{
    internal class CustomerFileProcess:BaseProcessor
    {
        private string CustomerFilePath { get; set; }
        private string FailReason { get; set; }
        private bool isValid { get; set; }
        private int StoreId { get; set; }
        
        private DataSet CustomerFileDataSet {  get; set; }
        public CustomerFileProcess(string customerFilePath,int storeid)
        {
            CustomerFilePath = customerFilePath;
            StoreId = storeid;
        }
        public void Process()
        {
            ReadFile();
            Validate();
            if (!isValid)
            {
                MoveFile(employeeFilePath, FileProcessStatus.Archive);
                return;
            }

            PushDataToDB();

            if (!isValid)
                MoveFile(employeeFilePath, FileProcessStatus.Archive);
            else
                MoveFile(employeeFilePath, FileProcessStatus.Processed);
        }
        private void ReadFile()
        {
            try
            {
                var CustomerFileDataSet = new DataSet();
                string Excelconn =string.Format(ExcelConnectionString,CustomerFilePath);
                using (OleDbConnection connection = new OleDbConnection(Excelconn))
                {
                    try
                    {
                        connection.Open();
                        OleDbDataAdapter customerDataAdapter = new OleDbDataAdapter("SELECT * FROM [Customer$]", connection);
                        OleDbDataAdapter customerOrderDataAdapter = new OleDbDataAdapter("SELECT * FROM [CustomerOrder$]", connection);
                        OleDbDataAdapter customerBillingDataAdapter = new OleDbDataAdapter("SELECT * FROM [CustomerBilling$]", connection);
                        
                        customerDataAdapter.Fill(CustomerFileDataSet, "Customer");
                        customerOrderDataAdapter.Fill(CustomerFileDataSet, "CustomerOrders");
                        customerBillingDataAdapter.Fill(CustomerFileDataSet, "CustomerBilling");

                        if (connection.State == ConnectionState.Open)
                        {
                            connection.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to read content of filename:" + CustomerFilePath);
                return;
            }

        }
    }
}
