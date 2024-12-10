using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data.SqlClient;

namespace Revanth_RSC.ProcessingFile
{
    internal class CustomerFileProcess:BaseProcessor
    {
        private string CustomerFilePath { get; set; }
        private string FailReason { get; set; }
        private bool isValid { get; set; }
        private int StoreId { get; set; }
        
        private DataSet CustomerFileDataSet {  get; set; }
        private List<CustomerModel> Customers { get; set; }
        public CustomerFileProcess(string customerFilePath,int storeid)
        {
            CustomerFilePath = customerFilePath;
            StoreId = storeid;
        }
        public void Process()
        {
            ReadFile();
            //Validate();
            if (!isValid)
            {
                MoveFile(CustomerFilePath, FileProcessStatus.Archive);
                return;
            }

            PushDataToDB();

            if (!isValid)
                MoveFile(CustomerFilePath, FileProcessStatus.Archive);
            else
                MoveFile(CustomerFilePath, FileProcessStatus.Processed);
        }

        private void PushDataToDB()
        {
            var Customers = PrepareData(CustomerFileDataSet);
            string query=string.Empty;
            foreach (var customer in Customers) 
            {
                if (customer != null) 
                {
                    query += $"IF EXISTS(select * from Customers where CustomerCode='{customer.CustomerCode}')\r\nBEGIN \r\nupdate Customers set \r\n\t\t\t\t\t [Name]='Sai Patel',\r\n\t\t\t\t\t Email='sai@gmail.com',\r\n\t\t\t\t\t ContactNumber='982346767'\r\nWHERE CustomerCode='CUST001' \r\nEND\r\n\r\nElse/*Insert*/\r\nBegin\r\ninsert into Customers (CustomerCode,Name,Email,ContactNumber)values('CUST001','Sai Patel','sai@gmail.com','9823467543')\r\nEND;";
                }
            }
            SqlConnection conn = new SqlConnection(rSCConnectionString);
            SqlCommand cmd = conn.CreateCommand();
            
        }

        private void ReadFile()
        {
            try
            {
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
        private static List<CustomerModel> PrepareData(DataSet dataSet)
        {
            DataTable dtCustomer = dataSet.Tables[0];
            DataTable dtOrder = dataSet.Tables[1];
            DataTable dtBilling = dataSet.Tables[2];

            var customers = new List<CustomerModel>();
            foreach (DataRow drCustomer in dtCustomer.Rows)
            {
                CustomerModel customerModel = new CustomerModel();
                customerModel.CustomerCode = Convert.ToString(drCustomer["CustomerCode"]);
                customerModel.CustomerName = Convert.ToString(drCustomer["Name"]);
                customerModel.CustomerEmail = Convert.ToString(drCustomer["Email"]);
                customerModel.ContactNumber = Convert.ToString(drCustomer["ContactNumber"]);
                customerModel.CustomerOders = new List<CustomerOderModel>();

                foreach (DataRow drOrder in dtOrder.Rows)
                {
                    string ordercustomerCode = Convert.ToString(drOrder["CustomerCode"]);
                    if (ordercustomerCode == customerModel.CustomerCode)
                    {
                        CustomerOderModel order = new CustomerOderModel();

                        order.OrderCode = Convert.ToString(drOrder["OrderCode"]);
                        //order.CustomerCode = Convert.ToString(dr2["CustomerCode"]);
                        order.EmployeeCode = Convert.ToString(drOrder["employeeCode"]);
                        order.StoreCode = Convert.ToString(drOrder["StoreCode"]);
                        order.ProductCode = Convert.ToString(drOrder["ProductCode"]);
                        order.OrderDate = Convert.ToDateTime(drOrder["OrderDate"]);
                        order.NofoIteams = Convert.ToInt32(drOrder["NoOfItems"]);
                        order.Amount = Convert.ToDecimal(drOrder["Amount"]);
                        order.Billings = new List<BillingModel>();


                        foreach (DataRow drBilling in dtBilling.Rows)
                        {
                            if (order.OrderCode == Convert.ToString(drBilling[0]))
                            {
                                BillingModel billingModel = new BillingModel();

                                billingModel.Date = Convert.ToDateTime(drBilling["BillingDate"]);
                                billingModel.Amount = Convert.ToDecimal(drBilling[2]);
                                order.Billings.Add(billingModel);

                            }

                        }
                        customerModel.CustomerOders.Add(order);
                    }

                }
                customers.Add(customerModel);;
            }
            return customers;

        }
    }
}




class CustomerModel
{
    public string CustomerCode { get; set; }
    public string CustomerName { get; set; }
    public string CustomerEmail { get; set; }
    public string ContactNumber { get; set; }
    public List<CustomerOderModel> CustomerOders { get; set; }
}

class CustomerOderModel
{
    public string OrderCode { get; set; }
    public string CustomerCode { get; set; }
    public string StoreCode { get; set; }
    public string EmployeeCode { get; set; }
    public string ProductCode { get; set; }

    public DateTime OrderDate { get; set; }
    public int NofoIteams { get; set; }
    public decimal Amount { get; set; }
    public List<BillingModel> Billings { get; set; }
}


class BillingModel
{
    public int BillingIdpk { get; set; }
    public int CustomerIdfk { get; set; }
    public int OrderIdfk { get; set; }
    public string OrderCode { get; set; }
    public DateTime Date { get; set; }
    public decimal Amount { get; set; }
}
