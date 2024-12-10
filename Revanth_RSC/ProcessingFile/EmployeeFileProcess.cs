using Revanth_RSC.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;

namespace Revanth_RSC.ProcessingFile
{
    internal class EmployeeFileProcess :BaseProcessor
    {
        private string employeeFilePath { get; set; }
        private string FailReason { get; set; }
        private bool isValid { get; set; }
        private string[] FileContent { get; set; }
        private List<EmployeeModel> employeeModelList { get; set; }
        public int storeid { get; set; }
        private string DirName { get { return Path.GetFileName(Path.GetDirectoryName(employeeFilePath)); } }
        private string DirPath { get { return (Path.GetDirectoryName(employeeFilePath)); } }
        private int NumberOfRowsEffetcted { get; set; }
        public EmployeeFileProcess(string employeFile, int StoreId)
        {
            employeeFilePath = employeFile;
            storeid = StoreId;
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
                FileContent = File.ReadAllLines(employeeFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to read content of filename:" + employeeFilePath);
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
                    string[] data = FileContent[i].Split('|');
                    if (data.Length != 9)
                    {
                        FailReason = "Log the error: Invalid data; Not matching with noOffieds expected i.e., 9."+ DirName;
                        break;
                    }

                    if (data[1].ToLower() != DirName.ToLower())
                    {
                        FailReason = "Log the error: Invalid data; store code not matching with current storecode.;";
                    }

                    if (!DateTime.TryParse(data[4], out DateTime dt))
                    {
                        FailReason += $"Error: invalid date; value:{data[4]};recordNumber:{i};";
                    }
                    if (!decimal.TryParse(data[8], out decimal price))
                    {
                        FailReason += $"Error: invalid salary; value:{data[8]};recordNumber:{i} ;";
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
        private bool PushDataToDB()
        {
            if (!isValid)
            {
                return false;
            }
            PrepareEmployeeModel();
            try
            {
                foreach (var employeeModel in employeeModelList)
                {
                    using (SqlConnection con = new SqlConnection(rSCConnectionString))
                    {

                        con.Open();
                        using (SqlCommand cmd = new SqlCommand(DBConstants.INSERTORUPDATEEMP, con))
                        {

                            DateTime? dateOfLeaving = null;

                            if (DateTime.TryParse(employeeModel.DateOfLeavingStr, out DateTime convertedDate))
                            {
                                dateOfLeaving = convertedDate;
                            }
                            cmd.Parameters.Add("@EmployeeCode", SqlDbType.VarChar).Value = employeeModel.EmployeeCode;
                            cmd.Parameters.Add("@StoreCode", SqlDbType.VarChar).Value = employeeModel.StoreCode;
                            cmd.Parameters.Add("@EmployeeName", SqlDbType.VarChar).Value = employeeModel.EmployeeName;
                            cmd.Parameters.Add("@Role", SqlDbType.VarChar).Value = employeeModel.Role;
                            cmd.Parameters.Add("@DOJ", SqlDbType.Date).Value = employeeModel.DateOfJoining;
                            cmd.Parameters.Add("@DOL", SqlDbType.DateTime).Value = dateOfLeaving ?? (object)DBNull.Value;
                            cmd.Parameters.Add("@ContactNumber", SqlDbType.VarChar).Value = employeeModel.ContactNumber;
                            cmd.Parameters.Add("@Gender", SqlDbType.VarChar).Value = employeeModel.Gender;
                            cmd.Parameters.Add("@Salary", SqlDbType.Decimal).Value = employeeModel.Salary;

                            this.NumberOfRowsEffetcted += cmd.ExecuteNonQuery();
                            con.Close();
                        }
                    }
                }
                isValid = true;
            }
            catch (Exception ex)
            {
                isValid = false;
            }
            return isValid;
        }

        private void PrepareEmployeeModel()
        {
            employeeModelList = new List<EmployeeModel>();
            for (int i = 1; i < FileContent.Length; i++)
            {
                string[] employeeFileData = FileContent[i].Split('|');

                EmployeeModel EmployeeData = new EmployeeModel();
                EmployeeData.EmployeeCode = employeeFileData[0];
                EmployeeData.StoreCode = employeeFileData[1];
                EmployeeData.EmployeeName = employeeFileData[2];
                EmployeeData.Role = employeeFileData[3];
                EmployeeData.DateOfJoiningStr= employeeFileData[4];
                EmployeeData.DateOfLeavingStr = employeeFileData[5];
                EmployeeData.ContactNumber = employeeFileData[6];
                EmployeeData.Gender = employeeFileData[7];
                EmployeeData.Salary = employeeFileData[8];
                employeeModelList.Add(EmployeeData);

            }
        }




    }
}
