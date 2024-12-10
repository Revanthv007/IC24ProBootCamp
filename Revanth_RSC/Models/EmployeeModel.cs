using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revanth_RSC.Models
{
    //EmployeeCode|StoreCode|EmployeeName|Role|DateOfJoining|DateOfLeaving|ContactNumber|Gender|Salary
    internal class EmployeeModel
    {
        public int StoreId { get; set; }
        public int EmployeeId { get; set; }
        public string EmployeeCode { get; set; }
        public string StoreCode { get; set; }
        public string EmployeeName { get; set; }
        public string Role { get; set; }
        public DateTime DateOfJoining { get { return Convert.ToDateTime(DateOfJoiningStr); } }
        public DateTime? DateOfLeaving
        {
            get
            {
                if (DateTime.TryParse(DateOfLeavingStr, out DateTime dt))
                {
                    return dt;
                }
                return null;
            }
        }
        public string DateOfJoiningStr { get; set; }
        public string DateOfLeavingStr { get; set; }
        public string ContactNumber { get; set; }
        public string Gender { get; set; }
        public string Salary { get; set; }
    }
}
