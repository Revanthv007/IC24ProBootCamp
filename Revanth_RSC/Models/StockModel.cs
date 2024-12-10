using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revanth_RSC.Models
{
    //ProductCode;StoreCode;ProductName;QuantityAvailable;Date;PricePerUnit
    internal class StockModel
    {
        public string ProductCode { get; set; }
        public string StoreCode { get; set; }
        public string ProductName { get; set; }
        public string QuantityAvailable { get; set; }
        public string Date { get; set; }
        public string PricePerUnit { get; set; }
    }
}
