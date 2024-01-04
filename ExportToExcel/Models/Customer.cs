using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportToExcel.Models
{
    public class Customer
    {
        public Guid CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string CompanyName { get; set; }

        public Customer()
        {
            CustomerId = Guid.Empty;
            CustomerName = string.Empty;
            CompanyName = string.Empty;
        }
    }
}
