using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Models
{
    public class Customer
    {
        public Guid customerID { get; set; }
        public string customerCode { get; set; }

        public string customerFullName { get; set; }
        public string memberCardCode { get; set; }
        public Guid? customerGroupID { get; set; }
        public string customerGroupName { get; set; }
        public string phoneNumber { get; set; }
        public DateTime? dateOfBirth { get; set; }

        public string companyName { get; set; }
        public string taxCode { get; set; }
        public string email { get; set; }
        public string address { get; set; }
        public string note { get; set; }

        public string customerStatus { get; set; }
    }
}
