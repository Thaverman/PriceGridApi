using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.UCommerceModels
{
    public class UCommerce_Address
    {
        public UCommerce_Address()
        {
        }

        public string AddressId { get; set; }
        public string Line1 { get; set; }
        public string Line2 { get; set; }
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string CountryId { get; set; }
        public string Attention { get; set; }
        public string CustomerId { get; set; }
        public string CompanyName { get; set; }
        public string AddressName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmailAddress { get; set; }
        public string PhoneNumber { get; set; }
        public string MobilePhoneNumber { get; set; }
    }
}