using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models
{
    public class Report
    {
        public Report(string sku, decimal totalRefund, decimal ourPrice, int count, decimal compPrice, string compName, decimal priceDifference, int quantity, decimal conversion, string customerPurchaseDate, string priceGridMatchDate)
        {
            this.sku = sku;
            TotalRefund = totalRefund;
            OurPrice = ourPrice;
            this.count = count;
            CompPrice = compPrice;
            CompName = compName;
            PriceDifference = priceDifference;
            Quantity = quantity;
            Conversion = conversion;
            CustomerPurchaseDate = customerPurchaseDate;
            PriceGridMatchDate = priceGridMatchDate;
        }
        public Report()
        {

        }

        public string sku { get; set; }
        public decimal TotalRefund { get; set; }
        public decimal OurPrice { get; set; }
        public int count { get; set; }
        public decimal CompPrice { get; set; }
        public string CompName { get; set; }
        public decimal PriceDifference { get; set; }
        public int Quantity { get; set; }
        public decimal Conversion { get; set; }
        public string CustomerPurchaseDate { get; set; } // for future use
        public string PriceGridMatchDate { get; set; } // for future use
    }
}