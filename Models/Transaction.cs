using PriceGridApi.Models.UCommerceModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models
{
    public class Transaction
    {
        public UCommerce_PurchaseOrder UCommerce_PurchaseOrder_Transaction { get; set; }
        public Dictionary<string, UCommerce_OrderLine> UCommerce_OrderLines_Transaction { get; set; }
        //public List<UCommerce_OrderLine> UCommerce_OrderLines_TransactionList { get; set; }
        //public static Dictionary<string, decimal> keySku_valueRunningTOtalOfSkuRefund = new Dictionary<string, decimal>();

        public Transaction()
        {
            UCommerce_PurchaseOrder UCommerce_PurchaseOrder_Transaction = new UCommerce_PurchaseOrder();
            Dictionary<string, UCommerce_OrderLine> UCommerce_OrderLines_Transaction = new Dictionary<string, UCommerce_OrderLine>();
        }
    }
}