using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.UCommerceModels
{
    public class UCommerce_OrderLine
    {
        public UCommerce_OrderLine(UCommerce_OrderLine uCommerce_OrderLine)
        {
            OrderLineId = uCommerce_OrderLine.OrderLineId;
            OrderId = uCommerce_OrderLine.OrderId;
            Sku = uCommerce_OrderLine.Sku;
            ProductName = uCommerce_OrderLine.ProductName;
            Price = uCommerce_OrderLine.Price;
            Quantity = uCommerce_OrderLine.Quantity;
            CreatedOn = uCommerce_OrderLine.CreatedOn;
            Discount = uCommerce_OrderLine.Discount;
            VAT = uCommerce_OrderLine.VAT;
            Total = uCommerce_OrderLine.Total;
            VATRate = uCommerce_OrderLine.VATRate;
            VariantSku = uCommerce_OrderLine.VariantSku;
            ShipmentId = uCommerce_OrderLine.ShipmentId;
            UnitDiscount = uCommerce_OrderLine.UnitDiscount;
            CreatedBy = uCommerce_OrderLine.CreatedBy;
        }
        public UCommerce_OrderLine()
        {

        }

        public int OrderLineId { get; set; } //PK not null
        public int OrderId { get; set; } // FK not null
        public string Sku { get; set; } //not null
        public string ProductName { get; set; } //not null
        public decimal? Price { get; set; } //not null
        public int Quantity { get; set; } //not null
        public DateTime? CreatedOn { get; set; } //not null
        public decimal? Discount { get; set; } //not null
        public decimal? VAT { get; set; } //not null
        public decimal? Total { get; set; } //nullable
        public decimal? VATRate { get; set; } //not null
        public string VariantSku { get; set; } //nullable
        public int? ShipmentId { get; set; } //FK nullable
        public decimal? UnitDiscount { get; set; } //nullable
        public string CreatedBy { get; set; } //nullable
    }
}