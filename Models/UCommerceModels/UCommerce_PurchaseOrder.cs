using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.UCommerceModels
{
    public class UCommerce_PurchaseOrder
    {
        public int? OrderId { get; set; } // PK not null
        public string OrderNumber { get; set; } // nullable
        public int? CustomerId { get; set; } // FK nullable
        public int? OrderStatusId { get; set; } // FK not null
        public DateTime CreatedDate { get; set; } // not null
        public DateTime? CompletedDate { get; set; } // nullable
        public int? CurrencyId { get; set; } //FK not null
        public int? ProductCatalogGroupId { get; set; } // not null
        public int? BillingAddressId { get; set; } // FK not null
        public string Note { get; set; } // nullable
        public string BasketId { get; set; } // Not null uniqueidentifier from SQL might need to be Guid
        public decimal? VAT { get; set; } // nullable
        public decimal? OrderTotal { get; set; } // nullable
        public decimal? ShippingTotal { get; set; } // nullable
        public decimal? PaymentTotal { get; set; } // nullable
        public decimal? TaxTotal { get; set; } // nullable
        public decimal? SubTotal { get; set; } // nullable
        public string OrderGuid { get; set; } // not null uniqueidentifier from SQL might need to be Guid
        public DateTime? ModifiedOn { get; set; } // not null 
        public string CultureCode { get; set; } // nullable
        public decimal? Discount { get; set; } // nullable
        public decimal? DiscountTotal { get; set; } // nullable
    }
}