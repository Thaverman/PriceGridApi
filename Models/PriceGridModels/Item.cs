using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models.PriceGridModels
{
    public class Item
    {
        public Item()
        {

        }

        public Item(int id, string name, string url, string sku, string brandsku, int productId, decimal price)
        {
            this.id = id;
            this.name = name;
            this.url = url;
            this.sku = sku;
            this.brandsku = brandsku;
            this.productId = productId;
            this.price = price;
        }

        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public string sku { get; set; }
        public string brandsku { get; set; }
        public int productId { get; set; }
        public decimal price { get; set; }
    }
}