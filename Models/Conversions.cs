using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceGridApi.Models
{
    public class Conversions
    {
        public string SkuCompetitorNameID { get; set; }
        public string ProductSKU { get; set; }
        public decimal COMP_QTY { get; set; }
        public decimal SSW_QTY { get; set; }
        public decimal? Conversion { get; set; }
        public string Competitor_Name { get; set; }
        public string Company { get; set; }
    }
}